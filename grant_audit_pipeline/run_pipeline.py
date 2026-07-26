import sys
from pathlib import Path
import pandas as pd
import re

# Add module path
sys.path.append(str(Path(__file__).parent))

from config import (
    TRIAGE_EXCEL_PATH, ALN_CSV_PATH, OSP_SOURCE_DIR, ORACLE_PARENT_DIR,
    MD_OUTPUT_DIR, OUTPUT_EXCEL_PATH, OUTPUT_SQLITE_PATH, TEST_RUN_LIMIT
)
from parsers.text_parser import process_document_pass_1
from core.crosswalk import build_dynamic_crosswalk
from core.deduplicator import stage0_binary_preflight, clone_binary_duplicate_metadata
from exporters.db_exporter import export_to_sqlite
from exporters.excel_exporter import export_audit_workbook


def is_fuzzy_duplicate(delta_candidate: float, b_start, exec_dt, seen_actions: list) -> bool:
    """
    Tier 2 Fuzzy Window Deduplication Check.
    Requires execution dates within 14 days AND budget start dates within 30 days (if both present).
    Prevents same-amount actions in different budget years from falsely matching as duplicates when NaT occurs.
    """
    for seen in seen_actions:
        amt_match = abs(delta_candidate - seen['amount']) <= 1.0
        if not amt_match:
            continue

        seen_b = seen['budget_start']
        seen_e = seen['exec_date']

        # 1. Execution date check: Must be within 14 days
        e_diff_days = abs((pd.to_datetime(exec_dt) - pd.to_datetime(seen_e)).days) if (pd.notna(exec_dt) and pd.notna(seen_e)) else 0
        if e_diff_days > 14:
            continue

        # 2. Budget start date check: If both present, must be within 30 days
        if pd.notna(b_start) and pd.notna(seen_b):
            b_diff_days = abs((pd.to_datetime(b_start) - pd.to_datetime(seen_b)).days)
            if b_diff_days > 30:
                continue

        return True

    return False


def build_transaction_ledger(doc_list: list) -> list:
    """
    STAGE 4: SPONSOR-AGNOSTIC DELTA SOLVER & CLUSTER LEDGER ENGINE
    Evaluates explicit action deltas (Box 20) vs cumulative step-functions (Box 27 - Prior Cumulative)
    to infer non-cash budget authorizations (e.g. Year 2 carryover / re-authorizations).
    Applies Tier 2 Fuzzy Window Deduplication (±14 days execution, ±30 days budget start, ±$1.00) to isolate shadow duplicates.
    """
    clusters = {}
    for doc in doc_list:
        ckey = str(doc.get("AWARD_CLUSTER_KEY", "UNKNOWN"))
        clusters.setdefault(ckey, []).append(doc)

    transactions = []

    for ckey, docs in clusters.items():
        # Sort cluster documents chronologically
        docs_sorted = sorted(
            docs,
            key=lambda d: pd.to_datetime(
                d.get("EXECUTION_DATE") or d.get("DOC_START_DATE") or "1970-01-01",
                errors='coerce'
            )
        )

        cluster_running_cum_obligated = 0.0
        seen_binding_actions = []

        for doc in docs_sorted:
            filename = str(doc.get("Filename") or doc.get("FILENAME") or "")
            action_cat = str(doc.get("ACTION_CATEGORY", "OFFICIAL_NOA"))
            is_binding = bool(doc.get("IS_BINDING_FINANCIAL_ACTION", True))
            sponsor_tmpl = str(doc.get("SPONSOR_TEMPLATE", "GENERIC_NOA"))

            fv = doc.get("FEATURE_VECTOR") or {}
            a_action = float(fv.get("ACTION_AMOUNT") or doc.get("DELTA_OBLIGATED") or 0.0)
            c_cum_doc = float(fv.get("CUMULATIVE_AMOUNT") or 0.0)
            m_ceiling = float(fv.get("PROJECT_CEILING") or doc.get("DOC_CEILING") or 0.0)
            raw_non_binding = float(doc.get("NON_BINDING_REPORTED_BUDGET") or 0.0)

            b_start = pd.to_datetime(
                doc.get("VISION_BUDGET_START") or doc.get("DOC_START_DATE"),
                format='mixed', errors='coerce'
            )
            exec_dt = pd.to_datetime(
                doc.get("VISION_EXECUTION_DATE") or doc.get("EXECUTION_DATE"),
                format='mixed', errors='coerce'
            )

            # 1. Delta Solver Engine: Direct vs Inferred Delta Resolution (3-Point Triangulation)
            delta_candidate = 0.0
            is_inferred = False

            if is_binding:
                if a_action > 0:
                    delta_candidate = a_action
                elif a_action == 0.0 and c_cum_doc > 0.0 and c_cum_doc > cluster_running_cum_obligated:
                    delta_candidate = c_cum_doc - cluster_running_cum_obligated
                    is_inferred = True

            # 2. Tier 2 Fuzzy Window Deduplication Check
            is_dup = is_fuzzy_duplicate(delta_candidate, b_start, exec_dt, seen_binding_actions) if (is_binding and delta_candidate > 0) else False

            # 3. Status Assignment & Running Cumulative State Update
            if is_binding and delta_candidate > 0 and is_dup:
                status = "DUPLICATE_SHADOW_RECORD"
                included = False
                obligation_amt = 0.00
                stated_ceiling = 0.00
                non_binding_amt = delta_candidate
            elif is_binding and delta_candidate > 0:
                status = "PRIMARY_ACTIVE_ACTION"
                included = True
                obligation_amt = delta_candidate
                stated_ceiling = m_ceiling
                non_binding_amt = 0.00
                if is_inferred:
                    action_cat = "OFFICIAL_NOA"

                cluster_running_cum_obligated += delta_candidate
                seen_binding_actions.append({
                    'budget_start': b_start,
                    'exec_date': exec_dt,
                    'amount': delta_candidate
                })
            else:
                status = "NON_BINDING_RECORD"
                included = False
                obligation_amt = 0.00
                stated_ceiling = 0.00
                non_binding_amt = delta_candidate or raw_non_binding or float(doc.get("TABLE_TOTAL") or 0.0)

            transactions.append({
                "TRANSACTION_ID": doc.get("FILE_HASH") or doc.get("Filename"),
                "AWARD_CLUSTER_KEY": ckey,
                "ORACLE_AWARD_NUMBER": doc.get("ORACLE_AWARD_NUMBER"),
                "CAYUSE_PROJECT_NUMBER": doc.get("CAYUSE_PROJECT_NUMBER"),
                "FILENAME": filename,
                "SPONSOR_TEMPLATE": sponsor_tmpl,
                "ACTION_CATEGORY": action_cat,
                "IS_BINDING_FINANCIAL_ACTION": is_binding,
                "IS_INFERRED_ACTION": is_inferred,
                "LEDGER_ACTION_STATUS": status,
                "INCLUDED_IN_CUMULATIVE_TOTAL": included,
                "ACTION_DATE": doc.get("EXECUTION_DATE") or doc.get("VISION_EXECUTION_DATE"),
                "BUDGET_PERIOD_START": doc.get("DOC_START_DATE") or doc.get("VISION_BUDGET_START"),
                "BUDGET_PERIOD_END": doc.get("DOC_END_DATE") or doc.get("VISION_BUDGET_END"),
                "OBLIGATION_ACTION_AMOUNT": obligation_amt,
                "STATED_RECORD_TOTAL_AWARD": stated_ceiling,
                "NON_BINDING_REPORTED_BUDGET": non_binding_amt,
                "DIRECT_AMOUNT": float(doc.get("TABLE_DIRECT") or 0.0) if included else 0.0,
                "INDIRECT_AMOUNT": float(doc.get("TABLE_INDIRECT") or 0.0) if included else 0.0,
                "EXTRACTION_METHOD": doc.get("EXTRACTION_METHOD"),
                "PDF_PATH": doc.get("PDF_Path")
            })

    return transactions


def aggregate_cluster_reconciliation(cluster_transactions: list) -> dict:
    """
    Aggregates transaction-level ledger records into cluster-level truth 
    for 3-way reconciliation against Cayuse and Oracle baselines.
    """
    if not cluster_transactions:
        return {}

    df_docs = pd.DataFrame(cluster_transactions)
    cluster_key = df_docs['AWARD_CLUSTER_KEY'].iloc[0] if 'AWARD_CLUSTER_KEY' in df_docs.columns else "UNKNOWN"
    
    cayuse_proj = (
        df_docs['CAYUSE_PROJECT_NUMBER'].dropna().iloc[0] 
        if 'CAYUSE_PROJECT_NUMBER' in df_docs.columns and not df_docs['CAYUSE_PROJECT_NUMBER'].dropna().empty 
        else None
    )
    oracle_award = (
        df_docs['ORACLE_AWARD_NUMBER'].dropna().iloc[0] 
        if 'ORACLE_AWARD_NUMBER' in df_docs.columns and not df_docs['ORACLE_AWARD_NUMBER'].dropna().empty 
        else None
    )

    # Effective Dates
    df_docs['EFFECTIVE_BUDGET_START'] = pd.to_datetime(
        df_docs['BUDGET_PERIOD_START'].fillna(df_docs['ACTION_DATE']),
        format='mixed', errors='coerce'
    )
    df_docs['EFFECTIVE_PROJECT_END'] = pd.to_datetime(
        df_docs['BUDGET_PERIOD_END'],
        format='mixed', errors='coerce'
    )

    # Triage System Ledger Values
    cayuse_ceiling = float(df_docs['CAYUSE_CEILING'].dropna().max()) if 'CAYUSE_CEILING' in df_docs.columns and not df_docs['CAYUSE_CEILING'].dropna().empty else 0.0
    cayuse_obligated = float(df_docs['CAYUSE_OBLIGATED'].dropna().max()) if 'CAYUSE_OBLIGATED' in df_docs.columns and not df_docs['CAYUSE_OBLIGATED'].dropna().empty else 0.0
    oracle_obligated = float(df_docs['ORACLE_OBLIGATED'].dropna().max()) if 'ORACLE_OBLIGATED' in df_docs.columns and not df_docs['ORACLE_OBLIGATED'].dropna().empty else 0.0
    oracle_ceiling = float(df_docs['ORACLE_CEILING'].dropna().max()) if 'ORACLE_CEILING' in df_docs.columns and not df_docs['ORACLE_CEILING'].dropna().empty else 0.0

    # Filter strictly for INCLUDED_IN_CUMULATIVE_TOTAL == True
    active_actions = df_docs[df_docs['INCLUDED_IN_CUMULATIVE_TOTAL'] == True]
    pdf_obligated = float(active_actions['OBLIGATION_ACTION_AMOUNT'].sum()) if not active_actions.empty else 0.0

    if not active_actions.empty and 'EFFECTIVE_BUDGET_START' in active_actions.columns:
        df_sorted = active_actions.sort_values(by='EFFECTIVE_BUDGET_START', ascending=True)
        initial_obligated_amount = float(df_sorted['OBLIGATION_ACTION_AMOUNT'].iloc[0])
    else:
        initial_obligated_amount = 0.0

    declared_ceilings = df_docs[df_docs['STATED_RECORD_TOTAL_AWARD'] > 0]['STATED_RECORD_TOTAL_AWARD'].dropna()
    max_stated_noa_ceiling = float(declared_ceilings.max()) if not declared_ceilings.empty else 0.0

    # Active ceiling calculation considering multi-year budget expansions
    total_awarded_amount = max(cayuse_ceiling, oracle_ceiling, max_stated_noa_ceiling) if (cayuse_ceiling or oracle_ceiling or max_stated_noa_ceiling) else None

    if total_awarded_amount is not None:
        pdf_active_ceiling = max(total_awarded_amount, pdf_obligated)
    else:
        pdf_active_ceiling = pdf_obligated

    start_dates = df_docs['EFFECTIVE_BUDGET_START'].dropna()
    end_dates = df_docs['EFFECTIVE_PROJECT_END'].dropna()
    
    pdf_start_date = start_dates.min() if not start_dates.empty else None
    pdf_end_date = end_dates.max() if not end_dates.empty else None

    verdict = "IN_SYNC"
    audit_status = "IN_SYNC"
    discrepancy_reason = ""
    ceiling_breach = False
    budget_expanded = (pdf_obligated > initial_obligated_amount)

    if total_awarded_amount is not None and pdf_obligated > total_awarded_amount:
        ceiling_breach = True
        verdict = "FLAG_CEILING_BREACH"
        audit_status = "HUMAN_REVIEW_REQUIRED"
        discrepancy_reason = (
            f"ANOMALY DETECTED: Cumulative PDF obligated funds (${pdf_obligated:,.2f}) "
            f"exceed max declared project ceiling (${total_awarded_amount:,.2f}). "
            f"Manual audit intervention required."
        )
    else:
        oracle_mismatch = (oracle_obligated > 0) and abs(oracle_obligated - pdf_obligated) > 1.0
        cayuse_mismatch = (cayuse_ceiling > 0) and abs(cayuse_ceiling - pdf_active_ceiling) > 1.0

        if oracle_mismatch and cayuse_mismatch:
            verdict = "BOTH_OUT_OF_SYNC"
            audit_status = "DISCREPANCY_DETECTED"
            discrepancy_reason = f"Oracle obligated (${oracle_obligated:,.2f}) != PDF (${pdf_obligated:,.2f}) AND Cayuse ceiling (${cayuse_ceiling:,.2f}) != PDF active ceiling (${pdf_active_ceiling:,.2f})."
        elif oracle_mismatch:
            verdict = "ORACLE_OUT_OF_SYNC"
            audit_status = "DISCREPANCY_DETECTED"
            discrepancy_reason = f"Oracle obligated (${oracle_obligated:,.2f}) does not match PDF obligated (${pdf_obligated:,.2f})."
        elif cayuse_mismatch:
            verdict = "CAYUSE_OUT_OF_SYNC"
            audit_status = "DISCREPANCY_DETECTED"
            discrepancy_reason = f"Cayuse ceiling (${cayuse_ceiling:,.2f}) does not match PDF active ceiling (${pdf_active_ceiling:,.2f})."
        elif total_awarded_amount is None:
            discrepancy_reason = f"NOTICE: Total Awarded Amount (project ceiling) not explicitly declared on NOAs/Cayuse. Active ceiling set to cumulative obligated (${pdf_obligated:,.2f})."
        elif budget_expanded:
            discrepancy_reason = (
                f"INFORMATIONAL: Award budget expanded from initial obligation (${initial_obligated_amount:,.2f}) "
                f"to active cumulative total (${pdf_obligated:,.2f}). System ledgers in sync."
            )

    return {
        "AWARD_CLUSTER_KEY": cluster_key,
        "ORACLE_AWARD_NUMBER": oracle_award,
        "CAYUSE_PROJECT_NUMBER": cayuse_proj,
        "DOCUMENT_COUNT": len(cluster_transactions),
        "PDF_START_DATE_TRUTH": pdf_start_date,
        "PDF_END_DATE_TRUTH": pdf_end_date,
        "TOTAL_AWARDED_AMOUNT": total_awarded_amount,
        "INITIAL_OBLIGATED_AMOUNT": initial_obligated_amount,
        "PDF_CUMULATIVE_OBLIGATED": pdf_obligated,
        "PDF_ACTIVE_CEILING": pdf_active_ceiling,
        "CAYUSE_OBLIGATED": cayuse_obligated,
        "ORACLE_OBLIGATED": oracle_obligated,
        "CAYUSE_CEILING": cayuse_ceiling,
        "ORACLE_CEILING": oracle_ceiling,
        "BUDGET_EXPANDED": budget_expanded,
        "CEILING_BREACH": ceiling_breach,
        "AUDIT_STATUS": audit_status,
        "RECONCILIATION_VERDICT": verdict,
        "DISCREPANCY_REASON": discrepancy_reason
    }


def build_award_header(recon_record: dict) -> dict:
    """Extracts macro-level header metadata from cluster reconciliation records."""
    return {
        "AWARD_CLUSTER_KEY": recon_record.get("AWARD_CLUSTER_KEY"),
        "ORACLE_AWARD_NUMBER": recon_record.get("ORACLE_AWARD_NUMBER"),
        "CAYUSE_PROJECT_NUMBER": recon_record.get("CAYUSE_PROJECT_NUMBER"),
        "DOCUMENT_COUNT": recon_record.get("DOCUMENT_COUNT"),
        "PROJECT_START_DATE": recon_record.get("PDF_START_DATE_TRUTH"),
        "PROJECT_END_DATE": recon_record.get("PDF_END_DATE_TRUTH"),
        "TOTAL_AWARDED_AMOUNT": recon_record.get("TOTAL_AWARDED_AMOUNT"),
        "INITIAL_OBLIGATED_AMOUNT": recon_record.get("INITIAL_OBLIGATED_AMOUNT"),
        "PDF_CUMULATIVE_OBLIGATED": recon_record.get("PDF_CUMULATIVE_OBLIGATED"),
        "CAYUSE_CEILING": recon_record.get("CAYUSE_CEILING"),
        "ORACLE_CEILING": recon_record.get("ORACLE_CEILING")
    }


def main():
    print("=== STARTING SPONSOR-AGNOSTIC REFACTORED PIPELINE ===")
    
    # 1. Discover PDFs across OSP & Oracle FY Folders
    discovered_pdfs = []
    if OSP_SOURCE_DIR.exists():
        for p in OSP_SOURCE_DIR.glob("*.pdf"):
            discovered_pdfs.append((p, "OSP"))

    if ORACLE_PARENT_DIR.exists():
        for fy_dir in sorted(ORACLE_PARENT_DIR.glob("FY 20*")):
            for p in fy_dir.rglob("*.pdf"):
                discovered_pdfs.append((p, "ORACLE"))

    if TEST_RUN_LIMIT is not None:
        discovered_pdfs = discovered_pdfs[:TEST_RUN_LIMIT]
        print(f"*** TEST MODE ACTIVE: Capped processing to first {len(discovered_pdfs)} files ***")

    print(f"Discovered {len(discovered_pdfs)} total PDF files across sources...")

    # STAGE 0: Binary Pre-Flight Optimization
    unique_pdfs, binary_duplicate_map = stage0_binary_preflight(discovered_pdfs)
    print(f"Stage 0 Pre-Flight complete: {len(unique_pdfs)} unique binary files ({len(discovered_pdfs) - len(unique_pdfs)} exact duplicates bypassed).")

    # STAGE 1 & 2 & 3: Pass 1 Ingestion & Immutable Feature Vector Extraction
    extracted_unique_docs = []
    for idx, (pdf_path, source_tag) in enumerate(unique_pdfs, 1):
        doc_data = process_document_pass_1(pdf_path, source_tag, MD_OUTPUT_DIR)
        extracted_unique_docs.append(doc_data)
        if idx % 1000 == 0 or idx == len(unique_pdfs):
            print(f" -> Processed {idx}/{len(unique_pdfs)} unique documents...")

    # Expand metadata for Stage 0 binary duplicates
    all_extracted_docs = clone_binary_duplicate_metadata(extracted_unique_docs, binary_duplicate_map)

    # Dynamic Crosswalk Resolution Pass (Attaches AWARD_CLUSTER_KEY)
    build_dynamic_crosswalk(all_extracted_docs, TRIAGE_EXCEL_PATH)

    # STAGE 4: Delta Solver & Cluster Reducer Transaction Ledger
    transaction_ledger = build_transaction_ledger(all_extracted_docs)
    print(f"Stage 4 Cluster Reduction complete: {len(transaction_ledger)} transactions processed.")

    # Group transactions by cluster key for reconciliation aggregation
    cluster_tx_map = {}
    for tx in transaction_ledger:
        ckey = tx["AWARD_CLUSTER_KEY"]
        cluster_tx_map.setdefault(ckey, []).append(tx)

    reconciliation_results = [
        aggregate_cluster_reconciliation(cluster_txs) 
        for cluster_txs in cluster_tx_map.values()
    ]

    award_headers = [build_award_header(r) for r in reconciliation_results]

    # Exports to Relational SQLite and Split Excel Audit Workbook
    export_to_sqlite(award_headers, transaction_ledger, reconciliation_results, OUTPUT_SQLITE_PATH)
    export_audit_workbook(
        pd.DataFrame(reconciliation_results), 
        pd.DataFrame(award_headers), 
        pd.DataFrame(transaction_ledger), 
        OUTPUT_EXCEL_PATH
    )

    print("\nPipeline Execution Complete!")
    print(f"-> SQLite Staged Database: {OUTPUT_SQLITE_PATH}")
    print(f"-> Enhanced Audit Workbook: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()