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
from core.deduplicator import deduplicate_and_merge_sources
from core.reconciler import synthesize_portfolio_pass_2
from exporters.db_exporter import export_to_sqlite
from exporters.excel_exporter import export_audit_workbook


def aggregate_cluster_reconciliation(doc_list: list) -> dict:
    """
    Aggregates document-level extractions into cluster-level truth 
    with explicit distinction between TOTAL_AWARDED_AMOUNT (max stated project ceiling)
    and INITIAL_OBLIGATED_AMOUNT (Year 1 funded budget).
    """
    if not doc_list:
        return {}

    df_docs = pd.DataFrame(doc_list)
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

    for col in ["VISION_BUDGET_START", "VISION_PROJECT_END", "VISION_EXECUTION_DATE"]:
        if col not in df_docs.columns:
            df_docs[col] = None

    if "IS_BINDING_FINANCIAL_ACTION" not in df_docs.columns:
        df_docs["IS_BINDING_FINANCIAL_ACTION"] = df_docs["ACTION_CATEGORY"].isin(["OFFICIAL_NOA", "NO_COST_EXTENSION", "PRIME_AWARD"])

    # Derive Effective Dates safely as uniform datetime64[ns]
    df_docs['EFFECTIVE_BUDGET_START'] = pd.to_datetime(
        df_docs['VISION_BUDGET_START']
        .fillna(df_docs['DOC_START_DATE'])
        .fillna(df_docs['EXECUTION_DATE']),
        format='mixed',
        errors='coerce'
    )
    
    df_docs['EFFECTIVE_PROJECT_END'] = pd.to_datetime(
        df_docs['VISION_PROJECT_END']
        .fillna(df_docs['DOC_END_DATE']),
        format='mixed',
        errors='coerce'
    )

    # System ledger values from Triage MASTER tab
    cayuse_ceiling = float(df_docs['CAYUSE_CEILING'].dropna().max()) if 'CAYUSE_CEILING' in df_docs.columns and not df_docs['CAYUSE_CEILING'].dropna().empty else 0.0
    cayuse_obligated = float(df_docs['CAYUSE_OBLIGATED'].dropna().max()) if 'CAYUSE_OBLIGATED' in df_docs.columns and not df_docs['CAYUSE_OBLIGATED'].dropna().empty else 0.0
    oracle_obligated = float(df_docs['ORACLE_OBLIGATED'].dropna().max()) if 'ORACLE_OBLIGATED' in df_docs.columns and not df_docs['ORACLE_OBLIGATED'].dropna().empty else 0.0
    oracle_ceiling = float(df_docs['ORACLE_CEILING'].dropna().max()) if 'ORACLE_CEILING' in df_docs.columns and not df_docs['ORACLE_CEILING'].dropna().empty else 0.0

    # Filter strictly for IS_BINDING_FINANCIAL_ACTION == True for obligation math
    is_prime_official = df_docs['IS_BINDING_FINANCIAL_ACTION'] == True
    df_financial_docs = df_docs[is_prime_official].copy()

    # Deduplicate identical NoA actions cleanly
    if not df_financial_docs.empty and 'DELTA_OBLIGATED' in df_financial_docs.columns:
        df_deduped_docs = df_financial_docs.drop_duplicates(
            subset=['AWARD_CLUSTER_KEY', 'EFFECTIVE_BUDGET_START', 'DELTA_OBLIGATED'],
            keep='first'
        )
    else:
        df_deduped_docs = df_financial_docs

    # Derive PDF Financial Metrics
    pdf_obligated = float(df_deduped_docs['DELTA_OBLIGATED'].sum()) if 'DELTA_OBLIGATED' in df_deduped_docs.columns else 0.0

    if not df_deduped_docs.empty and 'EFFECTIVE_BUDGET_START' in df_deduped_docs.columns:
        df_sorted_actions = df_deduped_docs.sort_values(by='EFFECTIVE_BUDGET_START', ascending=True)
        initial_obligated_amount = float(df_sorted_actions['DELTA_OBLIGATED'].iloc[0])
    else:
        initial_obligated_amount = 0.0

    declared_noa_ceilings = df_docs[(df_docs['DOC_CEILING'] > 0) & is_prime_official]['DOC_CEILING'].dropna() if 'DOC_CEILING' in df_docs.columns else pd.Series(dtype=float)
    max_stated_noa_ceiling = float(declared_noa_ceilings.max()) if not declared_noa_ceilings.empty else 0.0

    if max_stated_noa_ceiling > 0 and (cayuse_ceiling == 0 or max_stated_noa_ceiling >= cayuse_ceiling):
        total_awarded_amount = max_stated_noa_ceiling
    elif cayuse_ceiling > 0:
        total_awarded_amount = cayuse_ceiling
    else:
        total_awarded_amount = None

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
        "DOCUMENT_COUNT": len(doc_list),
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


def build_transaction_ledger(doc_list: list) -> list:
    """
    Transforms raw document extractions into a clean, relational transaction ledger 
    with explicit deduplication tracking and column segregation.
    """
    transactions = []
    seen_actions = set()

    for doc in doc_list:
        ckey = str(doc.get("AWARD_CLUSTER_KEY", "UNKNOWN"))
        filename = str(doc.get("Filename", ""))
        action_cat = str(doc.get("ACTION_CATEGORY", "OFFICIAL_NOA"))
        is_binding = bool(doc.get("IS_BINDING_FINANCIAL_ACTION", True))
        
        raw_delta = float(doc.get("DELTA_OBLIGATED") or 0.0)
        raw_ceiling = float(doc.get("DOC_CEILING") or 0.0)
        raw_non_binding = float(doc.get("NON_BINDING_REPORTED_BUDGET") or 0.0)
        
        # 1. Strict NCE Guardrail: Explicit NCE filenames cannot carry positive obligation funding
        if re.search(r'\bNCE\b|No\s*Cost\s*Extension', filename, re.IGNORECASE):
            action_cat = "NO_COST_EXTENSION"
            is_binding = True
            raw_delta = 0.0

        # 2. Date key fallback for deduplication (Budget Start -> Action Date -> Execution Date)
        eff_start = pd.to_datetime(
            doc.get("VISION_BUDGET_START") or doc.get("DOC_START_DATE") or doc.get("ACTION_DATE") or doc.get("EXECUTION_DATE") or doc.get("VISION_EXECUTION_DATE"),
            format='mixed', errors='coerce'
        )
        eff_start_str = eff_start.strftime("%Y-%m-%d") if pd.notna(eff_start) else "NO_DATE"

        # 3. Duplicate Shadow Record Check
        if is_binding and raw_delta > 0:
            action_key = (ckey, eff_start_str, round(raw_delta, 2))
            if action_key in seen_actions:
                status = "DUPLICATE_SHADOW_RECORD"
                included = False
                obligation_amt = 0.0
                stated_ceiling = 0.0
                non_binding_amt = raw_delta
            else:
                seen_actions.add(action_key)
                status = "PRIMARY_ACTIVE_ACTION"
                included = True
                obligation_amt = raw_delta
                stated_ceiling = raw_ceiling
                non_binding_amt = 0.0
        elif is_binding:
            status = "PRIMARY_ACTIVE_ACTION"
            included = True
            obligation_amt = 0.0
            stated_ceiling = raw_ceiling
            non_binding_amt = 0.0
        else:
            status = "NON_BINDING_RECORD"
            included = False
            obligation_amt = 0.0
            stated_ceiling = 0.0
            non_binding_amt = raw_delta or raw_non_binding or float(doc.get("TABLE_TOTAL") or 0.0)

        transactions.append({
            "TRANSACTION_ID": doc.get("FILE_HASH") or doc.get("Filename"),
            "AWARD_CLUSTER_KEY": doc.get("AWARD_CLUSTER_KEY"),
            "ORACLE_AWARD_NUMBER": doc.get("ORACLE_AWARD_NUMBER"),
            "CAYUSE_PROJECT_NUMBER": doc.get("CAYUSE_PROJECT_NUMBER"),
            "FILENAME": doc.get("Filename"),
            "ACTION_CATEGORY": action_cat,
            "IS_BINDING_FINANCIAL_ACTION": is_binding,
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
    print("=== STARTING MODULAR 3-WAY RECONCILIATION PIPELINE ===")
    
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
    else:
        print(f"Discovered {len(discovered_pdfs)} total PDF files across sources...")

    print(f"Discovered {len(discovered_pdfs)} PDF files.")

    # 2. Pass 1 Ingestion
    extracted_docs = []
    for idx, (pdf_path, source_tag) in enumerate(discovered_pdfs, 1):
        doc_data = process_document_pass_1(pdf_path, source_tag, MD_OUTPUT_DIR)
        extracted_docs.append(doc_data)
        if idx % 1000 == 0 or idx == len(discovered_pdfs):
            print(f" -> Processed {idx}/{len(discovered_pdfs)} documents...")

    # 3. Dynamic Crosswalk Resolution Pass
    build_dynamic_crosswalk(extracted_docs, TRIAGE_EXCEL_PATH)

    # 4. Deduplication & Merge
    merged_docs = deduplicate_and_merge_sources(extracted_docs)
    print(f"Deduplication complete: {len(merged_docs)} unique logical document actions.")

    # 5. Group by Award Cluster Key
    clusters = {}
    for doc in merged_docs:
        ckey = doc["AWARD_CLUSTER_KEY"]
        clusters.setdefault(ckey, []).append(doc)

    # 6. Refined Portfolio Synthesis & 3-Way Reconciliation
    reconciliation_results = [
        aggregate_cluster_reconciliation(doc_list) 
        for doc_list in clusters.values()
    ]

    # 7. Build Header & Transaction Relational Data Structures
    award_headers = [build_award_header(r) for r in reconciliation_results]
    transaction_ledger = build_transaction_ledger(merged_docs)

    # 8. Exports to Relational SQLite and Split Excel Audit Workbook
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