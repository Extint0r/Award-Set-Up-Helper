import sys
from pathlib import Path
import pandas as pd

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
    and applies human-in-the-loop review flags for ceiling breaches.
    
    Refined Reconciliation Rules:
    1. Isolate Prime Financials: Exclude outgoing subcontracts (SUBCONTRACT_OUT) 
       and internal draft/revised budget worksheets from prime obligation sums.
    2. Deduplicate NoAs: Filter out duplicate copies of the same NoA sharing 
       identical execution dates and dollar deltas within a cluster.
    3. Immutable NOA Ceiling: Preserve total authorized ceiling directly from 
       official NOA headers (DOC_CEILING > 0) without synthesizing from obligations.
    4. Human Intervention Trigger: Flag any anomaly where cumulative obligated 
       dollars exceed authorized total project ceiling for manual review.
    5. Metric Benchmarking: Compare Oracle Obligated directly to PDF Cumulative 
       Obligated, and Cayuse Ceiling to PDF Active Ceiling.
    """
    if not doc_list:
        return {}

    df_docs = pd.DataFrame(doc_list)
    cluster_key = df_docs['AWARD_CLUSTER_KEY'].iloc[0] if 'AWARD_CLUSTER_KEY' in df_docs.columns else "UNKNOWN"
    
    # Extract identifiers if present in metadata
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

    df_docs = pd.DataFrame(doc_list)

    # 1. Ensure optional Vision columns exist in DataFrame schema (defaults to None if missing)
    for col in ["VISION_BUDGET_START", "VISION_PROJECT_END", "VISION_EXECUTION_DATE"]:
        if col not in df_docs.columns:
            df_docs[col] = None

    # 2. Derive Effective Dates safely
    df_docs['EFFECTIVE_BUDGET_START'] = df_docs['VISION_BUDGET_START'].fillna(df_docs['DOC_START_DATE'])
    df_docs['EFFECTIVE_PROJECT_END'] = df_docs['VISION_PROJECT_END'].fillna(df_docs['DOC_END_DATE'])
    df_docs['EFFECTIVE_EXECUTION_DATE'] = df_docs['VISION_EXECUTION_DATE'].fillna(df_docs['EXECUTION_DATE'])

    # 3. Deduplicate actions cleanly using effective budget period & dollar amount
    df_deduped_docs = df_docs.drop_duplicates(
        subset=['AWARD_CLUSTER_KEY', 'EFFECTIVE_BUDGET_START', 'DELTA_OBLIGATED'],
        keep='first'
    )
    # System ledger values from document metadata
    cayuse_ceiling = float(df_docs['CAYUSE_CEILING'].dropna().max()) if 'CAYUSE_CEILING' in df_docs.columns and not df_docs['CAYUSE_CEILING'].dropna().empty else 0.0
    cayuse_obligated = float(df_docs['CAYUSE_OBLIGATED'].dropna().max()) if 'CAYUSE_OBLIGATED' in df_docs.columns and not df_docs['CAYUSE_OBLIGATED'].dropna().empty else 0.0
    oracle_obligated = float(df_docs['ORACLE_OBLIGATED'].dropna().sum()) if 'ORACLE_OBLIGATED' in df_docs.columns and not df_docs['ORACLE_OBLIGATED'].dropna().empty else 0.0
    oracle_ceiling = float(df_docs['ORACLE_CEILING'].dropna().max()) if 'ORACLE_CEILING' in df_docs.columns and not df_docs['ORACLE_CEILING'].dropna().empty else 0.0

    # 1. Filter out non-prime (SUBCONTRACT_OUT) & non-official/internal docs for obligation math
    action_cat_mask = (
        df_docs['ACTION_CATEGORY'] != 'SUBCONTRACT_OUT' 
        if 'ACTION_CATEGORY' in df_docs.columns 
        else pd.Series(True, index=df_docs.index)
    )
    
    filename_col = (
        df_docs['Filename'].astype(str) 
        if 'Filename' in df_docs.columns 
        else df_docs['PDF_Path'].astype(str) if 'PDF_Path' in df_docs.columns else pd.Series("", index=df_docs.index)
    )
    official_mask = ~filename_col.str.contains(r'no-notice|revbud|draft|internal|work-in-progress', case=False, na=False)

    is_prime_official = action_cat_mask & official_mask
    df_financial_docs = df_docs[is_prime_official].copy()

    # 2. Deduplicate identical NoA actions (same execution date & dollar delta within cluster)
    if not df_financial_docs.empty and 'EXECUTION_DATE' in df_financial_docs.columns and 'DELTA_OBLIGATED' in df_financial_docs.columns:
        # Deduplicate identical NoA actions that cover the exact same budget period & dollar amount
        df_deduped_docs = df_financial_docs.drop_duplicates(
            subset=['AWARD_CLUSTER_KEY', 'EFFECTIVE_BUDGET_START', 'DELTA_OBLIGATED'],
            keep='first'
)
    else:
        df_deduped_docs = df_financial_docs

    # 3. Derive PDF Obligated Truth (Sum of valid incremental NoA actions)
    pdf_obligated = float(df_deduped_docs['DELTA_OBLIGATED'].sum()) if 'DELTA_OBLIGATED' in df_deduped_docs.columns else 0.0

    # 4. Derive PDF Ceiling Truth (Explicitly declared on official NoA headers)
    if 'DOC_CEILING' in df_docs.columns:
        official_noas = df_docs[(df_docs['DOC_CEILING'] > 0) & official_mask]
        pdf_ceiling = float(official_noas['DOC_CEILING'].max()) if not official_noas.empty else 0.0
    else:
        pdf_ceiling = 0.0

    # Date horizon extraction
    start_dates = pd.to_datetime(df_docs['DOC_START_DATE'], errors='coerce').dropna() if 'DOC_START_DATE' in df_docs.columns else pd.Series(dtype='datetime64[ns]')
    end_dates = pd.to_datetime(df_docs['DOC_END_DATE'], errors='coerce').dropna() if 'DOC_END_DATE' in df_docs.columns else pd.Series(dtype='datetime64[ns]')
    
    pdf_start_date = start_dates.min() if not start_dates.empty else None
    pdf_end_date = end_dates.max() if not end_dates.empty else None

    # 5. Evaluate Compliance & Human Review Triggers
    verdict = "IN_SYNC"
    audit_status = "IN_SYNC"
    discrepancy_reason = ""
    ceiling_breach = False

    # ANOMALY CHECK: Cumulative obligations exceed authorized ceiling
    if pdf_ceiling > 0 and pdf_obligated > pdf_ceiling:
        ceiling_breach = True
        verdict = "FLAG_CEILING_BREACH"
        audit_status = "HUMAN_REVIEW_REQUIRED"
        discrepancy_reason = (
            f"ANOMALY DETECTED: Cumulative PDF obligated funds (${pdf_obligated:,.2f}) "
            f"exceed authorized total project ceiling (${pdf_ceiling:,.2f}). "
            f"Manual audit intervention required."
        )
    else:
        # Ledger system reconciliation against correct benchmark metrics
        oracle_mismatch = (oracle_obligated > 0) and abs(oracle_obligated - pdf_obligated) > 1.0
        cayuse_mismatch = (cayuse_ceiling > 0) and (pdf_ceiling > 0) and abs(cayuse_ceiling - pdf_ceiling) > 1.0

        if oracle_mismatch and cayuse_mismatch:
            verdict = "BOTH_OUT_OF_SYNC"
            discrepancy_reason = f"Oracle obligated (${oracle_obligated:,.2f}) != PDF (${pdf_obligated:,.2f}) AND Cayuse ceiling (${cayuse_ceiling:,.2f}) != PDF ceiling (${pdf_ceiling:,.2f})."
        elif oracle_mismatch:
            verdict = "ORACLE_OUT_OF_SYNC"
            discrepancy_reason = f"Oracle obligated (${oracle_obligated:,.2f}) does not match PDF obligated (${pdf_obligated:,.2f})."
        elif cayuse_mismatch:
            verdict = "CAYUSE_OUT_OF_SYNC"
            discrepancy_reason = f"Cayuse ceiling (${cayuse_ceiling:,.2f}) does not match PDF active ceiling (${pdf_ceiling:,.2f})."

    return {
        "AWARD_CLUSTER_KEY": cluster_key,
        "ORACLE_AWARD_NUMBER": oracle_award,
        "CAYUSE_PROJECT_NUMBER": cayuse_proj,
        "DOCUMENT_COUNT": len(doc_list),
        "PDF_START_DATE_TRUTH": pdf_start_date,
        "PDF_END_DATE_TRUTH": pdf_end_date,
        "PDF_CUMULATIVE_OBLIGATED": pdf_obligated,
        "PDF_ACTIVE_CEILING": pdf_ceiling,
        "CAYUSE_OBLIGATED": cayuse_obligated,
        "ORACLE_OBLIGATED": oracle_obligated,
        "CAYUSE_CEILING": cayuse_ceiling,
        "ORACLE_CEILING": oracle_ceiling,
        "CEILING_BREACH": ceiling_breach,
        "AUDIT_STATUS": audit_status,
        "RECONCILIATION_VERDICT": verdict,
        "DISCREPANCY_REASON": discrepancy_reason
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

    # Apply test cap if enabled
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

    # 3. Dynamic Crosswalk Resolution Pass (Resolves UNKNOWN & Split Clusters)
    build_dynamic_crosswalk(extracted_docs)

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

    # 7. Exports to SQLite and Excel Audit Workbook
    export_to_sqlite(merged_docs, reconciliation_results, OUTPUT_SQLITE_PATH)
    export_audit_workbook(pd.DataFrame(reconciliation_results), pd.DataFrame(merged_docs), OUTPUT_EXCEL_PATH)

    print("\nPipeline Execution Complete!")
    print(f"-> SQLite Staged Database: {OUTPUT_SQLITE_PATH}")
    print(f"-> Enhanced Audit Workbook: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()