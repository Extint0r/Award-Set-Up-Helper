import os
import re
import fitz  # PyMuPDF fallback
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

# Suppress PyMuPDF warnings
fitz.TOOLS.mupdf_display_errors(False)

# ==========================================
# PATH & ENVIRONMENT CONFIGURATION
# ==========================================
TRIAGE_EXCEL_PATH = Path(r"2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx")
ALN_CSV_PATH      = Path(r"ALN.csv")

PDF_SOURCE_DIR    = Path(r"D:\0-Batch-AWARDS\processed_files")
MD_OUTPUT_DIR     = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR        = Path(r"D:\0-Batch-AWARDS\processed_files\Review")

# Output Artifacts
OUTPUT_EXCEL_PATH  = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"
OUTPUT_SQLITE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.db"

# ==========================================
# REGEX PATTERNS & MATCHING ENGINES
# ==========================================
RE_CAYUSE_PROJ = re.compile(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])')
RE_CAYUSE_PROP = re.compile(r'(?<![A-Za-z0-9])(A\d{2}-\d{4})(?![A-Za-z0-9])', re.IGNORECASE)
RE_ORACLE_NUM  = re.compile(r'\b([1-9]\d{5})\b')

# Extraction Regexes
RE_EXECUTION_DATE = re.compile(
    r'(?:Date|Execution\s*Date|Notice\s*Date|Award\s*Date)\s*[:\=]?\s*(\d{1,2}/\d{1,2}/\d{2,4})',
    re.IGNORECASE
)

RE_DATE_RANGE = re.compile(
    r'(?:Project\s*Period|Award\s*Period|Performance\s*Period|Grant\s*Period)\s*[:\=]?\s*'
    r'(\d{1,2}/\d{1,2}/\d{2,4})\s*[-–—\s+to\s+]+\s*(\d{1,2}/\d{1,2}/\d{2,4})',
    re.IGNORECASE
)

RE_CEILING_AMOUNT = re.compile(
    r'(?:Project\s*Total\s*Amount|Total\s*Award\s*Amount|Total\s*Ceiling|Award\s*Ceiling|Total\s*Project\s*Ceiling)\s*(?:\([^)]*\))?\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE
)

RE_OBLIGATED_ACTION = re.compile(
    r'(?:Current\s*Action.*?totaling|Obligated\s*Amount|Amount\s*Awarded\s*This\s*Action|Funding\s*This\s*Action|Action\s*Amount|This\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE
)

RE_BUDGET_TOTAL = re.compile(
    r'(?:Total\s*Budget|Total\s*Costs|Total\s*Amount|Total\s*Direct\s*and\s*Indirect)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE
)

RE_INDIRECT_COST = re.compile(
    r'(?:Indirect\s*Costs?|F&A\s*Costs?|Facilities\s*&\s*Admin|Overhead)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)',
    re.IGNORECASE
)

BUDGET_PAGE_KEYWORDS = [
    "direct cost", "f&a", "facilities and administrative", "f&a rate", 
    "f&a cost", "indirect cost", "total direct", "modified total direct"
]

# Master Crosswalk Lookups
LOOKUP_PROJ_TO_ORACLE = {}
LOOKUP_PROP_TO_ORACLE = {}
LOOKUP_ORACLE_TO_PROJ = {}

# ==========================================
# HELPER FUNCTIONS
# ==========================================
def parse_dollar_amount(val_str: Any) -> float:
    """Cleans currency strings into standard float values."""
    if pd.isna(val_str) or val_str is None:
        return 0.0
    try:
        clean_str = re.sub(r'[^\d.]', '', str(val_str))
        return float(clean_str) if clean_str else 0.0
    except ValueError:
        return 0.0

def parse_date(date_str: Any) -> Optional[datetime]:
    """Parses date string into standard datetime object."""
    if pd.isna(date_str) or not date_str:
        return None
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(date_str).strip(), fmt)
        except ValueError:
            pass
    return None

def find_markdown_file(pdf_path: Path, md_dir: Path) -> Optional[Path]:
    """Finds existing Markdown file regardless of prefix variations."""
    exact_path = md_dir / f"{pdf_path.stem}.md"
    if exact_path.exists():
        return exact_path
    
    matches = list(md_dir.glob(f"*{pdf_path.stem}.md"))
    if matches:
        return matches[0]
    return None

def load_system_baselines(triage_excel_path: Path) -> Dict[str, Dict[str, Any]]:
    """Loads baseline figures and populates multi-directional lookup maps."""
    global LOOKUP_PROJ_TO_ORACLE, LOOKUP_PROP_TO_ORACLE, LOOKUP_ORACLE_TO_PROJ
    
    if not triage_excel_path.exists():
        print(f"Warning: Triage baseline file '{triage_excel_path}' not found.")
        return {}
    
    try:
        xls = pd.ExcelFile(triage_excel_path)
        sheet_to_load = 'TRIAGE' if 'TRIAGE' in xls.sheet_names else ('MASTER' if 'MASTER' in xls.sheet_names else xls.sheet_names[0])
        
        df = pd.read_excel(triage_excel_path, sheet_name=sheet_to_load)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        baselines = {}
        for _, row in df.iterrows():
            oracle_num = str(row.get("ORACLE_AWARD_NUMBER", "")).split('.')[0].strip()
            cayuse_proj = str(row.get("CAYUSE_PROJECT_NUMBER", "")).strip()
            cayuse_prop = str(row.get("CAYUSE_PROPOSAL_NUMBER", "")).strip()

            if oracle_num and oracle_num not in ('nan', 'None', ''):
                if cayuse_proj and cayuse_proj not in ('nan', 'None', ''):
                    LOOKUP_PROJ_TO_ORACLE[cayuse_proj] = oracle_num
                    LOOKUP_ORACLE_TO_PROJ[oracle_num] = cayuse_proj
                if cayuse_prop and cayuse_prop not in ('nan', 'None', ''):
                    LOOKUP_PROP_TO_ORACLE[cayuse_prop] = oracle_num

            data = {
                "ORACLE_AWARD_NUMBER": oracle_num if oracle_num not in ('nan', 'None') else "",
                "CAYUSE_PROJECT_NUMBER": cayuse_proj if cayuse_proj not in ('nan', 'None') else "",
                "CAYUSE_OBLIGATED": parse_dollar_amount(row.get("CAYUSE_OBLIGATED_TOTAL", 0)),
                "ORACLE_OBLIGATED": parse_dollar_amount(row.get("ORACLE_BUDGET_BURDENED_COST", row.get("ORACLE_HARD_LIMIT", 0))),
                "CAYUSE_TOTAL_AMOUNT": parse_dollar_amount(row.get("CAYUSE_TOTAL_AMOUNT", 0)),
                "ORACLE_HEADER_HARD_LIMIT": parse_dollar_amount(row.get("ORACLE_HEADER_HARD_LIMIT", 0))
            }

            if oracle_num and oracle_num not in ('nan', 'None', ''):
                baselines[oracle_num] = data
            if cayuse_proj and cayuse_proj not in ('nan', 'None', ''):
                baselines[cayuse_proj] = data

        print(f"Loaded {len(baselines)} baseline entries from sheet '{sheet_to_load}'.")
        return baselines
    except Exception as e:
        print(f"Error loading baseline file '{triage_excel_path}': {e}")
        return {}

# ==========================================
# PASS 1: PER-DOCUMENT EXTRACTION
# ==========================================
def process_document_pass_1(pdf_path: Path, md_dir: Path) -> Dict[str, Any]:
    """
    Extracts dates/financials AND writes individual Markdown files to disk.
    """
    filename = pdf_path.name
    text = ""
    md_path = find_markdown_file(pdf_path, md_dir)

    if md_path and md_path.exists():
        with open(md_path, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()
    else:
        # Fallback: Extract directly from PDF AND save .md to disk for manual inspection
        try:
            doc = fitz.open(pdf_path)
            text = "\n".join([page.get_text() for page in doc])
            doc.close()
            
            md_dir.mkdir(parents=True, exist_ok=True)
            target_md_path = md_dir / f"{pdf_path.stem}.md"
            with open(target_md_path, "w", encoding="utf-8") as f:
                f.write(f"# Document: {filename}\n\n{text}")
            md_path = target_md_path
        except Exception:
            text = ""

    # Identifier Extraction & Resolution
    cay_proj_matches = RE_CAYUSE_PROJ.findall(filename)
    cay_prop_matches = RE_CAYUSE_PROP.findall(filename)
    oracle_matches   = RE_ORACLE_NUM.findall(filename)

    cay_proj = cay_proj_matches[0] if cay_proj_matches else ""
    cay_prop = cay_prop_matches[0] if cay_prop_matches else ""
    oracle_num = oracle_matches[0] if oracle_matches else ""

    resolved_oracle = (
        oracle_num or 
        LOOKUP_PROJ_TO_ORACLE.get(cay_proj, "") or 
        LOOKUP_PROP_TO_ORACLE.get(cay_prop, "")
    )
    
    resolved_cayuse = (
        cay_proj or 
        LOOKUP_ORACLE_TO_PROJ.get(oracle_num, "") or 
        cay_prop
    )

    cluster_key = (
        resolved_oracle or 
        resolved_cayuse or 
        "UNKNOWN"
    )

    header_scope = text[:4000] if text else ""
    
    exec_match = RE_EXECUTION_DATE.search(header_scope)
    execution_date = parse_date(exec_match.group(1)) if exec_match else None
    
    range_match = RE_DATE_RANGE.search(header_scope)
    doc_start_date = parse_date(range_match.group(1)) if range_match else None
    doc_end_date = parse_date(range_match.group(2)) if range_match else None
    
    obligated_match = RE_OBLIGATED_ACTION.search(header_scope)
    delta_obligated = parse_dollar_amount(obligated_match.group(1)) if obligated_match else 0.0
    
    ceiling_match = RE_CEILING_AMOUNT.search(header_scope)
    doc_ceiling = parse_dollar_amount(ceiling_match.group(1)) if ceiling_match else 0.0

    has_budget_table = False
    table_total = 0.0
    table_indirect = 0.0
    table_direct = 0.0
    budget_split_status = "NO_BUDGET_TABLE"

    if text and any(kw in text.lower() for kw in BUDGET_PAGE_KEYWORDS):
        has_budget_table = True
        tot_match = RE_BUDGET_TOTAL.search(text)
        ind_match = RE_INDIRECT_COST.search(text)
        
        if tot_match:
            table_total = parse_dollar_amount(tot_match.group(1))
        if ind_match:
            table_indirect = parse_dollar_amount(ind_match.group(1))
            
        if table_total > 0:
            table_direct = max(0.0, table_total - table_indirect)
            if abs(table_total - delta_obligated) < 1.0:
                budget_split_status = "MATCHED_HEADER"
            else:
                budget_split_status = "PARTIAL_OR_UNMATCHED"

    return {
        "PDF_Path": str(pdf_path),
        "Markdown_Path": str(md_path) if md_path else "",
        "Filename": filename,
        "AWARD_CLUSTER_KEY": cluster_key,
        "ORACLE_AWARD_NUMBER": resolved_oracle,
        "CAYUSE_PROJECT_NUMBER": resolved_cayuse,
        "EXECUTION_DATE": execution_date,
        "DOC_START_DATE": doc_start_date,
        "DOC_END_DATE": doc_end_date,
        "DELTA_OBLIGATED": delta_obligated,
        "DOC_CEILING": doc_ceiling,
        "HAS_BUDGET_TABLE": has_budget_table,
        "TABLE_TOTAL": table_total,
        "TABLE_DIRECT": table_direct,
        "TABLE_INDIRECT": table_indirect,
        "BUDGET_SPLIT_STATUS": budget_split_status
    }

# ==========================================
# PASS 2 & 3: SYNTHESIS & RECONCILIATION
# ==========================================
def synthesize_portfolio_pass_2(cluster_docs: List[Dict[str, Any]], system_baselines: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    """Aggregates portfolio data and generates audit verdicts with clear reasons."""
    if not cluster_docs:
        return {}

    cluster_key = cluster_docs[0]["AWARD_CLUSTER_KEY"]
    
    oracle_num = next((d["ORACLE_AWARD_NUMBER"] for d in cluster_docs if d["ORACLE_AWARD_NUMBER"]), "")
    cayuse_proj = next((d["CAYUSE_PROJECT_NUMBER"] for d in cluster_docs if d["CAYUSE_PROJECT_NUMBER"]), "")

    baseline = (
        system_baselines.get(oracle_num) or 
        system_baselines.get(cayuse_proj) or 
        system_baselines.get(cluster_key, {})
    )

    if not oracle_num:
        oracle_num = baseline.get("ORACLE_AWARD_NUMBER", "")
    if not cayuse_proj:
        cayuse_proj = baseline.get("CAYUSE_PROJECT_NUMBER", "")

    cayuse_ob = baseline.get("CAYUSE_OBLIGATED", 0.0)
    oracle_ob = baseline.get("ORACLE_OBLIGATED", 0.0)
    cayuse_ceiling = baseline.get("CAYUSE_TOTAL_AMOUNT", 0.0)
    oracle_ceiling = baseline.get("ORACLE_HEADER_HARD_LIMIT", 0.0)

    c_0 = max(cayuse_ceiling, oracle_ceiling)
    active_ceiling = c_0

    sorted_docs = sorted(
        cluster_docs, 
        key=lambda x: (x["EXECUTION_DATE"] or x["DOC_START_DATE"] or datetime.min)
    )

    start_dates = [d["DOC_START_DATE"] for d in sorted_docs if d["DOC_START_DATE"]]
    end_dates = [d["DOC_END_DATE"] for d in sorted_docs if d["DOC_END_DATE"]]
    
    award_start_date_truth = min(start_dates) if start_dates else None
    award_end_date_truth = max(end_dates) if end_dates else None
    
    cum_obligated = 0.0
    highest_doc_ceiling = 0.0
    
    for doc in sorted_docs:
        cum_obligated += doc["DELTA_OBLIGATED"]
        if doc["DOC_CEILING"] > 0:
            active_ceiling = doc["DOC_CEILING"]
            highest_doc_ceiling = max(highest_doc_ceiling, doc["DOC_CEILING"])

    if active_ceiling == 0.0:
        active_ceiling = c_0

    ceiling_breach = False
    audit_status = "IN_SYNC"
    
    if cum_obligated > active_ceiling and active_ceiling > 0:
        if highest_doc_ceiling >= cum_obligated:
            active_ceiling = highest_doc_ceiling
            audit_status = "CEILING_RESOLVED_BY_CLUSTER_DOC"
        else:
            ceiling_breach = True
            audit_status = "CEILING_BREACH_UNAPPROVED"

    # Evaluate Verdict and Discrepancy Reason
    cay_sync = (abs(cum_obligated - cayuse_ob) < 1.0)
    orc_sync = (abs(cum_obligated - oracle_ob) < 1.0)
    
    if ceiling_breach:
        reconciliation_verdict = "FLAG_CEILING_BREACH"
        discrepancy_reason = f"Cumulative PDF obligations (${cum_obligated:,.2f}) exceed active ceiling (${active_ceiling:,.2f})."
    elif cay_sync and orc_sync:
        reconciliation_verdict = "IN_SYNC"
        discrepancy_reason = "PDF Truth, Cayuse, and Oracle amounts are fully aligned."
    elif not cay_sync and orc_sync:
        reconciliation_verdict = "CAYUSE_UPDATE_REQ"
        discrepancy_reason = f"PDF Truth (${cum_obligated:,.2f}) matches Oracle, but Cayuse (${cayuse_ob:,.2f}) requires update."
    elif cay_sync and not orc_sync:
        reconciliation_verdict = "ORACLE_UPDATE_REQ"
        discrepancy_reason = f"PDF Truth (${cum_obligated:,.2f}) matches Cayuse, but Oracle (${oracle_ob:,.2f}) requires update."
    else:
        reconciliation_verdict = "BOTH_OUT_OF_SYNC"
        if cayuse_ob == 0.0 and oracle_ob == 0.0 and cum_obligated > 0:
            discrepancy_reason = f"UNMATCHED_BASELINE: Extracted PDF funds (${cum_obligated:,.2f}), but no Cayuse/Oracle baseline record was matched."
        elif cum_obligated == 0.0 and (cayuse_ob > 0 or oracle_ob > 0):
            discrepancy_reason = f"ZERO_PDF_EXTRACTION: System baseline has obligations (Cayuse: ${cayuse_ob:,.2f}, Oracle: ${oracle_ob:,.2f}), but $0 action delta extracted from PDFs."
        else:
            discrepancy_reason = f"FINANCIAL_DISCREPANCY: PDF Truth (${cum_obligated:,.2f}) differs from Cayuse (${cayuse_ob:,.2f}) and Oracle (${oracle_ob:,.2f})."

    return {
        "AWARD_CLUSTER_KEY": cluster_key,
        "ORACLE_AWARD_NUMBER": oracle_num,
        "CAYUSE_PROJECT_NUMBER": cayuse_proj,
        "DOCUMENT_COUNT": len(sorted_docs),
        "PDF_START_DATE_TRUTH": award_start_date_truth.strftime("%Y-%m-%d") if award_start_date_truth else None,
        "PDF_END_DATE_TRUTH": award_end_date_truth.strftime("%Y-%m-%d") if award_end_date_truth else None,
        "PDF_CUMULATIVE_OBLIGATED": cum_obligated,
        "PDF_ACTIVE_CEILING": active_ceiling,
        "CAYUSE_OBLIGATED": cayuse_ob,
        "ORACLE_OBLIGATED": oracle_ob,
        "CAYUSE_CEILING": cayuse_ceiling,
        "ORACLE_CEILING": oracle_ceiling,
        "CEILING_BREACH": ceiling_breach,
        "AUDIT_STATUS": audit_status,
        "RECONCILIATION_VERDICT": reconciliation_verdict,
        "DISCREPANCY_REASON": discrepancy_reason
    }

# ==========================================
# DATABASE INITIALIZATION
# ==========================================
def initialize_database(db_path: Path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tbl_Documents (
        PDF_Path TEXT PRIMARY KEY,
        Markdown_Path TEXT,
        Filename TEXT,
        AWARD_CLUSTER_KEY TEXT,
        ORACLE_AWARD_NUMBER TEXT,
        CAYUSE_PROJECT_NUMBER TEXT,
        EXECUTION_DATE TEXT,
        DOC_START_DATE TEXT,
        DOC_END_DATE TEXT,
        DELTA_OBLIGATED REAL,
        DOC_CEILING REAL,
        HAS_BUDGET_TABLE INTEGER,
        TABLE_TOTAL REAL,
        TABLE_DIRECT REAL,
        TABLE_INDIRECT REAL,
        BUDGET_SPLIT_STATUS TEXT
    )
    """)
    
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tbl_Reconciliation_Summary (
        AWARD_CLUSTER_KEY TEXT PRIMARY KEY,
        ORACLE_AWARD_NUMBER TEXT,
        CAYUSE_PROJECT_NUMBER TEXT,
        DOCUMENT_COUNT INTEGER,
        PDF_START_DATE_TRUTH TEXT,
        PDF_END_DATE_TRUTH TEXT,
        PDF_CUMULATIVE_OBLIGATED REAL,
        PDF_ACTIVE_CEILING REAL,
        CAYUSE_OBLIGATED REAL,
        ORACLE_OBLIGATED REAL,
        CAYUSE_CEILING REAL,
        ORACLE_CEILING REAL,
        CEILING_BREACH INTEGER,
        AUDIT_STATUS TEXT,
        RECONCILIATION_VERDICT TEXT,
        DISCREPANCY_REASON TEXT,
        REVIEW_STATUS TEXT DEFAULT 'PENDING',
        REVIEWER_NOTES TEXT
    )
    """)
    
    conn.commit()
    conn.close()

# ==========================================
# MAIN EXECUTION PIPELINE
# ==========================================
def main():
    print("=== STARTING PRODUCTION 3-WAY RECONCILIATION PIPELINE ===")
    initialize_database(OUTPUT_SQLITE_PATH)
    
    # 1. Load Real System Baselines from Excel
    system_baselines = load_system_baselines(TRIAGE_EXCEL_PATH)
    
    # 2. Discover and Process Actual Document Corpus
    if not PDF_SOURCE_DIR.exists():
        print(f"Error: PDF source directory '{PDF_SOURCE_DIR}' does not exist.")
        return

    pdf_files = list(PDF_SOURCE_DIR.glob("*.pdf"))
    print(f"Found {len(pdf_files)} PDF documents in '{PDF_SOURCE_DIR}'...")
    
    extracted_documents = []
    for idx, pdf_path in enumerate(pdf_files, 1):
        doc_data = process_document_pass_1(pdf_path, MD_OUTPUT_DIR)
        extracted_documents.append(doc_data)
        
        if idx % 1000 == 0 or idx == len(pdf_files):
            print(f" -> Processed {idx}/{len(pdf_files)} documents...")

    if not extracted_documents:
        print("No documents found or extracted. Exiting pipeline.")
        return

    # 3. Group Extracted Records by Award Cluster Key
    clusters = {}
    for doc in extracted_documents:
        ckey = doc["AWARD_CLUSTER_KEY"]
        clusters.setdefault(ckey, []).append(doc)
        
    # 4. Pass 2 & Pass 3: Portfolio Accumulation & Reconciliation
    reconciliation_results = []
    for ckey, doc_list in clusters.items():
        synth = synthesize_portfolio_pass_2(doc_list, system_baselines)
        reconciliation_results.append(synth)

    # 5. Export Staging Data to SQLite
    conn = sqlite3.connect(OUTPUT_SQLITE_PATH)
    
    doc_df = pd.DataFrame(extracted_documents)
    for col in ["EXECUTION_DATE", "DOC_START_DATE", "DOC_END_DATE"]:
        if col in doc_df.columns:
            doc_df[col] = doc_df[col].astype(str)
            
    doc_df.to_sql("tbl_Documents", conn, if_exists="replace", index=False)
    
    recon_df = pd.DataFrame(reconciliation_results)
    recon_df.to_sql("tbl_Reconciliation_Summary", conn, if_exists="replace", index=False)
    conn.close()

    # 6. Export Final Audit Report Workbook
    OUTPUT_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_EXCEL_PATH, engine='openpyxl') as writer:
        recon_df.to_excel(writer, sheet_name="3Way_Reconciliation", index=False)
        doc_df.to_excel(writer, sheet_name="Document_Level_Detail", index=False)

    print(f"\nPipeline Execution Complete!")
    print(f"-> SQLite Staged Database: {OUTPUT_SQLITE_PATH}")
    print(f"-> Audit Workbook: {OUTPUT_EXCEL_PATH}")

if __name__ == "__main__":
    main()