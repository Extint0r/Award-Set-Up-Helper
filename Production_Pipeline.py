import os
import re
import glob
import fitz  # PyMuPDF
import pandas as pd
import yaml
import sqlite3
from pathlib import Path
from datetime import datetime

# ==============================================================================
# 0. TERMINAL WARNING SUPPRESSION & PATH CONFIGURATION
# ==============================================================================
fitz.TOOLS.mupdf_display_errors(False)

# Path & Environment Configuration
TRIAGE_EXCEL_PATH = Path(r"2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx")
ALN_CSV_PATH      = Path(r"ALN.csv")

PDF_SOURCE_DIR    = Path(r"D:\0-Batch-AWARDS\processed_files")
MD_OUTPUT_DIR     = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR        = Path(r"D:\0-Batch-AWARDS\processed_files\Review")

# Output Artifacts
OUTPUT_EXCEL_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"
OUTPUT_SQLITE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.db"

# Ensure output directories exist
MD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REVIEW_DIR.mkdir(parents=True, exist_ok=True)

# ==============================================================================
# 1. REGEX PATTERNS & BOUNDED MATCHING ENGINES
# ==============================================================================
# Bounded Identifiers
RE_CAYUSE_PROJ = re.compile(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])')
RE_CAYUSE_PROP = re.compile(r'(?<![A-Za-z0-9])(A\d{2}-\d{4})(?![A-Za-z0-9])', re.IGNORECASE)

# Oracle Award Numbers MUST start with 1-9 (No leading zeros permitted)
RE_ORACLE_NUM  = re.compile(r'\b([1-9]\d{5})\b')

# Banner UIDs MUST start with R, total 6 chars, contain >=1 digit, and handle _ boundaries
RE_BANNER_UID  = re.compile(r'(?<![A-Za-z0-9])(R(?=[A-Za-z0-9]{0,4}[0-9])[0-9A-Za-z]{5})(?![A-Za-z0-9])', re.IGNORECASE)

# Document Action Tags
RE_ACTION_TAGS = re.compile(
    r'\b(SubA(?:ward)?|SubB|Subgrant|Amd\s*\d*|Amendment\s*\d*|NCE|No-Cost\s*Ext\w*|'
    r'inv(?:ention)?\s*rpt|progress\s*rpt|PR|Y\d+-\d+\s*PRs?|MOU|A-133|Supplement)\b', 
    re.IGNORECASE
)

# Date Search Patterns
RE_DATE_GENERIC = re.compile(
    r'\b(\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2}|'
    r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4})\b', 
    re.IGNORECASE
)

# Budget Keyword Signals (Strategy 2 Tagging)
BUDGET_PAGE_KEYWORDS = [
    'budget summary', 'direct cost', 'indirect cost', 'f&a rate',
    'fringe benefit', 'personnel justification', 'total requested',
    'modified total direct', 'equipment', 'travel', 'contractual',
    'subaward budget', 'budget justification'
]

# ==============================================================================
# 2. MASTER LOOKUP DATA INGESTION
# ==============================================================================
LOOKUP_PROJ_TO_ORACLE = {}
LOOKUP_PROP_TO_ORACLE = {}
LOOKUP_BANNER_TO_ORACLE = {}
LOOKUP_ORACLE_TO_PI = {}
LOOKUP_ORACLE_TO_ALN = {}

def load_master_lookups():
    """Loads Master Triage Excel and ALN CSV into in-memory resolution dictionaries."""
    global LOOKUP_PROJ_TO_ORACLE, LOOKUP_PROP_TO_ORACLE
    global LOOKUP_BANNER_TO_ORACLE, LOOKUP_ORACLE_TO_PI, LOOKUP_ORACLE_TO_ALN
    
    if TRIAGE_EXCEL_PATH.exists():
        try:
            df_tr = pd.read_excel(TRIAGE_EXCEL_PATH)
            
            # Standardize columns
            df_tr.columns = [str(c).strip().upper() for c in df_tr.columns]
            
            for _, r in df_tr.iterrows():
                oracle = str(r.get('ORACLE_AWARD_NUMBER', '')).split('.')[0].strip()
                proj   = str(r.get('CAYUSE_PROJECT_NUMBER', '')).strip()
                prop   = str(r.get('CAYUSE_PROPOSAL_NUMBER', '')).strip()
                banner = str(r.get('BANNER_AWARD_UID', '')).strip().upper()
                pi     = str(r.get('LEAD_PI', '')).strip()
                aln    = str(r.get('ALN_NUMBER', '')).strip()

                if oracle and oracle != 'nan' and RE_ORACLE_NUM.match(oracle):
                    if proj and proj != 'nan':
                        LOOKUP_PROJ_TO_ORACLE[proj] = oracle
                    if prop and prop != 'nan':
                        LOOKUP_PROP_TO_ORACLE[prop] = oracle
                    if banner and banner != 'nan':
                        LOOKUP_BANNER_TO_ORACLE[banner] = oracle
                    if pi and pi != 'nan':
                        LOOKUP_ORACLE_TO_PI[oracle] = pi
                    if aln and aln != 'nan':
                        LOOKUP_ORACLE_TO_ALN[oracle] = aln
                        
            print(f"✅ Loaded Master Triage Lookups: {len(LOOKUP_PROJ_TO_ORACLE)} Project mappings, {len(LOOKUP_BANNER_TO_ORACLE)} Banner mappings.")
        except Exception as e:
            print(f"⚠️ Warning loading Master Triage Excel: {e}")
    else:
        print(f"⚠️ Master Triage File not found at {TRIAGE_EXCEL_PATH}. Proceeding with extraction rules.")

# ==============================================================================
# 3. HELPER FUNCTIONS
# ==============================================================================
def normalize_date(date_str: str) -> str:
    """Parses date string into YYYY-MM-DD format with 1980-2035 sanity bounds."""
    if not date_str:
        return ""
    try:
        dt = pd.to_datetime(date_str, errors='coerce')
        if pd.notnull(dt) and 1980 <= dt.year <= 2035:
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return ""

def extract_action_tags(filename: str, text_head: str) -> str:
    """Extracts administrative action tags from filename and text."""
    matches = set()
    for m in RE_ACTION_TAGS.findall(filename):
        matches.add(m.strip().replace(" ", ""))
    for m in RE_ACTION_TAGS.findall(text_head[:1000]):
        matches.add(m.strip().replace(" ", ""))
    return "; ".join(sorted(matches)) if matches else "Standard Award"

def scan_dates_from_text(text: str):
    """Scans header text for performance start and end dates."""
    start_date, end_date = "", ""
    lines = text.split('\n')
    
    for line in lines[:100]:
        line_lower = line.lower()
        if any(kw in line_lower for kw in ['start date', 'effective date', 'period of performance start', 'from']):
            m = RE_DATE_GENERIC.search(line)
            if m and not start_date:
                start_date = normalize_date(m.group(1))
        if any(kw in line_lower for kw in ['end date', 'expiration date', 'period of performance end', 'through', 'valid to']):
            m = RE_DATE_GENERIC.search(line)
            if m and not end_date:
                end_date = normalize_date(m.group(1))
                
    return start_date, end_date

def detect_budget_pages(doc: fitz.Document):
    """Identifies candidate pages containing budget schedules (Strategy 2)."""
    budget_pages = []
    for idx, page in enumerate(doc):
        p_text = page.get_text().lower()
        matches = sum(1 for kw in BUDGET_PAGE_KEYWORDS if kw in p_text)
        if matches >= 2:
            budget_pages.append(idx + 1)
    return len(budget_pages) > 0, budget_pages

# ==============================================================================
# 4. PASS 1: PER-FILE EXTRACTION & RESOLUTION ENGINE
# ==============================================================================
def process_pdf_pass_1(pdf_path: Path) -> dict:
    filename = pdf_path.name
    doc = fitz.open(pdf_path)
    
    header_text = ""
    for i in range(min(3, len(doc))):
        header_text += doc[i].get_text() + "\n"
        
    full_text = "\n".join([page.get_text() for page in doc])

    # --- Step 1: Identifier Extraction from Filename (Primary Baseline) ---
    fn_cayuse_proj = RE_CAYUSE_PROJ.findall(filename)
    fn_cayuse_prop = RE_CAYUSE_PROP.findall(filename)
    fn_oracle_num  = RE_ORACLE_NUM.findall(filename)
    fn_banner_uid  = RE_BANNER_UID.findall(filename)

    cay_proj_val = fn_cayuse_proj[0] if fn_cayuse_proj else ""
    cay_prop_val = fn_cayuse_prop[0] if fn_cayuse_prop else ""
    oracle_val   = fn_oracle_num[0] if fn_oracle_num else ""
    banner_val   = fn_banner_uid[0].upper() if fn_banner_uid else ""

    # --- Step 2: Master Lookup Resolution (Querying Master Triage) ---
    resolved_oracle = oracle_val
    if not resolved_oracle:
        if cay_proj_val in LOOKUP_PROJ_TO_ORACLE:
            resolved_oracle = LOOKUP_PROJ_TO_ORACLE[cay_proj_val]
        elif cay_prop_val in LOOKUP_PROP_TO_ORACLE:
            resolved_oracle = LOOKUP_PROP_TO_ORACLE[cay_prop_val]
        elif banner_val in LOOKUP_BANNER_TO_ORACLE:
            resolved_oracle = LOOKUP_BANNER_TO_ORACLE[banner_val]

    # --- Step 3: Constrained Header Text Fallback (Only if still unresolved) ---
    if not resolved_oracle and not cay_proj_val and not banner_val:
        txt_cayuse_proj = RE_CAYUSE_PROJ.findall(header_text)
        txt_cayuse_prop = RE_CAYUSE_PROP.findall(header_text)
        txt_oracle_num  = RE_ORACLE_NUM.findall(header_text)
        txt_banner_uid  = RE_BANNER_UID.findall(header_text)

        cay_proj_val = cay_proj_val or (txt_cayuse_proj[0] if txt_cayuse_proj else "")
        cay_prop_val = cay_prop_val or (txt_cayuse_prop[0] if txt_cayuse_prop else "")
        banner_val   = banner_val or (txt_banner_uid[0].upper() if txt_banner_uid else "")
        
        # Re-check lookup with header text values
        if cay_proj_val in LOOKUP_PROJ_TO_ORACLE:
            resolved_oracle = LOOKUP_PROJ_TO_ORACLE[cay_proj_val]
        elif txt_oracle_num:
            resolved_oracle = txt_oracle_num[0]

    # --- Step 4: Metadata Reconciliation ---
    lead_pi = LOOKUP_ORACLE_TO_PI.get(resolved_oracle, "EXTRACTED_PI")
    aln_num = LOOKUP_ORACLE_TO_ALN.get(resolved_oracle, "N/A")

    # --- Step 5: Action Tag & Date Extraction ---
    action_tag = extract_action_tags(filename, header_text)
    doc_start, doc_end = scan_dates_from_text(header_text)
    has_budget, budget_pages = detect_budget_pages(doc)

    # --- Step 6: Determine Match Confidence & Canonical Cluster Key ---
    match_conf = "Unmatched"
    match_reason = "No known master key or legacy UID matched"
    
    if resolved_oracle:
        match_conf = "Active Index Match"
        match_reason = f"Matched Oracle Award Number ({resolved_oracle})"
    elif banner_val:
        match_conf = "Legacy Banner Match"
        match_reason = f"Matched Legacy Banner Award UID ({banner_val})"
    elif cay_proj_val:
        match_conf = "Cayuse Inference"
        match_reason = f"Matched Cayuse Project Number ({cay_proj_val})"
    elif cay_prop_val:
        match_conf = "Cayuse Proposal Match"
        match_reason = f"Matched Cayuse Proposal Number ({cay_prop_val})"

    # Unified Primary Cluster Key (Ensures related documents group together)
    cluster_key = resolved_oracle or banner_val or cay_proj_val or cay_prop_val or "UNMATCHED"

    # Generate Structured Markdown
    md_filename = f"{match_conf.replace(' ', '_')}_{pdf_path.stem}.md"
    md_path = MD_OUTPUT_DIR / md_filename

    frontmatter = {
        "original_filename": filename,
        "cayuse_project_number": cay_proj_val,
        "cayuse_proposal_number": cay_prop_val,
        "oracle_award_number": resolved_oracle,
        "banner_award_uid": banner_val,
        "lead_pi": lead_pi,
        "aln_number": aln_num,
        "action_tag": action_tag,
        "doc_start_date": doc_start,
        "doc_end_date": doc_end,
        "has_budget_table": has_budget,
        "budget_pages": budget_pages,
        "match_confidence": match_conf,
        "match_reason": match_reason
    }

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("---\n")
        yaml.dump(frontmatter, f, default_flow_style=False)
        f.write("---\n\n")
        f.write(f"# Document: {filename}\n\n")
        f.write(full_text)

    doc.close()

    return {
        "PDF_Path": str(pdf_path),
        "Markdown_Path": str(md_path),
        "Original_Filename": filename,
        "Markdown_Filename": md_filename,
        "CAYUSE_PROJECT_NUMBER": cay_proj_val,
        "CAYUSE_PROPOSAL_NUMBER": cay_prop_val,
        "ORACLE_AWARD_NUMBER": resolved_oracle,
        "BANNER_AWARD_UID": banner_val,
        "LEAD_PI": lead_pi,
        "ALN_NUMBER": aln_num,
        "ALN_PROGRAM_TITLE": "N/A",
        "ALN_SOURCE": "N/A",
        "DOCUMENT_ACTION_TAG": action_tag,
        "DOC_START_DATE": doc_start,
        "DOC_END_DATE": doc_end,
        "AWARD_CLUSTER_KEY": cluster_key,
        "HAS_BUDGET_TABLE": has_budget,
        "BUDGET_PAGES": str(budget_pages),
        "Match_Confidence": match_conf,
        "Match_Reason": match_reason
    }

# ==============================================================================
# 5. PASS 2: PORTFOLIO AGGREGATION & RELATIONAL INDEX BUILD
# ==============================================================================
def run_pipeline():
    print("🚀 Initializing Pipeline Processing...")
    
    # Load Master Lookups
    load_master_lookups()

    pdf_files = list(PDF_SOURCE_DIR.glob("*.pdf"))
    print(f"📁 Found {len(pdf_files)} PDF documents to process.")

    records = []
    for idx, pdf in enumerate(pdf_files, 1):
        try:
            rec = process_pdf_pass_1(pdf)
            records.append(rec)
        except Exception as e:
            print(f"⚠️ Error processing {pdf.name}: {e}")

        if idx % 1000 == 0 or idx == len(pdf_files):
            print(f"  -> Processed {idx}/{len(pdf_files)} files...")

    df = pd.DataFrame(records)

    # --------------------------------------------------------------------------
    # PASS 2: Portfolio Lifecycle Aggregation Logic
    # --------------------------------------------------------------------------
    print("\n🔄 Running Pass 2: Aggregating Portfolio Lifecycles & Chronology...")

    valid_clusters = df[df['AWARD_CLUSTER_KEY'] != 'UNMATCHED']

    # Aggregate start and end dates per cluster
    agg_dates = valid_clusters.groupby('AWARD_CLUSTER_KEY').agg(
        AWARD_CUMULATIVE_START=('DOC_START_DATE', lambda s: min([d for d in s if d] or [''])),
        AWARD_ULTIMATE_END=('DOC_END_DATE', lambda s: max([d for d in s if d] or ['']))
    ).reset_index()

    # Merge lifecycle dates back into portfolio
    df = df.merge(agg_dates, on='AWARD_CLUSTER_KEY', how='left')
    df['AWARD_CUMULATIVE_START'] = df['AWARD_CUMULATIVE_START'].fillna('')
    df['AWARD_ULTIMATE_END']     = df['AWARD_ULTIMATE_END'].fillna('')

    # Calculate Chronological Sequence Tag (e.g. "Doc 1 of 3")
    df['DOC_SORT_KEY'] = df['DOC_START_DATE'].replace('', '9999-99-99')
    df = df.sort_values(by=['AWARD_CLUSTER_KEY', 'DOC_SORT_KEY', 'Original_Filename'])
    
    df['CLUSTER_TOTAL_DOCS'] = df.groupby('AWARD_CLUSTER_KEY')['Original_Filename'].transform('count')
    df['CLUSTER_DOC_RANK']   = df.groupby('AWARD_CLUSTER_KEY').cumcount() + 1
    
    df['CHRONO_SEQUENCE'] = df.apply(
        lambda r: f"Doc {r['CLUSTER_DOC_RANK']} of {r['CLUSTER_TOTAL_DOCS']}" 
        if r['AWARD_CLUSTER_KEY'] != 'UNMATCHED' else "N/A", 
        axis=1
    )

    # Clean temporary calculation columns
    df.drop(columns=['AWARD_CLUSTER_KEY', 'DOC_SORT_KEY', 'CLUSTER_TOTAL_DOCS', 'CLUSTER_DOC_RANK'], inplace=True)

    # Enforce Canonical Output Schema
    canonical_columns = [
        "PDF_Path", "Markdown_Path", "Original_Filename", "Markdown_Filename",
        "CAYUSE_PROJECT_NUMBER", "CAYUSE_PROPOSAL_NUMBER", "ORACLE_AWARD_NUMBER",
        "BANNER_AWARD_UID", "LEAD_PI", "ALN_NUMBER", "ALN_PROGRAM_TITLE",
        "ALN_SOURCE", "DOCUMENT_ACTION_TAG", "DOC_START_DATE", "DOC_END_DATE",
        "AWARD_CUMULATIVE_START", "AWARD_ULTIMATE_END", "CHRONO_SEQUENCE",
        "HAS_BUDGET_TABLE", "BUDGET_PAGES", "Match_Confidence", "Match_Reason"
    ]
    df = df[canonical_columns]

    # Save to Master Audit Excel
    df.to_excel(OUTPUT_EXCEL_PATH, index=False)
    
    # Save to SQLite Relational DB
    conn = sqlite3.connect(OUTPUT_SQLITE_PATH)
    df.to_sql("tbl_Documents", conn, if_exists="replace", index=False)
    conn.close()

    # Summary Stats
    print("\n=================== PIPELINE EXECUTION SUMMARY ===================")
    print(f"Total Files Processed   : {len(df)}")
    print("Match Confidence Breakdown:")
    print(df['Match_Confidence'].value_counts().to_string())
    print("------------------------------------------------------------------")
    print(f"Budget Tables Tagged    : {df['HAS_BUDGET_TABLE'].sum()} documents")
    print(f"Output Audit Excel      : {OUTPUT_EXCEL_PATH}")
    print(f"Output Portfolio SQLite : {OUTPUT_SQLITE_PATH}")
    print("==================================================================")

if __name__ == "__main__":
    run_pipeline()