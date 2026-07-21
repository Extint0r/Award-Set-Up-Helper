import os
import re
import glob
import fitz  # PyMuPDF
import pandas as pd
import yaml
from pathlib import Path
from datetime import datetime

# ==============================================================================
# 0. TERMINAL WARNING SUPPRESSION & PATH CONFIGURATION
# ==============================================================================
# Suppress non-fatal MuPDF C-library rendering syntax noise
fitz.TOOLS.mupdf_display_errors(False)

# ==============================================================================
# PATH & ENVIRONMENT CONFIGURATION
# ==============================================================================
TRIAGE_EXCEL_PATH = r"2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx"
ALN_CSV_PATH = r"ALN.csv"

PDF_SOURCE_DIR = Path(r"D:\0-Batch-AWARDS\processed_files")
MD_OUTPUT_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Review")

# Output Artifacts
REVIEW_TABLE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"
OUTPUT_SQLITE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.db"

# Ensure directories exist
MD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REVIEW_DIR.mkdir(parents=True, exist_ok=True)

# ==============================================================================
# 1. REGEX PATTERNS & REGEX MATCHING ENGINES
# ==============================================================================
# Standard Identifiers
RE_CAYUSE_PROJ = re.compile(r'\b(\d{2}-\d{4})\b')
RE_CAYUSE_PROP = re.compile(r'\b(A\d{2}-\d{4})\b')
RE_ORACLE_NUM  = re.compile(r'\b(\d{6})\b')
RE_BANNER_UID  = re.compile(r'\b(R[0-9A-Z]{5})\b', re.IGNORECASE)

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

# Budget Keyword Signals (Strategy 2 Manifest Tagging)
BUDGET_PAGE_KEYWORDS = [
    'budget summary', 'direct cost', 'indirect cost', 'f&a rate',
    'fringe benefit', 'personnel justification', 'total requested',
    'modified total direct', 'equipment', 'travel', 'contractual',
    'subaward budget', 'budget justification'
]

# ==============================================================================
# 2. HELPER FUNCTIONS
# ==============================================================================
def normalize_date(date_str: str) -> str:
    """Parses various raw date strings into standard YYYY-MM-DD format."""
    if not date_str:
        return ""
    try:
        dt = pd.to_datetime(date_str, errors='coerce')
        if pd.notnull(dt):
            return dt.strftime('%Y-%m-%d')
    except Exception:
        pass
    return ""

def extract_action_tags(filename: str, text_head: str) -> str:
    """Extracts administrative action tags (SubA, Amd01, NCE, etc.) from filename and text."""
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
    
    for line in lines[:100]:  # Limit scan to top administrative portion
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
    """Identifies candidate pages containing budget tables or justifications (Strategy 2)."""
    budget_pages = []
    for idx, page in enumerate(doc):
        p_text = page.get_text().lower()
        matches = sum(1 for kw in BUDGET_PAGE_KEYWORDS if kw in p_text)
        # Higher keyword density indicates explicit budget table/justification page
        if matches >= 2:
            budget_pages.append(idx + 1)  # 1-indexed page numbering
    return len(budget_pages) > 0, budget_pages

# ==============================================================================
# 3. PASS 1: PER-FILE PROCESSING ENGINE
# ==============================================================================
def process_pdf_pass_1(pdf_path: Path) -> dict:
    filename = pdf_path.name
    doc = fitz.open(pdf_path)
    
    # Extract raw text (First 3 pages for administrative headers)
    header_text = ""
    for i in range(min(3, len(doc))):
        header_text += doc[i].get_text() + "\n"
        
    full_text = "\n".join([page.get_text() for page in doc])

    # 1. Identifier Extraction
    cayuse_proj = RE_CAYUSE_PROJ.findall(filename) or RE_CAYUSE_PROJ.findall(header_text)
    cayuse_prop = RE_CAYUSE_PROP.findall(filename) or RE_CAYUSE_PROP.findall(header_text)
    oracle_num  = RE_ORACLE_NUM.findall(filename) or RE_ORACLE_NUM.findall(header_text)
    banner_uid  = RE_BANNER_UID.findall(filename) or RE_BANNER_UID.findall(header_text)

    cay_proj_val = cayuse_proj[0] if cayuse_proj else ""
    cay_prop_val = cayuse_prop[0] if cayuse_prop else ""
    oracle_val   = oracle_num[0] if oracle_num else ""
    banner_val   = banner_uid[0].upper() if banner_uid else ""

    # 2. Action Tag Extraction
    action_tag = extract_action_tags(filename, header_text)

    # 3. Document Performance Dates (Pass 1)
    doc_start, doc_end = scan_dates_from_text(header_text)

    # 4. Strategy 2 Budget Table Detection
    has_budget, budget_pages = detect_budget_pages(doc)

    # 5. Determine Award Match Hierarchy
    match_conf = "Unmatched"
    match_reason = "No known master key or legacy UID matched"
    
    if oracle_val and oracle_val != "000000":
        match_conf = "Active Index Match"
        match_reason = "Matched Oracle Award Number"
    elif banner_val:
        match_conf = "Legacy Banner Match"
        match_reason = f"Matched Legacy Banner Award UID ({banner_val})"
    elif cay_proj_val:
        match_conf = "Cayuse Inference"
        match_reason = "Matched Cayuse Project Number"
    elif cay_prop_val:
        match_conf = "Cayuse Proposal Match"
        match_reason = "Matched Cayuse Proposal Number"

    # Define Unified Cluster Key for Pass 2 Grouping
    cluster_key = oracle_val or banner_val or cay_proj_val or cay_prop_val or "UNMATCHED"

    # Generate Markdown Output File
    md_filename = f"{match_conf.replace(' ', '_')}_{pdf_path.stem}.md"
    md_path = MD_OUTPUT_DIR / md_filename

    frontmatter = {
        "original_filename": filename,
        "cayuse_project_number": cay_proj_val,
        "cayuse_proposal_number": cay_prop_val,
        "oracle_award_number": oracle_val,
        "banner_award_uid": banner_val,
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
        "ORACLE_AWARD_NUMBER": oracle_val,
        "BANNER_AWARD_UID": banner_val,
        "LEAD_PI": "EXTRACTED_PI",  # Reconciled in Master Merge
        "ALN_NUMBER": "N/A",        # Reconciled in Master Merge
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
# 4. PASS 2: PORTFOLIO-LEVEL DATE AGGREGATION & INDEX BUILD
# ==============================================================================
def run_pipeline():
    print("🚀 Starting v6 Processing Pipeline (Strategy 2 Ready)...")
    
    # FIX 1: Use PDF_SOURCE_DIR defined in config block
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
    # PASS 2: Award Cluster Aggregation Logic
    # --------------------------------------------------------------------------
    print("\n🔄 Running Pass 2: Aggregating Award Portfolio Lifecycles...")

    # Calculate min start date and max end date per cluster (excluding unmatched)
    valid_clusters = df[df['AWARD_CLUSTER_KEY'] != 'UNMATCHED']

    # Groupby aggregations
    agg_dates = valid_clusters.groupby('AWARD_CLUSTER_KEY').agg(
        AWARD_CUMULATIVE_START=('DOC_START_DATE', lambda s: min([d for d in s if d] or [''])),
        AWARD_ULTIMATE_END=('DOC_END_DATE', lambda s: max([d for d in s if d] or ['']))
    ).reset_index()

    # Merge aggregated dates back into dataframe
    df = df.merge(agg_dates, on='AWARD_CLUSTER_KEY', how='left')
    df['AWARD_CUMULATIVE_START'] = df['AWARD_CUMULATIVE_START'].fillna('')
    df['AWARD_ULTIMATE_END']     = df['AWARD_ULTIMATE_END'].fillna('')

    # Calculate Chronological Sequence Tag (e.g., "Doc 1 of 3")
    df['DOC_SORT_KEY'] = df['DOC_START_DATE'].replace('', '9999-99-99')
    df = df.sort_values(by=['AWARD_CLUSTER_KEY', 'DOC_SORT_KEY', 'Original_Filename'])
    
    df['CLUSTER_TOTAL_DOCS'] = df.groupby('AWARD_CLUSTER_KEY')['Original_Filename'].transform('count')
    df['CLUSTER_DOC_RANK']   = df.groupby('AWARD_CLUSTER_KEY').cumcount() + 1
    
    df['CHRONO_SEQUENCE'] = df.apply(
        lambda r: f"Doc {r['CLUSTER_DOC_RANK']} of {r['CLUSTER_TOTAL_DOCS']}" 
        if r['AWARD_CLUSTER_KEY'] != 'UNMATCHED' else "N/A", 
        axis=1
    )

    # Clean up temporary processing columns
    df.drop(columns=['AWARD_CLUSTER_KEY', 'DOC_SORT_KEY', 'CLUSTER_TOTAL_DOCS', 'CLUSTER_DOC_RANK'], inplace=True)

    # Reorder columns to match canonical schema
    canonical_columns = [
        "PDF_Path", "Markdown_Path", "Original_Filename", "Markdown_Filename",
        "CAYUSE_PROJECT_NUMBER", "CAYUSE_PROPOSAL_NUMBER", "ORACLE_AWARD_NUMBER",
        "BANNER_AWARD_UID", "LEAD_PI", "ALN_NUMBER", "ALN_PROGRAM_TITLE",
        "ALN_SOURCE", "DOCUMENT_ACTION_TAG", "DOC_START_DATE", "DOC_END_DATE",
        "AWARD_CUMULATIVE_START", "AWARD_ULTIMATE_END", "CHRONO_SEQUENCE",
        "HAS_BUDGET_TABLE", "BUDGET_PAGES", "Match_Confidence", "Match_Reason"
    ]
    df = df[canonical_columns]

    # Save to Excel & SQLite Database
    # FIX 2: Use REVIEW_TABLE_PATH defined in config block
    df.to_excel(REVIEW_TABLE_PATH, index=False)
    
    import sqlite3
    conn = sqlite3.connect(OUTPUT_SQLITE_PATH)
    df.to_sql("tbl_Documents", conn, if_exists="replace", index=False)
    conn.close()

    # Pipeline Summary Output
    print("\n=================== PIPELINE EXECUTION SUMMARY ===================")
    print(f"Total Files Processed   : {len(df)}")
    print("Match Confidence Breakdown:")
    print(df['Match_Confidence'].value_counts().to_string())
    print("------------------------------------------------------------------")
    print(f"Budget Tables Tagged    : {df['HAS_BUDGET_TABLE'].sum()} documents")
    # FIX 3: Use REVIEW_TABLE_PATH defined in config block
    print(f"Output Audit Excel      : {REVIEW_TABLE_PATH}")
    print(f"Output Portfolio SQLite : {OUTPUT_SQLITE_PATH}")
    print("==================================================================")

if __name__ == "__main__":
    run_pipeline()