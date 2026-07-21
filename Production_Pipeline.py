import os
import re
import yaml
import fitz  # PyMuPDF for fast PDF extraction
import pandas as pd
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

# ---------------------------------------------------------------------------
# CONFIGURATION & PATHS (UPDATED FOR YOUR LOCAL DRIVE)
# ---------------------------------------------------------------------------
TRIAGE_EXCEL_PATH = r"2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx"

# Windows directory paths
PDF_SOURCE_DIR = Path(r"D:\0-Batch-AWARDS\processed_files")
MD_OUTPUT_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Review")
REVIEW_TABLE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"

# Ensure destination directories exist prior to run
MD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REVIEW_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 1. LOAD MASTER REFERENCE TABLE
# ---------------------------------------------------------------------------
def load_master_lookup(excel_path):
    df_master = pd.read_excel(excel_path, sheet_name='MASTER')
    
    # Fill NA and standardize text
    df_master['CAYUSE_PROJECT_NUMBER'] = df_master['CAYUSE_PROJECT_NUMBER'].astype(str).str.strip()
    df_master['CAYUSE_PROPOSAL_NUMBER'] = df_master['CAYUSE_PROPOSAL_NUMBER'].astype(str).str.strip()
    df_master['ORACLE_AWARD_NUMBER'] = df_master['ORACLE_AWARD_NUMBER'].fillna(0).astype(int).astype(str).str.strip()
    df_master['FEDERAL_AWARD_IDENTIFIER'] = df_master['FEDERAL_AWARD_IDENTIFIER'].astype(str).str.strip()
    
    return df_master

# ---------------------------------------------------------------------------
# 2. MATCHING & CLASSIFICATION ENGINE
# ---------------------------------------------------------------------------
def classify_document(text):
    """Detects whether document is an Award or Amendment/Modification."""
    text_upper = text.upper()
    amendment_keywords = ['AMENDMENT', 'MODIFICATION', 'REVISION', 'NOGA AMENDMENT', 'EXTENSION']
    if any(kw in text_upper for kw in amendment_keywords):
        return "Amendment/Modification"
    return "Award/Notice of Award"

def match_pdf_to_master(filename, text_sample, df_master):
    """
    Tiered Cross-Referencing Engine:
    Returns (Matched Row Dict, Match Confidence, Match Reason)
    """
    combined_target = f"{filename} {text_sample}"
    
    # Tier 1: CAYUSE_PROJECT_NUMBER match
    for _, row in df_master.iterrows():
        proj_no = row['CAYUSE_PROJECT_NUMBER']
        if proj_no and proj_no != 'nan' and len(proj_no) > 3:
            if re.search(r'\b' + re.escape(proj_no) + r'\b', combined_target, re.IGNORECASE):
                return row.to_dict(), "High", f"Matched CAYUSE_PROJECT_NUMBER ({proj_no})"

    # Tier 2: CAYUSE_PROPOSAL_NUMBER match
    for _, row in df_master.iterrows():
        prop_no = row['CAYUSE_PROPOSAL_NUMBER']
        if prop_no and prop_no != 'nan' and len(prop_no) > 3:
            if re.search(r'\b' + re.escape(prop_no) + r'\b', combined_target, re.IGNORECASE):
                return row.to_dict(), "High", f"Matched CAYUSE_PROPOSAL_NUMBER ({prop_no})"

    # Tier 3: FEDERAL_AWARD_IDENTIFIER match
    for _, row in df_master.iterrows():
        fed_id = row['FEDERAL_AWARD_IDENTIFIER']
        if fed_id and fed_id not in ['nan', '0'] and len(fed_id) > 4:
            if re.search(r'\b' + re.escape(fed_id) + r'\b', combined_target, re.IGNORECASE):
                return row.to_dict(), "Medium-High", f"Matched FEDERAL_AWARD_IDENTIFIER ({fed_id})"

    # Tier 4: ORACLE_AWARD_NUMBER match
    for _, row in df_master.iterrows():
        orc_no = row['ORACLE_AWARD_NUMBER']
        if orc_no and orc_no not in ['nan', '0'] and len(orc_no) > 4:
            if re.search(r'\b' + re.escape(orc_no) + r'\b', combined_target, re.IGNORECASE):
                return row.to_dict(), "Medium", f"Matched ORACLE_AWARD_NUMBER ({orc_no})"

    # Fallback: Unmatched
    return None, "Low", "No matching keys found in filename or first 5 pages"

# ---------------------------------------------------------------------------
# 3. SINGLE PDF PROCESSING WORKER
# ---------------------------------------------------------------------------
def process_single_pdf(pdf_path, df_master_records):
    try:
        doc = fitz.open(pdf_path)
        
        # Read first 5 pages for indexing sample
        sample_pages = [page.get_text() for idx, page in enumerate(doc) if idx < 5]
        sample_text = "\n".join(sample_pages)
        
        # Read full text for markdown conversion
        full_text_pages = [f"## Page {i+1}\n\n" + page.get_text() for i, page in enumerate(doc)]
        full_markdown_body = "\n\n".join(full_text_pages)
        doc.close()

        # Execute Matcher
        match_data, confidence, reason = match_pdf_to_master(pdf_path.name, sample_text, df_master_records)
        doc_type = classify_document(sample_text)

        cayuse_proj_no = match_data.get('CAYUSE_PROJECT_NUMBER', 'UNMATCHED') if match_data else 'UNMATCHED'
        oracle_award_no = match_data.get('ORACLE_AWARD_NUMBER', 'UNMATCHED') if match_data else 'UNMATCHED'
        lead_pi = match_data.get('LEAD_PI', 'UNKNOWN') if match_data else 'UNKNOWN'

        # Construct YAML Frontmatter
        yaml_header = {
            'document_type': doc_type,
            'cayuse_project_number': cayuse_proj_no,
            'cayuse_proposal_number': match_data.get('CAYUSE_PROPOSAL_NUMBER', 'N/A') if match_data else 'N/A',
            'oracle_award_number': oracle_award_no,
            'lead_pi': lead_pi,
            'federal_award_identifier': match_data.get('FEDERAL_AWARD_IDENTIFIER', 'N/A') if match_data else 'N/A',
            'sponsor_name': match_data.get('CAYUSE_SPONSOR_NAME', 'N/A') if match_data else 'N/A',
            'match_confidence': confidence,
            'match_reason': reason,
            'original_pdf_name': pdf_path.name
        }

        yaml_str = "---\n" + yaml.dump(yaml_header, sort_keys=False) + "---\n\n"
        final_md_content = yaml_str + full_markdown_body

        # Save Markdown File
        new_md_filename = f"{cayuse_proj_no}_{pdf_path.stem}.md"
        output_md_path = MD_OUTPUT_DIR / new_md_filename
        
        with open(output_md_path, 'w', encoding='utf-8') as f:
            f.write(final_md_content)

        return {
            'PDF_Path': str(pdf_path.resolve()),
            'Markdown_Path': str(output_md_path.resolve()),
            'Original_Filename': pdf_path.name,
            'Markdown_Filename': new_md_filename,
            'CAYUSE_PROJECT_NUMBER': cayuse_proj_no,
            'ORACLE_AWARD_NUMBER': oracle_award_no,
            'LEAD_PI': lead_pi,
            'Document_Type': doc_type,
            'Match_Confidence': confidence,
            'Match_Reason': reason
        }

    except Exception as e:
        return {
            'PDF_Path': str(pdf_path.resolve()),
            'Markdown_Path': 'ERROR',
            'Original_Filename': pdf_path.name,
            'Markdown_Filename': 'ERROR',
            'CAYUSE_PROJECT_NUMBER': 'ERROR',
            'ORACLE_AWARD_NUMBER': 'ERROR',
            'LEAD_PI': 'ERROR',
            'Document_Type': 'ERROR',
            'Match_Confidence': 'Failed',
            'Match_Reason': f"Processing Error: {str(e)}"
        }

# ---------------------------------------------------------------------------
# 4. BATCH EXECUTION & AUDIT LOG GENERATION
# ---------------------------------------------------------------------------
def run_pipeline():
    print("Loading Master Triage Index...")
    df_master = load_master_lookup(TRIAGE_EXCEL_PATH)
    
    pdf_files = list(PDF_SOURCE_DIR.glob("*.pdf"))
    print(f"Found {len(pdf_files)} PDF files to process.")

    audit_records = []
    
    # Parallel processing using multi-threading/processing for speed
    with ProcessPoolExecutor() as executor:
        futures = [executor.submit(process_single_pdf, pdf, df_master) for pdf in pdf_files]
        for idx, future in enumerate(as_completed(futures)):
            result = future.result()
            audit_records.append(result)
            if (idx + 1) % 500 == 0 or (idx + 1) == len(pdf_files):
                print(f"Processed [{idx + 1}/{len(pdf_files)}] files...")

    # Output Audit Index Table
    df_audit = pd.DataFrame(audit_records)
    df_audit.to_excel(REVIEW_TABLE_PATH, index=False)
    print(f"\nPipeline Complete! Audit review table saved to: {REVIEW_TABLE_PATH}")

if __name__ == "__main__":
    run_pipeline()