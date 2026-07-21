import os
import re
import yaml
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

# ---------------------------------------------------------------------------
# PATH CONFIGURATION
# ---------------------------------------------------------------------------
TRIAGE_EXCEL_PATH = r"2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx"
ALN_CSV_PATH = r"ALN.csv"

PDF_SOURCE_DIR = Path(r"D:\0-Batch-AWARDS\processed_files")
MD_OUTPUT_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Review")
REVIEW_TABLE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"

MD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
REVIEW_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 1. LOAD & SANITIZE REFERENCE DATA (MASTER & ALN)
# ---------------------------------------------------------------------------
def normalize_aln_key(val):
    """Ensures ALN format is padded correctly (e.g., 47.07 -> 47.070)."""
    val = str(val).strip()
    if '.' in val:
        parts = val.split('.')
        agency = parts[0].zfill(2)
        prog = parts[1].ljust(3, '0')[:3]
        return f"{agency}.{prog}"
    return val

def load_reference_data(excel_path, aln_path):
    df_master = pd.read_excel(excel_path, sheet_name='MASTER')
    df_master = df_master.fillna('')
    
    proj_map, prop_map, oracle_map, fed_map = {}, {}, {}, {}
    
    for idx, row in df_master.iterrows():
        proj = str(row.get('CAYUSE_PROJECT_NUMBER', '')).replace('.0', '').strip()
        prop = str(row.get('CAYUSE_PROPOSAL_NUMBER', '')).replace('.0', '').strip()
        orc = str(row.get('ORACLE_AWARD_NUMBER', '')).replace('.0', '').strip()
        fed = str(row.get('FEDERAL_AWARD_IDENTIFIER', '')).replace('.0', '').strip()
        
        proj = '' if proj.lower() in ['nan', 'none', '0', ''] else proj
        prop = '' if prop.lower() in ['nan', 'none', '0', ''] else prop
        orc = '' if orc.lower() in ['nan', 'none', '0', ''] else orc
        fed = '' if fed.lower() in ['nan', 'none', '0', ''] else fed

        row_dict = row.to_dict()

        if proj: proj_map[proj] = row_dict
        if prop: prop_map[prop] = row_dict
        if orc: oracle_map[orc] = row_dict
        if fed and len(fed) >= 4: fed_map[fed] = row_dict

    aln_lookup = {}
    if Path(aln_path).exists():
        aln_df = pd.read_csv(aln_path, dtype=str).fillna('')
        for idx, row in aln_df.iterrows():
            raw_aln = str(row.get('ALN', '')).strip()
            title = str(row.get('ALN_PROGRAM_TITLE', '')).strip()
            if raw_aln and raw_aln.lower() != 'nan':
                norm_aln = normalize_aln_key(raw_aln)
                aln_lookup[norm_aln] = title
                aln_lookup[raw_aln] = title

    return proj_map, prop_map, oracle_map, fed_map, aln_lookup

# ---------------------------------------------------------------------------
# 2. HELPER ENGINES (ALN & MUTUALLY EXCLUSIVE HISTORICAL MATCHING)
# ---------------------------------------------------------------------------
def extract_aln_details(filename, page_texts, aln_lookup):
    """Scans filename and text for ALN/CFDA patterns (XX.XXX)."""
    aln_pattern = r'(?<!\d)(\d{2}\.\d{3})(?!\d)'
    
    # 1. Check Filename
    fn_matches = re.findall(aln_pattern, filename)
    if fn_matches:
        aln_num = fn_matches[0]
        title = aln_lookup.get(aln_num, aln_lookup.get(normalize_aln_key(aln_num), 'Title Not in ALN.csv'))
        return aln_num, title, 'Filename'

    # 2. Check Pages
    for page_idx, page_text in enumerate(page_texts):
        page_matches = re.findall(aln_pattern, page_text)
        if page_matches:
            aln_num = page_matches[0]
            title = aln_lookup.get(aln_num, aln_lookup.get(normalize_aln_key(aln_num), 'Title Not in ALN.csv'))
            return aln_num, title, f'Page {page_idx + 1}'

    return 'N/A', 'N/A', 'Not Found'

def classify_document(text):
    text_upper = text.upper()
    amendment_keywords = ['AMENDMENT', 'MODIFICATION', 'REVISION', 'NOGA AMENDMENT', 'EXTENSION', 'AMD']
    if any(kw in text_upper for kw in amendment_keywords):
        return "Amendment/Modification"
    return "Award/Notice of Award"

def progressive_match_with_crossref(filename, page_texts, proj_map, prop_map, oracle_map, fed_map):
    """
    1. Check Active MASTER Index (Pass 1: Front Pages, Pass 2: Deep Scan)
    2. If Unmatched in MASTER: Strictly evaluate whether it's a Project ID (A-prefix) OR Proposal ID (bare YY-XXXX)
    """
    full_text = " ".join(page_texts)
    pass1_target = f"{filename} " + " ".join(page_texts[:5])

    # --- PASS 1: Active MASTER (Front Pages + Filename) ---
    for proj, row in proj_map.items():
        if proj and re.search(r'(?<![A-Za-z0-9])' + re.escape(proj) + r'(?![A-Za-z0-9])', pass1_target, re.I):
            return row, "High", f"Matched Active CAYUSE_PROJECT_NUMBER ({proj}) [Pass 1]", None, None
    for prop, row in prop_map.items():
        if prop and re.search(r'(?<!\d)' + re.escape(prop) + r'(?!\d)', pass1_target):
            return row, "High", f"Matched Active CAYUSE_PROPOSAL_NUMBER ({prop}) [Pass 1]", None, None
    for orc, row in oracle_map.items():
        if orc and re.search(r'(?<!\d)' + re.escape(orc) + r'(?!\d)', pass1_target):
            return row, "Medium-High", f"Matched Active ORACLE_AWARD_NUMBER ({orc}) [Pass 1]", None, None
    for fed, row in fed_map.items():
        if fed and re.search(r'(?<![A-Za-z0-9])' + re.escape(fed) + r'(?![A-Za-z0-9])', pass1_target, re.I):
            return row, "Medium", f"Matched Active FEDERAL_AWARD_IDENTIFIER ({fed}) [Pass 1]", None, None

    # --- PASS 2: Active MASTER Deep Scan (Pages 6+) ---
    if len(page_texts) > 5:
        for page_idx in range(5, len(page_texts)):
            p_text = page_texts[page_idx]
            for proj, row in proj_map.items():
                if proj and re.search(r'(?<![A-Za-z0-9])' + re.escape(proj) + r'(?![A-Za-z0-9])', p_text, re.I):
                    return row, "Low-Medium", f"Matched Active CAYUSE_PROJECT_NUMBER ({proj}) [Pass 2: Page {page_idx+1}]", None, None
            for prop, row in prop_map.items():
                if prop and re.search(r'(?<!\d)' + re.escape(prop) + r'(?!\d)', p_text):
                    return row, "Low-Medium", f"Matched Active CAYUSE_PROPOSAL_NUMBER ({prop}) [Pass 2: Page {page_idx+1}]", None, None

    # --- PASS 3: LEGACY / HISTORICAL MUTUALLY EXCLUSIVE MATCHING ---
    # Search for A-prefix project code (A21-1107 or A21-1107-001) in filename OR full text
    proj_code_matches = re.findall(r'(?<![A-Za-z0-9])(A\d{2}-\d{3,4}(?:-\d{2,4})?)(?![A-Za-z0-9])', f"{filename} {full_text}", re.I)
    
    if proj_code_matches:
        found_proj = proj_code_matches[0].upper()
        # Strictly a Project ID -> Proposal ID remains N/A
        return None, "Legacy/Historical Project", f"Extracted Cayuse Project Code ({found_proj}) from text/filename", found_proj, "N/A"

    # Search for bare proposal number YY-XXXX or 8-digit in filename or text
    fn_code_match = re.search(r'(?<!\d)(\d{2}-\d{4}|\d{8})(?!\d)', f"{filename} {full_text}")
    if fn_code_match:
        cand_prop = fn_code_match.group(1)
        # Strictly a Proposal ID -> Project ID remains NOT_IN_MASTER
        return None, "Legacy/Historical Proposal", f"Extracted Proposal ID ({cand_prop}) from filename/text", "NOT_IN_MASTER", cand_prop

    return None, "Unmatched", "No matching keys found in MASTER tab, text, or Filename", "NOT_IN_MASTER", "N/A"

# ---------------------------------------------------------------------------
# 3. WORKER FUNCTION FOR SINGLE PDF
# ---------------------------------------------------------------------------
def process_single_pdf(pdf_path, proj_map, prop_map, oracle_map, fed_map, aln_lookup):
    try:
        doc = fitz.open(pdf_path)
        page_texts = [page.get_text() for page in doc]
        doc.close()

        # Build Markdown Body
        full_text_pages = [f"## Page {i+1}\n\n" + text for i, text in enumerate(page_texts)]
        full_markdown_body = "\n\n".join(full_text_pages)

        # Document Classification
        combined_head = f"{pdf_path.name} " + " ".join(page_texts[:3])
        doc_type = classify_document(combined_head)

        # Cross-Referenced Matching
        match_data, confidence, reason, leg_proj, leg_prop = progressive_match_with_crossref(
            pdf_path.name, page_texts, proj_map, prop_map, oracle_map, fed_map
        )

        aln_num, aln_title, aln_source = extract_aln_details(pdf_path.name, page_texts, aln_lookup)

        # Assign Key Identifiers
        if match_data:
            cayuse_proj_no = match_data.get('CAYUSE_PROJECT_NUMBER', 'UNMATCHED')
            cayuse_prop_no = match_data.get('CAYUSE_PROPOSAL_NUMBER', 'N/A')
            oracle_award_no = match_data.get('ORACLE_AWARD_NUMBER', 'N/A')
            lead_pi = match_data.get('LEAD_PI', 'UNKNOWN')
            fed_id = match_data.get('FEDERAL_AWARD_IDENTIFIER', 'N/A')
            sponsor = match_data.get('CAYUSE_SPONSOR_NAME', 'N/A')
            prefix_id = cayuse_proj_no if cayuse_proj_no != 'UNMATCHED' else f"PROP_{cayuse_prop_no}"
        else:
            cayuse_proj_no = leg_proj if leg_proj else "NOT_IN_MASTER"
            cayuse_prop_no = leg_prop if leg_prop else "N/A"
            oracle_award_no = "NOT_IN_MASTER"
            lead_pi = "UNKNOWN"
            fed_id = "N/A"
            sponsor = "N/A"
            
            if "Project" in confidence:
                prefix_id = f"HISTORICAL_PROJ_{cayuse_proj_no}"
            elif "Proposal" in confidence:
                prefix_id = f"HISTORICAL_PROP_{cayuse_prop_no}"
            else:
                prefix_id = "UNMATCHED"

        # YAML Frontmatter Header
        yaml_header = {
            'document_type': doc_type,
            'cayuse_project_number': cayuse_proj_no,
            'cayuse_proposal_number': cayuse_prop_no,
            'oracle_award_number': oracle_award_no,
            'lead_pi': lead_pi,
            'federal_award_identifier': fed_id,
            'sponsor_name': sponsor,
            'aln_number': aln_num,
            'aln_program_title': aln_title,
            'aln_source': aln_source,
            'match_confidence': confidence,
            'match_reason': reason,
            'original_pdf_name': pdf_path.name
        }

        yaml_str = "---\n" + yaml.dump(yaml_header, sort_keys=False) + "---\n\n"
        final_md_content = yaml_str + full_markdown_body

        # Save Markdown File
        new_md_filename = f"{prefix_id}_{pdf_path.stem}.md"
        new_md_filename = re.sub(r'[\\/*?:"<>|]', '_', new_md_filename)
        output_md_path = MD_OUTPUT_DIR / new_md_filename
        
        with open(output_md_path, 'w', encoding='utf-8') as f:
            f.write(final_md_content)

        return {
            'PDF_Path': str(pdf_path.resolve()),
            'Markdown_Path': str(output_md_path.resolve()),
            'Original_Filename': pdf_path.name,
            'Markdown_Filename': new_md_filename,
            'CAYUSE_PROJECT_NUMBER': cayuse_proj_no,
            'CAYUSE_PROPOSAL_NUMBER': cayuse_prop_no,
            'ORACLE_AWARD_NUMBER': oracle_award_no,
            'LEAD_PI': lead_pi,
            'ALN_NUMBER': aln_num,
            'ALN_PROGRAM_TITLE': aln_title,
            'ALN_SOURCE': aln_source,
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
            'CAYUSE_PROPOSAL_NUMBER': 'ERROR',
            'ORACLE_AWARD_NUMBER': 'ERROR',
            'LEAD_PI': 'ERROR',
            'ALN_NUMBER': 'ERROR',
            'ALN_PROGRAM_TITLE': 'ERROR',
            'ALN_SOURCE': 'ERROR',
            'Document_Type': 'ERROR',
            'Match_Confidence': 'Failed',
            'Match_Reason': f"Processing Error: {str(e)}"
        }

# ---------------------------------------------------------------------------
# 4. MAIN BATCH RUNNER
# ---------------------------------------------------------------------------
def run_pipeline():
    print("Loading Master Triage Index & ALN Lookup Table...")
    proj_map, prop_map, oracle_map, fed_map, aln_lookup = load_reference_data(
        TRIAGE_EXCEL_PATH, ALN_CSV_PATH
    )
    print(f"ALN Lookup loaded with {len(aln_lookup)} program titles.")

    pdf_files = [f for f in PDF_SOURCE_DIR.glob("*.pdf") if f.is_file()]
    print(f"Found {len(pdf_files)} top-level PDF files in {PDF_SOURCE_DIR}.")

    audit_records = []
    with ProcessPoolExecutor() as executor:
        futures = [
            executor.submit(process_single_pdf, pdf, proj_map, prop_map, oracle_map, fed_map, aln_lookup) 
            for pdf in pdf_files
        ]
        for idx, future in enumerate(as_completed(futures)):
            result = future.result()
            audit_records.append(result)
            if (idx + 1) % 1000 == 0 or (idx + 1) == len(pdf_files):
                print(f"Processed [{idx + 1}/{len(pdf_files)}] files...")

    # Output Audit Index Table
    df_audit = pd.DataFrame(audit_records)
    df_audit.to_excel(REVIEW_TABLE_PATH, index=False)
    
    print("\n" + "="*60)
    print("PIPELINE EXECUTION COMPLETE")
    print("="*60)
    print(f"Total Files Processed: {len(df_audit)}")
    print("\nMatch Confidence Breakdown:")
    print(df_audit['Match_Confidence'].value_counts())
    print("\nALN Extraction Summary:")
    print(df_audit['ALN_SOURCE'].value_counts())
    print(f"\n -> Markdown directory: {MD_OUTPUT_DIR}")
    print(f" -> Master review index: {REVIEW_TABLE_PATH}")

if __name__ == "__main__":
    run_pipeline()