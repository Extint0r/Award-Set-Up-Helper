import os
import re
import yaml
import fitz  # PyMuPDF
import difflib
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
# 1. LOAD & SANITIZE REFERENCE DATA (MASTER, ALN & PIs)
# ---------------------------------------------------------------------------
def clean_str_id(val):
    """Sanitizes floating point artifacts (.0) and converts inputs to clean strings."""
    if pd.isna(val) or val is None:
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return '' if s.lower() in ['nan', 'none', '0', ''] else s

def normalize_aln_key(val):
    """Standardizes ALN formatting (e.g. 81.87 -> 81.087)."""
    val = clean_str_id(val)
    if '.' in val:
        parts = val.split('.')
        agency = parts[0].zfill(2)
        prog = parts[1].ljust(3, '0')[:3]
        return f"{agency}.{prog}"
    return val

def load_reference_data(excel_path, aln_path):
    df_master = pd.read_excel(excel_path, sheet_name='MASTER').fillna('')
    
    proj_map, prop_map, oracle_map, fed_map = {}, {}, {}, {}
    master_pi_list = [p for p in df_master['LEAD_PI'].dropna().astype(str).str.strip().unique().tolist() if p]
    
    for idx, row in df_master.iterrows():
        proj = clean_str_id(row.get('CAYUSE_PROJECT_NUMBER', ''))
        prop = clean_str_id(row.get('CAYUSE_PROPOSAL_NUMBER', ''))
        orc = clean_str_id(row.get('ORACLE_AWARD_NUMBER', ''))
        fed = clean_str_id(row.get('FEDERAL_AWARD_IDENTIFIER', ''))
        
        row_dict = row.to_dict()
        for k, v in row_dict.items():
            if isinstance(v, float) and v.is_integer():
                row_dict[k] = str(int(v))
            elif isinstance(v, float) and pd.isna(v):
                row_dict[k] = ''

        if proj: proj_map[proj.upper()] = row_dict
        if prop: prop_map[prop] = row_dict
        if orc: oracle_map[orc.upper()] = row_dict
        if fed and len(fed) >= 4: fed_map[fed.upper()] = row_dict

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

    return proj_map, prop_map, oracle_map, fed_map, aln_lookup, master_pi_list

# ---------------------------------------------------------------------------
# 2. HELPER ENGINES (ALN, STICKY PI RECONCILIATION & LEGACY RESOLUTION)
# ---------------------------------------------------------------------------
def extract_aln_details(filename, page_texts, aln_lookup):
    aln_pattern = r'(?<!\d)(\d{2}\.\d{3})(?!\d)'
    fn_matches = re.findall(aln_pattern, filename)
    if fn_matches:
        aln_num = fn_matches[0]
        title = aln_lookup.get(aln_num, aln_lookup.get(normalize_aln_key(aln_num), 'Title Not in ALN.csv'))
        return aln_num, title, 'Filename'

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

def extract_and_reconcile_pi(filename, page_1_text, master_pi_list):
    """
    Extracts PI from Page 1 or filename and fuzzy-matches against MASTER PIs 
    enforcing a strict >= 90% match threshold.
    """
    raw_pi_candidate = None

    # 1. Check Page 1 text for PI header block (e.g. PI: Aazhang, Behnaam)
    pi_text_match = re.search(r'\bPI\b.*?\n(?:[^\n]+\n){0,3}?([A-Za-z\-]+,\s*[A-Za-z\-]+)', page_1_text, re.IGNORECASE)
    if pi_text_match:
        parts = pi_text_match.group(1).split(',')
        raw_pi_candidate = f"{parts[1].strip()} {parts[0].strip()}"

    # 2. Check Filename fallback (e.g. Aazhang_Behnaam_...)
    if not raw_pi_candidate:
        fn_parts = filename.split('_')
        if len(fn_parts) >= 2 and not fn_parts[0].isdigit() and not re.match(r'^\d{2}-\d{4}', fn_parts[0]):
            raw_pi_candidate = f"{fn_parts[1].strip()} {fn_parts[0].strip()}"

    if not raw_pi_candidate:
        return "UNKNOWN"

    # 3. Fuzzy match against canonical MASTER PIs (Strict Threshold >= 0.90)
    best_match = None
    best_score = 0.0

    for master_pi in master_pi_list:
        ratio = difflib.SequenceMatcher(None, raw_pi_candidate.lower(), master_pi.lower()).ratio()
        parts = raw_pi_candidate.split()
        if len(parts) >= 2:
            swapped = f"{parts[1]} {parts[0]}".lower()
            ratio_swapped = difflib.SequenceMatcher(None, swapped, master_pi.lower()).ratio()
            ratio = max(ratio, ratio_swapped)

        if ratio > best_score:
            best_score = ratio
            best_match = master_pi

    # Require >= 90% similarity score
    if best_score >= 0.90:
        return best_match
    else:
        return raw_pi_candidate

def resolve_legacy_identifiers(filename, full_text):
    """
    Chronological multi-tier evaluator for legacy/unindexed identifiers.
    """
    # 1. Search for explicit 'A'-prefixed Project Codes (AYY-XXXX)
    proj_code_matches = re.findall(
        r'(?<![A-Za-z0-9])(A\d{2}-\d{3,4}(?:-\d{2,4})?)(?![A-Za-z0-9])', 
        f"{filename} {full_text}", re.I
    )
    
    # 2. Search for all bare YY-XXXX identifiers (excluding those inside AYY-XXXX)
    bare_matches = re.findall(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])', f"{filename} {full_text}")
    unique_bare = list(dict.fromkeys(bare_matches))

    # --- TIER 1: Explicit 'A' Prefix Present ---
    if proj_code_matches:
        found_proj = proj_code_matches[0].upper()
        found_prop = unique_bare[0] if unique_bare else "N/A"
        return found_proj, found_prop, "Legacy/Historical Project", f"Extracted Project Code ({found_proj}) and Proposal ID ({found_prop})"

    # --- TIER 2: Multiple Bare YY-XXXX Codes Found (Chronological Sequence Split) ---
    if len(unique_bare) >= 2:
        parsed = []
        for code in unique_bare:
            yy, xxxx = code.split('-')
            parsed.append((int(yy), int(xxxx), code))
        
        # Sort ascending: lower (Year, Sequence) comes first
        parsed_sorted = sorted(parsed, key=lambda x: (x[0], x[1]))
        
        lower_prop = parsed_sorted[0][2]   # Proposal ID (submitted earlier)
        higher_proj = parsed_sorted[-1][2]  # Legacy Project ID (awarded later)
        
        return higher_proj, lower_prop, "Legacy/Historical Split", f"Chronological Sequence Split: Proposal ({lower_prop}) < Project ({higher_proj})"

    # --- TIER 3: Single Bare Identifier Found ---
    if len(unique_bare) == 1:
        cand_prop = unique_bare[0]
        return "NOT_IN_MASTER", cand_prop, "Legacy/Historical Proposal", f"Extracted Single Proposal Identifier ({cand_prop})"

    return "NOT_IN_MASTER", "N/A", "Unmatched", "No Cayuse identifiers found"

def precision_progressive_match(filename, page_texts, proj_map, prop_map, oracle_map, fed_map):
    """
    Accuracy-First Progressive Matching Engine with Alphanumeric Oracle Support.
    """
    full_text = " ".join(page_texts)
    pass1_target = f"{filename} " + " ".join(page_texts[:5])

    # --- PASS 1: Active MASTER Match (Front Pages + Filename) ---
    # 1. Project ID candidate (AYY-XXXX)
    proj_cands = re.findall(r'(?<![A-Za-z0-9])(A\d{2}-\d{3,4}(?:-\d{2,4})?)(?![A-Za-z0-9])', pass1_target, re.I)
    for pc in proj_cands:
        clean_pc = pc.upper()
        if clean_pc in proj_map:
            return proj_map[clean_pc], "High", f"Matched Active CAYUSE_PROJECT_NUMBER ({clean_pc}) [Pass 1]", None, None, None
        base_pc = clean_pc.rsplit('-', 1)[0] if clean_pc.count('-') > 1 else clean_pc
        if base_pc in proj_map:
            return proj_map[base_pc], "High", f"Matched Active CAYUSE_PROJECT_NUMBER ({base_pc}) [Pass 1]", None, None, None

    # 2. Proposal ID candidate (YY-XXXX)
    prop_cands = re.findall(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])', pass1_target)
    for prc in prop_cands:
        if prc in prop_map:
            return prop_map[prc], "High", f"Matched Active CAYUSE_PROPOSAL_NUMBER ({prc}) [Pass 1]", None, None, None

    # 3. Oracle Award ID candidate (6-digit numeric OR R-prefixed e.g. R66720, 137033)
    orc_cands = re.findall(r'(?<![A-Za-z0-9])(1\d{5}|R\d[A-Za-z0-9]{4,6}(?:-[A-Za-z0-9]+)?)(?![A-Za-z0-9])', pass1_target, re.I)
    for oc in orc_cands:
        clean_oc = oc.upper()
        if clean_oc in oracle_map:
            return oracle_map[clean_oc], "Medium-High", f"Matched Active ORACLE_AWARD_NUMBER ({clean_oc}) [Pass 1]", None, None, None

    # 4. Federal Award ID
    for fed_key, row in fed_map.items():
        if fed_key in pass1_target.upper():
            return row, "Medium", f"Matched Active FEDERAL_AWARD_IDENTIFIER ({fed_key}) [Pass 1]", None, None, None

    # --- PASS 2: Active MASTER Deep Scan (Pages 6+) ---
    if len(page_texts) > 5:
        deep_target = " ".join(page_texts[5:])
        deep_proj = re.findall(r'(?<![A-Za-z0-9])(A\d{2}-\d{3,4}(?:-\d{2,4})?)(?![A-Za-z0-9])', deep_target, re.I)
        for pc in deep_proj:
            clean_pc = pc.upper()
            if clean_pc in proj_map:
                return proj_map[clean_pc], "Low-Medium", f"Matched Active CAYUSE_PROJECT_NUMBER ({clean_pc}) [Pass 2]", None, None, None

        deep_prop = re.findall(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])', deep_target)
        for prc in deep_prop:
            if prc in prop_map:
                return prop_map[prc], "Low-Medium", f"Matched Active CAYUSE_PROPOSAL_NUMBER ({prc}) [Pass 2]", None, None, None

    # --- PASS 3: LEGACY / HISTORICAL EXTRACTION ENGINE ---
    leg_oracle = orc_cands[0].upper() if orc_cands else "NOT_IN_MASTER"
    leg_proj, leg_prop, leg_conf, leg_reason = resolve_legacy_identifiers(filename, full_text)

    return None, leg_conf, leg_reason, leg_proj, leg_prop, leg_oracle

# ---------------------------------------------------------------------------
# 3. WORKER FUNCTION FOR SINGLE PDF
# ---------------------------------------------------------------------------
def process_single_pdf(pdf_path, proj_map, prop_map, oracle_map, fed_map, aln_lookup, master_pi_list):
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

        # High-Accuracy Matching Engine
        match_data, confidence, reason, leg_proj, leg_prop, leg_oracle = precision_progressive_match(
            pdf_path.name, page_texts, proj_map, prop_map, oracle_map, fed_map
        )

        aln_num, aln_title, aln_source = extract_aln_details(pdf_path.name, page_texts, aln_lookup)

        # Assign Key Identifiers
        if match_data:
            cayuse_proj_no = clean_str_id(match_data.get('CAYUSE_PROJECT_NUMBER', 'UNMATCHED'))
            cayuse_prop_no = clean_str_id(match_data.get('CAYUSE_PROPOSAL_NUMBER', 'N/A'))
            oracle_award_no = clean_str_id(match_data.get('ORACLE_AWARD_NUMBER', 'N/A'))
            lead_pi = clean_str_id(match_data.get('LEAD_PI', 'UNKNOWN'))
            fed_id = clean_str_id(match_data.get('FEDERAL_AWARD_IDENTIFIER', 'N/A'))
            sponsor = clean_str_id(match_data.get('CAYUSE_SPONSOR_NAME', 'N/A'))
            prefix_id = cayuse_proj_no if cayuse_proj_no != 'UNMATCHED' else f"PROP_{cayuse_prop_no}"
        else:
            cayuse_proj_no = leg_proj if leg_proj else "NOT_IN_MASTER"
            cayuse_prop_no = leg_prop if leg_prop else "N/A"
            oracle_award_no = leg_oracle if leg_oracle else "NOT_IN_MASTER"
            
            # Reconcile PI via Fuzzy Matching (Strict >= 90%)
            p1_text = page_texts[0] if page_texts else ""
            lead_pi = extract_and_reconcile_pi(pdf_path.name, p1_text, master_pi_list)
            
            fed_id = "N/A"
            sponsor = "N/A"
            
            if cayuse_proj_no != "NOT_IN_MASTER":
                prefix_id = f"HISTORICAL_PROJ_{cayuse_proj_no}"
            elif cayuse_prop_no != "N/A":
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
    print("Loading Master Triage Index, ALN Lookup & Canonical PIs...")
    proj_map, prop_map, oracle_map, fed_map, aln_lookup, master_pi_list = load_reference_data(
        TRIAGE_EXCEL_PATH, ALN_CSV_PATH
    )
    print(f"Loaded {len(master_pi_list)} Canonical PIs and {len(aln_lookup)} ALN titles.")

    pdf_files = [f for f in PDF_SOURCE_DIR.glob("*.pdf") if f.is_file()]
    print(f"Found {len(pdf_files)} top-level PDF files in {PDF_SOURCE_DIR}.")

    audit_records = []
    with ProcessPoolExecutor() as executor:
        futures = [
            executor.submit(process_single_pdf, pdf, proj_map, prop_map, oracle_map, fed_map, aln_lookup, master_pi_list) 
            for pdf in pdf_files
        ]
        for idx, future in enumerate(as_completed(futures)):
            result = future.result()
            audit_records.append(result)
            if (idx + 1) % 2500 == 0 or (idx + 1) == len(pdf_files):
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