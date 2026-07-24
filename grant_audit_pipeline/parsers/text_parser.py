import re
import fitz
import hashlib
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd

from config import (
    RE_ORACLE_NUM, RE_CAYUSE_PROJ, RE_CAYUSE_PROP, RE_BANNER_UID, RE_ALN,
    RE_EXECUTION_DATE, RE_DATE_RANGE, RE_CEILING_AMOUNT, RE_OBLIGATED_ACTION,
    RE_BUDGET_TOTAL, RE_INDIRECT_COST, BUDGET_PAGE_KEYWORDS, ENABLE_VISION_FALLBACK,
    parse_dollar_amount, parse_date, format_master_yaml_header, find_markdown_file
)
from parsers.table_finder import extract_budget_from_pdf_tables
from parsers.vision_parser import parse_budget_table_with_vision


def compute_file_hash(pdf_path: Path) -> str:
    """Computes MD5 hash for exact binary duplicate matching."""
    hasher = hashlib.md5()
    try:
        with open(pdf_path, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return ""


def strip_yaml_frontmatter(text: str) -> str:
    """Strips YAML frontmatter header (--- ... ---) to prevent header text pollution."""
    return re.sub(r'^---.*?---\s*', '', text, flags=re.DOTALL)


def extract_cayuse_year(raw_cayuse_proj: str, filename: str) -> Optional[int]:
    """Extracts the 2-digit integer year (YY) from a Cayuse Project Number."""
    if raw_cayuse_proj:
        m = re.search(r'^(\d{2})-\d{4}$', raw_cayuse_proj)
        if m:
            return int(m.group(1))
            
    m_fn = re.search(r'(?<!\d)(\d{2})-\d{4}(?!\d)', filename)
    if m_fn:
        return int(m_fn.group(1))
    return None


def extract_budget_pages_from_md(md_text: str) -> List[int]:
    """Extracts budget_pages array from Markdown YAML frontmatter header."""
    m = re.search(r'budget_pages:\s*\[(.*?)\]', md_text)
    if m:
        raw_str = m.group(1).strip()
        if raw_str:
            return [int(p.strip()) for p in raw_str.split(',') if p.strip().isdigit()]
    return []


def extract_header_date_context(text: str, max_chars: int = 5000, window: int = 200) -> List[Dict[str, str]]:
    """
    Scans the first `max_chars` of a document for date patterns and captures 
    a context window of `window` characters on either side to disambiguate date intent.
    """
    if not text or not isinstance(text, str):
        return []

    date_pattern = r'\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2},? \d{4}|\d{4}-\d{2}-\d{2})\b'
    
    header_text = text[:max_chars]
    snippets = []

    for match in re.finditer(date_pattern, header_text, re.IGNORECASE):
        start = max(0, match.start() - window)
        end = min(len(header_text), match.end() + window)
        snippet = " ".join(header_text[start:end].split())  # Clean extra whitespace/newlines
        
        snippets.append({
            "date": match.group(0),
            "context": f"...{snippet}..."
        })

    return snippets


def parse_osr_matrix_grid(body_text: str) -> Tuple[float, float, float, float]:
    """
    Parses the 4-column AWARDS grid on OSR Award Data Sheets (2014-2018)
    by reading ordered TOTAL, Direct, and Indirect arrays from body text.
    Returns: (Current Action Direct, Current Action Indirect, Current Action Total, Total Project Ceiling)
    """
    direct, indirect, total, ceiling = 0.0, 0.0, 0.0, 0.0
    
    if "AWARDS" in body_text or "Award Data Sheet" in body_text or "A. Current Action" in body_text:
        totals = re.findall(r'TOTAL:\s*([\d,]+(?:\.\d{2})?)', body_text, re.IGNORECASE)
        directs = re.findall(r'Direct\s*Costs:\s*([\d,]+(?:\.\d{2})?)', body_text, re.IGNORECASE)
        indirects = re.findall(r'Indirect\s*Costs:\s*([\d,]+(?:\.\d{2})?)', body_text, re.IGNORECASE)
        
        if totals:
            total = parse_dollar_amount(totals[0])  # Position 0: Current Action Total
            
        if len(totals) >= 4:
            ceiling = parse_dollar_amount(totals[3])  # Position 3: Total Project Amount
        elif len(totals) >= 2 and ceiling == 0.0:
            ceiling = parse_dollar_amount(totals[1])  # Fallback to Awarded to Date
            
        if directs:
            direct = parse_dollar_amount(directs[0])
            
        if indirects:
            indirect = parse_dollar_amount(indirects[0])
            
    return direct, indirect, total, ceiling


def route_era_budget_extraction(
    body_text: str, 
    yy_year: Optional[int]
) -> Tuple[float, float, float, float, str]:
    """
    Routes document body text to era-specific extraction routines based on YY year.
    Returns: (direct, indirect, total, ceiling, era_template_tag)
    """
    direct, indirect, total, ceiling = 0.0, 0.0, 0.0, 0.0
    template_tag = "GENERIC_KEY_VALUE"

    # 1. Classic OSR Era (2014 - 2018): 4-Column Award Data Sheet Matrix
    if yy_year is not None and 14 <= yy_year <= 18:
        template_tag = "OSR_Matrix_2014_2018"
        direct, indirect, total, ceiling = parse_osr_matrix_grid(body_text)

    # 2. Mid-Era & Modern Cayuse/Oracle Era (2019+): Web-Form & Header Key-Values
    elif yy_year is not None and yy_year >= 19:
        template_tag = "Cayuse_Oracle_2019_Plus"
        tot_m = re.search(
            r'(?:Action\s*Total|Current\s*Action\s*Amount|This\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', 
            body_text, re.IGNORECASE
        )
        if tot_m:
            total = parse_dollar_amount(tot_m.group(1))

    # 3. Legacy Era (2013 and older): Legacy Key-Value Patterns
    else:
        template_tag = "Legacy_Pre_2014"
        tot_m = re.search(r'Total\s*Amount\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', body_text, re.IGNORECASE)
        if tot_m:
            total = parse_dollar_amount(tot_m.group(1))

    return direct, indirect, total, ceiling, template_tag


def check_subcontract_out(filename: str, body_text: str) -> bool:
    """Strictly checks if a document represents an outgoing subcontract."""
    if re.search(r'(?<![A-Za-z0-9])(?:Sub-[A-Za-z0-9\-_]+|SubA|Subaward|Subcontract)(?![A-Za-z0-9])', filename, re.IGNORECASE):
        return True
    
    if re.search(r'SUBCONTRACT\s*OUT\s*[\r\n\s]*[☒\[X\]]\s*YES', body_text, re.IGNORECASE):
        return True
        
    return False


def process_document_pass_1(
    pdf_path: Path, 
    source_tag: str, 
    md_dir: Path, 
    aln_df: Optional[pd.DataFrame] = None
) -> Dict[str, Any]:
    filename = pdf_path.name
    file_hash = compute_file_hash(pdf_path)
    raw_md_text = ""
    md_path = find_markdown_file(pdf_path, md_dir)

    # 1. Identifier Extraction
    cay_proj_matches = RE_CAYUSE_PROJ.findall(filename)
    cay_prop_matches = RE_CAYUSE_PROP.findall(filename)
    oracle_matches   = RE_ORACLE_NUM.findall(filename)
    banner_matches   = RE_BANNER_UID.findall(filename)
    aln_matches      = RE_ALN.findall(filename)
    
    aln_number  = aln_matches[0] if aln_matches else ""
    raw_banner  = banner_matches[0].upper() if banner_matches else ""
    
    pi_match = re.match(r'^([A-Za-z]+_[A-Za-z]+)', filename)
    lead_pi  = pi_match.group(1) if pi_match else "EXTRACTED_PI"

    raw_cay_proj   = cay_proj_matches[0] if cay_proj_matches else ""
    raw_cay_prop   = cay_prop_matches[0] if cay_prop_matches else ""
    raw_oracle_num = oracle_matches[0] if oracle_matches else ""

    # Extract Cayuse Year (YY)
    cayuse_yy = extract_cayuse_year(raw_cay_proj, filename)

    # 2. Text Extraction & Page Inspection
    budget_pages: List[int] = []
    if md_path and md_path.exists():
        with open(md_path, 'r', encoding='utf-8', errors='ignore') as f:
            raw_md_text = f.read()
        budget_pages = extract_budget_pages_from_md(raw_md_text)
        body_text = strip_yaml_frontmatter(raw_md_text)
    else:
        try:
            doc = fitz.open(pdf_path)
            pages_text = []
            for pno, page in enumerate(doc, 1):
                p_text = page.get_text()
                pages_text.append(p_text)
                if any(kw in p_text.lower() for kw in BUDGET_PAGE_KEYWORDS) or "Award Data Sheet" in p_text:
                    budget_pages.append(pno)
            body_text = "\n".join(pages_text)
            doc.close()
        except Exception:
            body_text = ""

    # Contextual date snippet extraction (Runs AFTER body_text is assigned)
    date_snippets = extract_header_date_context(body_text, max_chars=5000, window=200)

    if not raw_banner and body_text:
        banner_txt = re.search(r'(?:Rice\s*Fund\s*No\.?|Rfund\s*#?)\s*[:\=]?\s*(R\d{4,6})\b', body_text, re.IGNORECASE)
        if banner_txt:
            raw_banner = banner_txt.group(1).upper()

    if not aln_number and body_text:
        cfda_txt = re.search(r'(?:CFDA|ALN)\s*Number\s*[:\=]?\s*(\d{2}\.\d{3})', body_text, re.IGNORECASE)
        if cfda_txt:
            aln_number = cfda_txt.group(1)

    # 3. Action Type & Subcontract Categorization
    is_nce_action = bool(re.search(r'\b(?:NCE|No\s*Cost\s*Extension)\b', filename, re.IGNORECASE))
    is_subcontract_out = check_subcontract_out(filename, body_text)

    action_category = "SUBCONTRACT_OUT" if is_subcontract_out else ("NO_COST_EXTENSION" if is_nce_action else "PRIME_AWARD")

    # 4. Header Financial & Date Extraction
    header_scope = body_text[:4000] if body_text else ""
    
    exec_match = RE_EXECUTION_DATE.search(header_scope)
    execution_date = parse_date(exec_match.group(1)) if exec_match else None
    
    range_match = RE_DATE_RANGE.search(header_scope)
    doc_start_date = parse_date(range_match.group(1)) if range_match else None
    doc_end_date = parse_date(range_match.group(2)) if range_match else None

    # Route Era-Specific Budget Extraction on pure body text
    era_dir, era_ind, era_tot, era_ceil, era_template_tag = route_era_budget_extraction(body_text, cayuse_yy)
    
    obligated_match = RE_OBLIGATED_ACTION.search(header_scope)
    delta_obligated = era_tot or (parse_dollar_amount(obligated_match.group(1)) if obligated_match else 0.0)
    
    ceiling_match = RE_CEILING_AMOUNT.search(header_scope)
    doc_ceiling = era_ceil or (parse_dollar_amount(ceiling_match.group(1)) if ceiling_match else 0.0)

    # 5. Budget Table Parsing & Fallbacks
    has_budget_table = len(budget_pages) > 0 or (bool(body_text) and any(kw in body_text.lower() for kw in BUDGET_PAGE_KEYWORDS))
    if has_budget_table and not budget_pages:
        budget_pages = [1]

    table_total, table_indirect, table_direct = era_tot, era_ind, era_dir
    budget_split_status = "NO_BUDGET_TABLE"
    extraction_method = "PyMuPDF_Fast"

    if has_budget_table and not is_nce_action and not is_subcontract_out:
        # Fallback 1: Text regex
        if table_total == 0.0:
            tot_match = RE_BUDGET_TOTAL.search(body_text)
            ind_match = RE_INDIRECT_COST.search(body_text)
            if tot_match:
                table_total = parse_dollar_amount(tot_match.group(1))
            if ind_match:
                table_indirect = parse_dollar_amount(ind_match.group(1))

        # Fallback 2: PyMuPDF table structure
        if table_total == 0.0:
            tb_direct, tb_indirect, tb_total = extract_budget_from_pdf_tables(pdf_path)
            if tb_total > 0:
                table_direct = tb_direct
                table_indirect = tb_indirect
                table_total = tb_total

        # Vision AI Evaluation (Triggered as Primary Handoff or Verification)
        if ENABLE_VISION_FALLBACK and budget_pages:
            v_res = parse_budget_table_with_vision(pdf_path, budget_pages[0])
            if v_res.get("VISION_EVAL_STATUS") == "PASSED" and v_res.get("VISION_PARSED_TOTAL", 0.0) > 0:
                # If PyMuPDF missed the total OR Vision provided a higher-confidence budget
                if table_total == 0.0 or v_res.get("VISION_CONFIDENCE_SCORE", 0.0) >= 0.8:
                    table_direct = v_res["VISION_PARSED_DIRECT"]
                    table_indirect = v_res["VISION_PARSED_INDIRECT"]
                    table_total = v_res["VISION_PARSED_TOTAL"]
                    v_exec_date = v_res.get("VISION_EXECUTION_DATE")
                    v_budget_start = v_res.get("VISION_BUDGET_START")
                    v_budget_end = v_res.get("VISION_BUDGET_END")
                    v_proj_start = v_res.get("VISION_PROJECT_START")
                    v_proj_end = v_res.get("VISION_PROJECT_END")
                    if v_res["VISION_PARSED_CEILING"] > 0 and doc_ceiling == 0.0:
                        doc_ceiling = v_res["VISION_PARSED_CEILING"]
                    extraction_method = "Multimodal_Vision_AI"
                    

        if table_total > 0:
            if table_direct == 0.0:
                table_direct = max(0.0, table_total - table_indirect)
            
            if delta_obligated == 0.0 or (delta_obligated < 1000 and table_total >= 10000) or (delta_obligated / table_total < 0.01):
                delta_obligated = table_total
                budget_split_status = "FALLBACK_TABLE_TOTAL"
            else:
                budget_split_status = "MATCHED_HEADER" if abs(table_total - delta_obligated) < 1.0 else "PARTIAL_OR_UNMATCHED"
    elif is_nce_action:
        budget_split_status = "NO_COST_EXTENSION"
    elif is_subcontract_out:
        budget_split_status = "SUBCONTRACT_OUT"

    action_tag = "NCE" if is_nce_action else ("Subcontract" if is_subcontract_out else "Standard Award")

    master_metadata = {
        "original_filename": filename,
        "file_hash": file_hash,
        "source_tag": source_tag,
        "raw_oracle_number": raw_oracle_num,
        "raw_cayuse_project": raw_cay_proj,
        "raw_cayuse_proposal": raw_cay_prop,
        "banner_award_uid": raw_banner,
        "cayuse_yy": cayuse_yy,
        "era_template_tag": era_template_tag,
        "lead_pi": lead_pi,
        "action_tag": action_tag,
        "action_category": action_category,
        "aln_number": aln_number,
        "execution_date": execution_date.strftime("%Y-%m-%d") if execution_date else "",
        "doc_start_date": doc_start_date.strftime("%Y-%m-%d") if doc_start_date else "",
        "doc_end_date": doc_end_date.strftime("%Y-%m-%d") if doc_end_date else "",
        "delta_obligated": delta_obligated,
        "doc_ceiling": doc_ceiling,
        "has_budget_table": has_budget_table,
        "budget_pages": budget_pages,
        "budget_split_status": budget_split_status,
        "table_total": table_total,
        "table_direct": table_direct,
        "table_indirect": table_indirect,
        "extraction_method": extraction_method,
        "date_snippets": date_snippets
    }

    if not (md_path and md_path.exists()):
        md_dir.mkdir(parents=True, exist_ok=True)
        target_md_path = md_dir / f"{pdf_path.stem}.md"
        header_str = format_master_yaml_header(master_metadata)
        
        with open(target_md_path, "w", encoding="utf-8") as f:
            f.write(f"{header_str}# Document: {filename}\n\n{body_text}")
        md_path = target_md_path

    return {
        "PDF_Path": str(pdf_path),
        "Markdown_Path": str(md_path) if md_path else "",
        "Filename": filename,
        "FILE_HASH": file_hash,
        "SOURCE_TAG": source_tag,
        "RAW_CAYUSE_PROJ": raw_cay_proj,
        "RAW_CAYUSE_PROP": raw_cay_prop,
        "RAW_ORACLE_NUM": raw_oracle_num,
        "RAW_BANNER_UID": raw_banner,
        "CAYUSE_YY": cayuse_yy,
        "ERA_TEMPLATE_TAG": era_template_tag,
        "ACTION_CATEGORY": action_category,
        "EXECUTION_DATE": execution_date,
        "DOC_START_DATE": doc_start_date,
        "DOC_END_DATE": doc_end_date,
        "DELTA_OBLIGATED": delta_obligated,
        "DOC_CEILING": doc_ceiling,
        "HAS_BUDGET_TABLE": has_budget_table,
        "TABLE_TOTAL": table_total,
        "TABLE_DIRECT": table_direct,
        "TABLE_INDIRECT": table_indirect,
        "BUDGET_SPLIT_STATUS": budget_split_status,
        "EXTRACTION_METHOD": extraction_method,
        "LEAD_PI": lead_pi,
        "ALN_NUMBER": aln_number,
        "TEXT_BODY": body_text,
        "BUDGET_PAGES": budget_pages,
        "DATE_SNIPPETS": date_snippets
    }