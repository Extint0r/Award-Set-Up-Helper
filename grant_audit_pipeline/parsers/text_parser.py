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

# Regex patterns for De-obligation & Administrative Action filtering
RE_DEOBLIGATION_PATTERNS = re.compile(
    r'\b(?:de-?obligat(?:ion|ed|e)|funding\s+reduction|reduction\s+of\s+funds|decrease\s+award|net\s+reduction|de-?commit(?:ment)?)\b|'
    r'\(\s*\$\s*[\d,]+(?:\.\d{2})?\s*\)',
    re.IGNORECASE
)

RE_ADMIN_MOD_PATTERNS = re.compile(
    r'\b(?:pi\s+change|change\s+of\s+pi|principal\s+investigator\s+change|'
    r'key\s+personnel|administrative\s+amendment|administrative\s+modification|'
    r'scope\s+change|change\s+in\s+scope|institutional\s+name\s+change|'
    r'address\s+change|no-cost\s+amendment)\b',
    re.IGNORECASE
)


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

    date_pattern = r'\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}|\d{4}-\d{2}-\d{2})\b'
    
    header_text = text[:max_chars]
    snippets = []

    for match in re.finditer(date_pattern, header_text, re.IGNORECASE):
        start = max(0, match.start() - window)
        end = min(len(header_text), match.end() + window)
        snippet = " ".join(header_text[start:end].split())
        
        snippets.append({
            "date": match.group(0),
            "context": f"...{snippet}..."
        })

    return snippets


def parse_dates_from_snippets(snippets: List[Dict[str, str]]) -> Dict[str, Optional[str]]:
    """
    Parses start, end, and execution dates from context snippets 
    with label-anchored regex to prevent label mismatch.
    """
    if not snippets:
        return {'start_date': None, 'end_date': None, 'exec_date': None}

    full_snippet_text = " ".join([s.get("context", "") for s in snippets if isinstance(s, dict)])
    date_pat = r'\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}|\d{4}-\d{2}-\d{2})\b'
    
    extracted = {'start_date': None, 'end_date': None, 'exec_date': None}

    budget_match = re.search(
        r'(?:Budget\s*Period|Current\s*Budget)\b.{0,60}?(?P<start>' + date_pat + r').{0,30}?(?:to|through|-|–|End\s*Date)\s*(?P<end>' + date_pat + r')', 
        full_snippet_text, re.IGNORECASE
    )
    if budget_match:
        extracted['start_date'] = budget_match.group('start')
        extracted['end_date'] = budget_match.group('end')

    exec_match = re.search(
        r'(?:Award\s*Date|Issue\s*Date|Date\s*Issued|Federal\s*Award\s*Date)\b.{0,30}?(?P<exec>' + date_pat + r')', 
        full_snippet_text, re.IGNORECASE
    )
    if exec_match:
        extracted['exec_date'] = exec_match.group('exec')

    return extracted


def check_subcontract_out(filename: str, body_text: str = "") -> bool:
    """Strictly checks if a document represents an outgoing subcontract."""
    if re.search(r'(?<![A-Za-z0-9])(?:Sub-[A-Za-z0-9\-_]+|SubA|Subaward|Subcontract)(?![A-Za-z0-9])', filename, re.IGNORECASE):
        return True
    
    if re.search(r'SUBCONTRACT\s*OUT.{0,100}?[☒\[X\]]\s*YES', body_text, re.DOTALL | re.IGNORECASE):
        return True
        
    return False


def classify_document_type(filename: str, body_text: str = "") -> Tuple[str, bool]:
    """
    STAGE 2: IMMUTABLE CLASSIFICATION ENGINE
    Classifies documents into granular categories and determines binding status.
    This category is immutable and will not be mutated downstream based on parsed values.
    """
    fn = str(filename)
    header_text = body_text[:3000]
    
    # 1. Outgoing Subcontracts
    if check_subcontract_out(fn, body_text):
        return ("SUBCONTRACT_OUT", False)
        
    # 2. Internal Advance Spend / Pre-Award Risk Authorizations
    if re.search(r'advspd|advance\s*spend|pre-?award', fn, re.IGNORECASE):
        return ("INTERNAL_ADVANCE_SPEND", False)

    # 3. Recipient Progress Reports (RPPR / Progress Report)
    if re.search(r'(?:^|[-_.\s])PR(?:[-_.\s]|$)|RPPR|ProgressReport', fn, re.IGNORECASE) or re.search(r'\bRPPR\b|Research\s+Performance\s+Progress\s+Report', body_text[:2000], re.IGNORECASE):
        return ("RECIPIENT_PROGRESS_REPORT", False)
        
    # 4. Informal Rebudgets, Proposals, Internal Drafts
    if re.search(r'no-notice|revbud|draft|internal|work-in-progress', fn, re.IGNORECASE):
        return ("PROPOSAL_OR_REVISED_BUDGET", False)
        
    # 5. Technical Narratives
    if re.search(r'-tech|Tech-chgs', fn, re.IGNORECASE):
        return ("TECHNICAL_NARRATIVE", False)

    # 6. De-obligation Actions (Binding Financial Action with Negative Funding)
    if RE_DEOBLIGATION_PATTERNS.search(fn) or RE_DEOBLIGATION_PATTERNS.search(header_text):
        return ("DEOBLIGATION", True)

    # 7. Non-Financial Administrative Modifications
    # Restrict RE_ADMIN_MOD_PATTERNS strictly to filenames to prevent body text false positives
    if RE_ADMIN_MOD_PATTERNS.search(fn):
        return ("ADMINISTRATIVE_MODIFICATION", False)

    # 8. No-Cost Extensions (Non-financial period extensions)
    if re.search(r'\b(?:NCE|No\s*Cost\s*Extension)\b', fn, re.IGNORECASE) or re.search(r'NO[-_\s]*COST[-_\s]*EXTENSION', body_text[:2000], re.IGNORECASE):
        return ("NO_COST_EXTENSION", False)
        
    # 9. Official Notice of Award / Sponsor Notice
    return ("OFFICIAL_NOA", True)


def detect_sponsor_template(filename: str, body_text: str = "", aln_number: str = "") -> str:
    """
    SPONSOR CLASSIFIER REGISTRY (REFINED):
    Auto-detects sponsor template taxonomy upfront to route to specific pattern matchers.
    
    CRITICAL FIX: Agency-First Priority.
    Checks ALN numbers and explicit agency markers BEFORE checking subaward fallbacks.
    Removes 'amd' keyword over-matching (amd = amendment, NOT subaward).
    """
    fn = str(filename).lower()
    text = body_text[:4000].lower() if body_text else ""

    # 1. NIH / PHS (ALN 93.xxx or explicit NIH/PHS markers)
    if "nih" in fn or "phs" in text or "national institutes of health" in text or "department of health and human services" in text or aln_number.startswith("93"):
        # Check if it's a pass-through subaward from another university/institution under NIH
        if any(k in fn for k in ["ucsf", "bcm", "uthsc", "uc-sf", "subaward_in"]):
            return "SUBCONTRACT_IN"
        return "NIH_PHS"

    # 2. NSF (ALN 47.xxx or explicit NSF markers)
    if "nsf" in fn or "national science foundation" in text or aln_number.startswith("47"):
        if any(k in fn for k in ["ucsf", "bcm", "uthsc", "uc-sf", "subaward_in"]):
            return "SUBCONTRACT_IN"
        return "NSF_STANDARD"

    # 3. DOE (ALN 81.xxx or explicit DOE markers)
    if "doe" in fn or "department of energy" in text or "assistance agreement" in text or aln_number.startswith("81"):
        if any(k in fn for k in ["ucsf", "bcm", "uthsc", "uc-sf", "subaward_in"]):
            return "SUBCONTRACT_IN"
        return "DOE_HQ"

    # 4. DOD / Military
    if any(k in fn for k in ["dod", "onr", "afosr", "darpa", "army", "navy"]) or "dd1155" in text or "sf30" in text:
        if any(k in fn for k in ["ucsf", "bcm", "uthsc", "uc-sf", "subaward_in"]):
            return "SUBCONTRACT_IN"
        return "DOD_GENERIC"

    # 5. Incoming Subawards (Pass-Through from other universities / prime recipients)
    if any(k in fn for k in ["ucsf", "bcm", "uthsc", "uc-sf", "subaward_in"]) or "subrecipient agreement" in text or "pass-through entity" in text:
        return "SUBCONTRACT_IN"

    # 6. Internal OSR Data Sheet (Fallback if no federal sponsor detected)
    if "award data sheet" in text or "a. current action" in text or "osr" in fn:
        return "OSR_INTERNAL"

    return "GENERIC_NOA"


def extract_sponsor_feature_vector(
    body_text: str, 
    template_type: str, 
    table_direct: float = 0.0, 
    table_indirect: float = 0.0, 
    table_total: float = 0.0
) -> Dict[str, float]:
    """
    STAGE 3: SPONSOR-SPECIFIC FEATURE VECTOR EXTRACTION
    Extracts up to 4 financial anchors into a normalized vector:
    [ACTION_AMOUNT, PERIOD_AMOUNT, CUMULATIVE_AMOUNT, PROJECT_CEILING]
    """
    text = body_text[:5000] if body_text else ""
    
    a_action = 0.0
    b_period = 0.0
    c_cum = 0.0
    m_ceiling = 0.0

    if template_type == "NIH_PHS":
        # Box 20: Action Obligation
        m_20 = re.search(r'(?:20\.\s*Total\s*Amount\s*of\s*Federal\s*Funds\s*Obligated\s*by\s*this\s*Action|Obligated\s*by\s*this\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_20:
            a_action = parse_dollar_amount(m_20.group(1))

        # Box 23: Budget Period Obligation
        m_23 = re.search(r'(?:23\.\s*Total\s*Amount\s*of\s*Federal\s*Funds\s*Obligated\s*this\s*budget\s*period|Obligated\s*this\s*budget\s*period)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_23:
            b_period = parse_dollar_amount(m_23.group(1))

        # Box 27: Stated Cumulative Award Total to Date
        m_27 = re.search(r'(?:27\.\s*Total\s*Amount\s*of\s*the\s*Federal\s*Award|including\s*Approved\s*Cost\s*Sharing\s*this\s*Project\s*Period)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_27:
            c_cum = parse_dollar_amount(m_27.group(1))

    elif template_type == "SUBCONTRACT_IN":
        # Enhanced Subaward re-anchoring: Prioritize Subrecipient Direct + Indirect Local Allocation
        a_action = (table_direct + table_indirect) if (table_direct + table_indirect) > 0 else table_direct
        if a_action == 0.0:
            m_sub_tot = re.search(r'(?:Subaward\s*Total|Total\s*Subaward\s*Amount|Amount\s*Funded|Direct\s*Costs?)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
            if m_sub_tot:
                a_action = parse_dollar_amount(m_sub_tot.group(1))

        # Capture prime grant total as ceiling reference
        m_prime = re.search(r'(?:Prime\s*Award|Grant\ Total|Total\ Federal\ Award|Total\ Estimated\ Cost)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_prime:
            m_ceiling = parse_dollar_amount(m_prime.group(1))

    elif template_type == "NSF_STANDARD":
        m_action = re.search(r'(?:Amount\s*of\s*This\s*Action|Obligated\s*Amount\s*This\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_action:
            a_action = parse_dollar_amount(m_action.group(1))

        m_cum = re.search(r'(?:Cumulative\s*Obligated\s*Amount|Total\s*Awarded\s*to\s*Date)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_cum:
            c_cum = parse_dollar_amount(m_cum.group(1))

        m_ceil = re.search(r'(?:Total\s*Expected\s*Award\s*Amount|Total\s*Estimated\s*Cost)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_ceil:
            m_ceiling = parse_dollar_amount(m_ceil.group(1))

    elif template_type == "DOE_HQ":
        m_action = re.search(r'(?:13\.\s*Action\s*Obligation|Action\s*Obligation)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_action:
            a_action = parse_dollar_amount(m_action.group(1))

        m_cum = re.search(r'(?:14\.\s*Total\s*Obligated\s*Funds|Total\s*Obligated\s*Funds)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_cum:
            c_cum = parse_dollar_amount(m_cum.group(1))

        m_ceil = re.search(r'(?:15\.\s*Total\s*Estimated\s*Cost|Total\s*Estimated\s*Cost)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_ceil:
            m_ceiling = parse_dollar_amount(m_ceil.group(1))

    elif template_type == "OSR_INTERNAL":
        totals = re.findall(r'TOTAL:\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if totals:
            a_action = parse_dollar_amount(totals[0])
        if len(totals) >= 3:
            c_cum = parse_dollar_amount(totals[2])
        if len(totals) >= 4:
            m_ceiling = parse_dollar_amount(totals[3])

    else: # GENERIC_NOA fallback
        m_action = RE_OBLIGATED_ACTION.search(text)
        if m_action:
            a_action = parse_dollar_amount(m_action.group(1))

        m_ceil = RE_CEILING_AMOUNT.search(text)
        if m_ceil:
            m_ceiling = parse_dollar_amount(m_ceil.group(1))

    return {
        "ACTION_AMOUNT": a_action,
        "PERIOD_AMOUNT": b_period,
        "CUMULATIVE_AMOUNT": c_cum,
        "PROJECT_CEILING": m_ceiling
    }


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

    date_snippets = extract_header_date_context(body_text, max_chars=5000, window=200)
    snippet_dates = parse_dates_from_snippets(date_snippets)

    if not raw_banner and body_text:
        banner_txt = re.search(r'(?:Rice\s*Fund\s*No\.?|Rfund\s*#?)\s*[:\=]?\s*(R\d{4,6})\b', body_text, re.IGNORECASE)
        if banner_txt:
            raw_banner = banner_txt.group(1).upper()

    if not aln_number and body_text:
        cfda_txt = re.search(r'(?:CFDA|ALN)\s*Number\s*[:\=]?\s*(\d{2}\.\d{3})', body_text, re.IGNORECASE)
        if cfda_txt:
            aln_number = cfda_txt.group(1)

    # 3. STAGE 2 IMMUTABLE Classification & Sponsor Auto-Detection
    action_category, is_binding_financial_action = classify_document_type(filename, body_text)
    sponsor_template = detect_sponsor_template(filename, body_text, aln_number)

    # 4. Header Financial & Date Extraction
    header_scope = body_text[:4000] if body_text else ""
    
    exec_match = RE_EXECUTION_DATE.search(header_scope)
    execution_date = parse_date(exec_match.group(1)) if exec_match else None
    
    range_match = RE_DATE_RANGE.search(header_scope)
    doc_start_date = parse_date(range_match.group(1)) if range_match else None
    doc_end_date = parse_date(range_match.group(2)) if range_match else None

    if not doc_start_date and snippet_dates['start_date']:
        doc_start_date = parse_date(snippet_dates['start_date'])
    if not doc_end_date and snippet_dates['end_date']:
        doc_end_date = parse_date(snippet_dates['end_date'])
    if not execution_date and snippet_dates['exec_date']:
        execution_date = parse_date(snippet_dates['exec_date'])

    # 5. Budget Table Parsing
    table_total, table_indirect, table_direct = 0.0, 0.0, 0.0
    has_budget_table = len(budget_pages) > 0 or (bool(body_text) and any(kw in body_text.lower() for kw in BUDGET_PAGE_KEYWORDS))
    if has_budget_table and not budget_pages:
        budget_pages = [1]

    budget_split_status = "NO_BUDGET_TABLE"
    extraction_method = "PyMuPDF_Fast"

    v_exec_date, v_budget_start, v_budget_end, v_proj_start, v_proj_end = None, None, None, None, None

    if has_budget_table and is_binding_financial_action:
        tot_match = RE_BUDGET_TOTAL.search(body_text)
        ind_match = RE_INDIRECT_COST.search(body_text)
        if tot_match:
            table_total = parse_dollar_amount(tot_match.group(1))
        if ind_match:
            table_indirect = parse_dollar_amount(ind_match.group(1))

        if table_total == 0.0:
            tb_direct, tb_indirect, tb_total = extract_budget_from_pdf_tables(pdf_path)
            if tb_total > 0:
                table_direct, table_indirect, table_total = tb_direct, tb_indirect, tb_total

        if ENABLE_VISION_FALLBACK and budget_pages:
            target_vision_page = 1 if (1 in budget_pages or not budget_pages) else budget_pages[0]
            v_res = parse_budget_table_with_vision(pdf_path, target_vision_page)
            if v_res.get("VISION_EVAL_STATUS") == "PASSED" and v_res.get("VISION_PARSED_TOTAL", 0.0) > 0:
                if table_total == 0.0 or v_res.get("VISION_CONFIDENCE_SCORE", 0.0) >= 0.8:
                    table_direct = v_res["VISION_PARSED_DIRECT"]
                    table_indirect = v_res["VISION_PARSED_INDIRECT"]
                    table_total = v_res["VISION_PARSED_TOTAL"]
                    v_exec_date = v_res.get("VISION_EXECUTION_DATE")
                    v_budget_start = v_res.get("VISION_BUDGET_START")
                    v_budget_end = v_res.get("VISION_BUDGET_END")
                    v_proj_start = v_res.get("VISION_PROJECT_START")
                    v_proj_end = v_res.get("VISION_PROJECT_END")
                    extraction_method = "Multimodal_Vision_AI"

        if table_total > 0 and table_direct == 0.0:
            table_direct = max(0.0, table_total - table_indirect)

    # 6. STAGE 3 FEATURE VECTOR EXTRACTION
    feature_vector = extract_sponsor_feature_vector(
        body_text, sponsor_template, table_direct, table_indirect, table_total
    )

    extracted_delta = feature_vector["ACTION_AMOUNT"]
    doc_ceiling = feature_vector["PROJECT_CEILING"]

    # GUARDRAIL 1: Subaward Prime Federal Grant Ceiling Isolation
    if sponsor_template == "SUBCONTRACT_IN":
        if extracted_delta > 1000000 and 0 < table_direct < 500000:
            extracted_delta = table_direct + table_indirect if (table_direct + table_indirect) > 0 else table_direct
            budget_split_status = "REANCHORED_SUBAWARD_TOTAL"

    # GUARDRAIL 2: De-obligation Negative Delta Normalization
    if action_category == "DEOBLIGATION":
        if extracted_delta > 0.0:
            extracted_delta = -abs(extracted_delta)
        budget_split_status = "ENFORCED_DEOBLIGATION_NEGATIVE"

    # GUARDRAIL 3: Strict NCE Non-Financial Enforcement (IMMUTABLE)
    if action_category == "NO_COST_EXTENSION":
        non_binding_reported_budget = extracted_delta or feature_vector["CUMULATIVE_AMOUNT"] or table_total
        delta_obligated = 0.00
        doc_ceiling = 0.00
        budget_split_status = "NCE_NON_FINANCIAL_RESTATED_BUDGET"

    # GUARDRAIL 4: Column Segregation based on Binding Status
    elif is_binding_financial_action:
        delta_obligated = extracted_delta
        non_binding_reported_budget = 0.00
    else:
        delta_obligated = 0.00
        doc_ceiling = 0.00
        non_binding_reported_budget = extracted_delta or table_total

    master_metadata = {
        "original_filename": filename,
        "file_hash": file_hash,
        "source_tag": source_tag,
        "raw_oracle_number": raw_oracle_num,
        "raw_cayuse_project": raw_cay_proj,
        "raw_cayuse_proposal": raw_cay_prop,
        "banner_award_uid": raw_banner,
        "cayuse_yy": cayuse_yy,
        "sponsor_template": sponsor_template,
        "lead_pi": lead_pi,
        "action_category": action_category,
        "is_binding_financial_action": is_binding_financial_action,
        "aln_number": aln_number,
        "execution_date": execution_date.strftime("%Y-%m-%d") if execution_date else "",
        "doc_start_date": doc_start_date.strftime("%Y-%m-%d") if doc_start_date else "",
        "doc_end_date": doc_end_date.strftime("%Y-%m-%d") if doc_end_date else "",
        "delta_obligated": delta_obligated,
        "doc_ceiling": doc_ceiling,
        "non_binding_reported_budget": non_binding_reported_budget,
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
        "SPONSOR_TEMPLATE": sponsor_template,
        "ACTION_CATEGORY": action_category,
        "IS_BINDING_FINANCIAL_ACTION": is_binding_financial_action,
        "EXECUTION_DATE": execution_date,
        "DOC_START_DATE": doc_start_date,
        "DOC_END_DATE": doc_end_date,
        "DELTA_OBLIGATED": delta_obligated,
        "DOC_CEILING": doc_ceiling,
        "NON_BINDING_REPORTED_BUDGET": non_binding_reported_budget,
        "FEATURE_VECTOR": feature_vector,
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
        "DATE_SNIPPETS": date_snippets,
        "VISION_EXECUTION_DATE": v_exec_date,
        "VISION_BUDGET_START": v_budget_start,
        "VISION_BUDGET_END": v_budget_end,
        "VISION_PROJECT_START": v_proj_start,
        "VISION_PROJECT_END": v_proj_end
    }