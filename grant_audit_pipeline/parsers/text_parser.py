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
from parsers.ai_reader import parse_ai_header_payload

# Regex patterns for De-obligation & Administrative Action filtering
RE_DEOBLIGATION_PATTERNS = re.compile(
    r'(?:^|[-_.\s])(?:de-?obligat(?:ion|ed|e)?|de-?ob|funding[-_\s]+reduction|reduction[-_\s]+of[-_\s]+funds|decrease[-_\s]+award|net[-_\s]+reduction|de-?commit(?:ment)?)(?:$|[-_.\s])|'
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
    
    # 1. Outgoing Subcontracts (Non-binding transfers from institutional perspective)
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
    if RE_ADMIN_MOD_PATTERNS.search(fn):
        return ("ADMINISTRATIVE_MODIFICATION", False)

    # 8. No-Cost Extensions (Non-financial period extensions)
    if re.search(r'\b(?:NCE|No\s*Cost\s*Extension)\b', fn, re.IGNORECASE) or re.search(r'NO[-_\s]*COST[-_\s]*EXTENSION', body_text[:2000], re.IGNORECASE):
        return ("NO_COST_EXTENSION", False)
        
    # 9. Official Notice of Award / Sponsor Notice
    return ("OFFICIAL_NOA", True)


def detect_sponsor_template(filename: str, body_text: str = "", aln_number: str = "") -> str:
    """
    SPONSOR CLASSIFIER REGISTRY:
    Auto-detects sponsor template taxonomy upfront to route to specific pattern matchers.
    """
    fn = str(filename).lower()
    text = body_text[:4000].lower() if body_text else ""

    # 1. NIH / PHS (ALN 93.xxx or explicit NIH/PHS markers)
    if "nih" in fn or "phs" in text or "national institutes of health" in text or "department of health and human services" in text or aln_number.startswith("93"):
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
    STAGE 3: SPONSOR-SPECIFIC FEATURE VECTOR EXTRACTION (REGEX BACKUP)
    Extracts up to 4 financial anchors into a normalized vector:
    [ACTION_AMOUNT, PERIOD_AMOUNT, CUMULATIVE_AMOUNT, PROJECT_CEILING]
    """
    text = body_text[:5000] if body_text else ""
    
    a_action = 0.0
    b_period = 0.0
    c_cum = 0.0
    m_ceiling = 0.0

    if template_type == "NIH_PHS":
        m_20 = re.search(r'(?:20\.\s*Total\s*Amount\s*of\s*Federal\s*Funds\s*Obligated\s*by\s*this\s*Action|Obligated\s*by\s*this\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_20:
            a_action = parse_dollar_amount(m_20.group(1))

        m_23 = re.search(r'(?:23\.\s*Total\s*Amount\s*of\s*Federal\s*Funds\s*Obligated\s*this\s*budget\s*period|Obligated\s*this\s*budget\s*period)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_23:
            b_period = parse_dollar_amount(m_23.group(1))

        m_27 = re.search(r'(?:27\.\s*Total\s*Amount\s*of\s*the\s*Federal\s*Award|including\s*Approved\s*Cost\s*Sharing\s*this\s*Project\s*Period)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
        if m_27:
            c_cum = parse_dollar_amount(m_27.group(1))

    elif template_type == "SUBCONTRACT_IN":
        a_action = (table_direct + table_indirect) if (table_direct + table_indirect) > 0 else table_direct
        if a_action == 0.0:
            m_sub_tot = re.search(r'(?:Subaward\s*Total|Total\s*Subaward\s*Amount|Amount\s*Funded|Direct\s*Costs?)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE)
            if m_sub_tot:
                a_action = parse_dollar_amount(m_sub_tot.group(1))

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

    # 1. Identifier Extraction (Filename & Raw Text Base)
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

    # 2. Text Body Extraction
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

    # 3. SECONDARY VERIFICATION PASS: PyMuPDF / Regex Classification & Extraction
    regex_category, regex_is_binding = classify_document_type(filename, body_text)
    regex_template = detect_sponsor_template(filename, body_text, aln_number)

    table_total, table_indirect, table_direct = 0.0, 0.0, 0.0
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

    regex_fv = extract_sponsor_feature_vector(
        body_text, regex_template, table_direct, table_indirect, table_total
    )
    regex_action = regex_fv.get("ACTION_AMOUNT", 0.0)

    # 4. PRIMARY PASS: Multimodal AI Reader (4-Page Ingestion & Local Cache)
    ai_payload = parse_ai_header_payload(pdf_path, max_pages=4)

    extraction_method = "Multimodal_AI_Reader"
    hitl_review_required = False
    discrepancy_note = ""

    if ai_payload and ai_payload.get("EVAL_STATUS") in ("PASSED", "PASSED_CACHE_HIT"):
        ai_meta = ai_payload.get("DOCUMENT_METADATA") or {}
        ai_dates = ai_payload.get("DATES") or {}
        ai_fin = ai_payload.get("FINANCIAL_VECTOR") or {}
        ai_conf = ai_payload.get("CONFIDENCE_SCORE", 0.95)

        # Primary Assignments from AI JSON
        sponsor_template = ai_meta.get("detected_sponsor_template") or regex_template
        is_official = bool(ai_meta.get("is_official_sponsor_notice", True))
        is_admin = bool(ai_meta.get("is_administrative_non_financial_action", False))

        if is_admin:
            action_category = "ADMINISTRATIVE_MODIFICATION"
            is_binding_financial_action = False
        elif not is_official:
            action_category = regex_category if regex_category != "OFFICIAL_NOA" else "TECHNICAL_NARRATIVE"
            is_binding_financial_action = False
        else:
            action_category = regex_category if regex_category not in ("OFFICIAL_NOA", "ADMINISTRATIVE_MODIFICATION") else "OFFICIAL_NOA"
            is_binding_financial_action = True

        execution_date = parse_date(ai_dates.get("execution_date")) or parse_date(snippet_dates['exec_date'])
        doc_start_date = parse_date(ai_dates.get("budget_period_start")) or parse_date(snippet_dates['start_date'])
        doc_end_date   = parse_date(ai_dates.get("budget_period_end")) or parse_date(snippet_dates['end_date'])

        a_action = parse_dollar_amount(ai_fin.get("action_obligation_amount"))
        b_period = parse_dollar_amount(ai_fin.get("current_budget_period_total"))
        c_cum    = parse_dollar_amount(ai_fin.get("cumulative_obligated_to_date"))
        m_ceil   = parse_dollar_amount(ai_fin.get("total_project_ceiling"))
        sub_dir  = parse_dollar_amount(ai_fin.get("local_subrecipient_direct_cost"))
        sub_ind  = parse_dollar_amount(ai_fin.get("local_subrecipient_indirect_cost"))

        # Subaward Pass-Through Re-anchoring
        if sponsor_template == "SUBCONTRACT_IN" and (sub_dir + sub_ind) > 0:
            a_action = sub_dir + sub_ind
            table_direct, table_indirect = sub_dir, sub_ind

        feature_vector = {
            "ACTION_AMOUNT": a_action,
            "PERIOD_AMOUNT": b_period,
            "CUMULATIVE_AMOUNT": c_cum,
            "PROJECT_CEILING": m_ceil
        }

        # Dual-Source Delta Discrepancy & HITL Resolution Check
        delta_diff = abs(a_action - regex_action)
        if a_action > 0 and regex_action > 0 and delta_diff > 1000.0 and ai_conf < 0.80:
            hitl_review_required = True
            discrepancy_note = f"HITL_REVIEW_REQUIRED: Dual-Source Delta Discrepancy (AI=${a_action:,.2f} vs Regex=${regex_action:,.2f})"

    else:
        # Fallback to PyMuPDF / Regex if AI Reader is unconfigured or failed
        extraction_method = "PyMuPDF_Regex_Fallback"
        sponsor_template = regex_template
        action_category = regex_category
        is_binding_financial_action = regex_is_binding
        feature_vector = regex_fv
        execution_date = parse_date(snippet_dates['exec_date'])
        doc_start_date = parse_date(snippet_dates['start_date'])
        doc_end_date   = parse_date(snippet_dates['end_date'])

    # 5. MULTI-SIGNAL CATEGORY VALIDATION CONTRACTS
    extracted_delta = feature_vector["ACTION_AMOUNT"]
    doc_ceiling     = feature_vector["PROJECT_CEILING"]

    # CONTRACT RULE 1: Admin Mod / NCE Re-Classification Override
    if action_category in ("ADMINISTRATIVE_MODIFICATION", "NO_COST_EXTENSION") and extracted_delta > 0.0:
        action_category = "OFFICIAL_NOA"
        is_binding_financial_action = True
        budget_split_status = "UPGRADED_FROM_ADMIN_MOD_TO_NOA"
    else:
        budget_split_status = "NORMAL_PARSED"

    # CONTRACT RULE 2: De-obligation Negative Delta Normalization
    if action_category == "DEOBLIGATION":
        is_binding_financial_action = True
        if extracted_delta > 0.0:
            extracted_delta = -abs(extracted_delta)
            feature_vector["ACTION_AMOUNT"] = extracted_delta
        budget_split_status = "ENFORCED_DEOBLIGATION_NEGATIVE"

    # CONTRACT RULE 3: Strict Outgoing Subcontracts Isolation (Non-Binding Transfers)
    if action_category == "SUBCONTRACT_OUT":
        is_binding_financial_action = False
        budget_split_status = "SUBCONTRACT_OUT_NON_BINDING"

    # CONTRACT RULE 4: Strict NCE Non-Financial Enforcement & Binding Column Segregation
    if action_category == "NO_COST_EXTENSION":
        non_binding_reported_budget = extracted_delta or feature_vector["CUMULATIVE_AMOUNT"] or table_total
        delta_obligated = 0.00
        doc_ceiling = 0.00
        budget_split_status = "NCE_NON_FINANCIAL_RESTATED_BUDGET"
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
        "has_budget_table": len(budget_pages) > 0,
        "budget_pages": budget_pages,
        "budget_split_status": budget_split_status,
        "table_total": table_total,
        "table_direct": table_direct,
        "table_indirect": table_indirect,
        "extraction_method": extraction_method,
        "hitl_review_required": hitl_review_required,
        "discrepancy_note": discrepancy_note
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
        "HAS_BUDGET_TABLE": len(budget_pages) > 0,
        "TABLE_TOTAL": table_total,
        "TABLE_DIRECT": table_direct,
        "TABLE_INDIRECT": table_indirect,
        "BUDGET_SPLIT_STATUS": budget_split_status,
        "EXTRACTION_METHOD": extraction_method,
        "HITL_REVIEW_REQUIRED": hitl_review_required,
        "DISCREPANCY_NOTE": discrepancy_note,
        "LEAD_PI": lead_pi,
        "ALN_NUMBER": aln_number,
        "TEXT_BODY": body_text,
        "BUDGET_PAGES": budget_pages
    }