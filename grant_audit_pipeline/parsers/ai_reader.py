import os
import json
import re
import fitz  # PyMuPDF
from pathlib import Path
from typing import Dict, Any, List, Optional

from config import VISION_CACHE_DIR, parse_dollar_amount, ENABLE_VISION_FALLBACK, VISION_API_KEY

# SDK Check for Google GenAI (New SDK)
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


def render_pdf_pages_to_images(pdf_path: Path, max_pages: int = 4, target_dpi: int = 200) -> List[Path]:
    """
    Renders up to `max_pages` (Pages 1 to 4) of a PDF to high-resolution PNG images 
    for Multimodal AI inspection.
    """
    VISION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    rendered_image_paths: List[Path] = []

    try:
        doc = fitz.open(pdf_path)
        pages_to_render = min(len(doc), max_pages)

        for p_idx in range(pages_to_render):
            page_num = p_idx + 1
            out_img_path = VISION_CACHE_DIR / f"{pdf_path.stem}_p{page_num}.png"

            if not out_img_path.exists():
                page = doc.load_page(p_idx)
                zoom = target_dpi / 72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                pix.save(str(out_img_path))

            rendered_image_paths.append(out_img_path)

        doc.close()
    except Exception as e:
        print(f"[Warning] Error rendering pages for {pdf_path.name}: {e}")

    return rendered_image_paths


def sanitize_ai_identifiers(idents: dict) -> dict:
    """
    Post-processing defense-in-depth sanitization:
    - Oracle award number must be strictly 6 numeric digits.
    - Cayuse project number must strictly follow YY-XXXX format (auto-formatting 6-digit variants if needed).
    - Eliminates status words, internal vendor IDs, or misallocated strings.
    """
    if not isinstance(idents, dict):
        return {"oracle_award_number": None, "cayuse_project_number": None, "aln_cfda_number": None, "principal_investigator": None}

    orc = str(idents.get("oracle_award_number") or "").strip()
    cay = str(idents.get("cayuse_project_number") or "").strip()

    # Strict Oracle Validation: Exactly 6 digits
    clean_orc = orc if re.match(r'^[1-9]\d{5}$', orc) else None

    # Strict Cayuse Validation: YY-XXXX format
    clean_cay = None
    if re.match(r'^\d{2}-\d{4}$', cay):
        clean_cay = cay
    elif re.match(r'^\d{6}$', cay):
        clean_cay = f"{cay[:2]}-{cay[2:]}"

    return {
        "oracle_award_number": clean_orc,
        "cayuse_project_number": clean_cay,
        "aln_cfda_number": idents.get("aln_cfda_number"),
        "principal_investigator": idents.get("principal_investigator")
    }


def parse_ai_header_payload(
    pdf_path: Path, 
    max_pages: int = 4, 
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    STEPS 1 & 2: Hardened Multimodal AI Reader Engine (4-Page Ingestion)
    
    Renders pages 1 through 4 of raw PDFs to PNG images and submits them to Gemini Flash.
    Extracts structured audit payloads with strict format boundaries and negative constraints
    on identifiers, eliminating cross-contamination and status value leaks.
    """
    json_cache_path = VISION_CACHE_DIR / f"{pdf_path.stem}_header.json"

    fallback_result = {
        "EXTRACTION_METHOD": "Fallback_Rules",
        "EVAL_STATUS": "SKIPPED",
        "CONFIDENCE_SCORE": 0.0,
        "DOCUMENT_METADATA": {
            "detected_sponsor_template": "GENERIC_NOA",
            "document_title_header": "",
            "is_official_sponsor_notice": True,
            "is_administrative_non_financial_action": False,
            "pages_inspected": 0
        },
        "IDENTIFIERS": {
            "oracle_award_number": None,
            "cayuse_project_number": None,
            "aln_cfda_number": None,
            "principal_investigator": None
        },
        "DATES": {
            "execution_date": None,
            "budget_period_start": None,
            "budget_period_end": None,
            "project_period_start": None,
            "project_period_end": None
        },
        "FINANCIAL_VECTOR": {
            "action_obligation_amount": 0.0,
            "current_budget_period_total": 0.0,
            "cumulative_obligated_to_date": 0.0,
            "total_project_ceiling": 0.0,
            "local_subrecipient_direct_cost": 0.0,
            "local_subrecipient_indirect_cost": 0.0
        },
        "EXTRACTION_NOTES": "Vision AI disabled or unconfigured."
    }

    if not ENABLE_VISION_FALLBACK:
        return fallback_result

    # 1. FAST PATH: Return cached JSON from disk if already processed
    if json_cache_path.exists():
        try:
            with open(json_cache_path, "r", encoding="utf-8") as jf:
                cached_data = json.load(jf)
            cached_data["EVAL_STATUS"] = "PASSED_CACHE_HIT"
            return cached_data
        except Exception:
            pass

    effective_key = api_key or VISION_API_KEY or os.getenv("GEMINI_API_KEY") or os.getenv("VISION_API_KEY")
    if not effective_key:
        fallback_result["EVAL_STATUS"] = "FAILED"
        fallback_result["EXTRACTION_NOTES"] = "Missing GEMINI_API_KEY or VISION_API_KEY in environment/.env file."
        return fallback_result

    if not GENAI_AVAILABLE:
        fallback_result["EVAL_STATUS"] = "FAILED"
        fallback_result["EXTRACTION_NOTES"] = "google-genai SDK not installed (run `pip install google-genai`)."
        return fallback_result

    # 2. Render up to `max_pages` PNG images
    img_paths = render_pdf_pages_to_images(pdf_path, max_pages=max_pages)
    if not img_paths:
        fallback_result["EVAL_STATUS"] = "FAILED"
        fallback_result["EXTRACTION_NOTES"] = "Failed to render PDF pages to images."
        return fallback_result

    uploaded_remote_files = []
    try:
        client = genai.Client(api_key=effective_key)

        # 3. Upload rendered images to Google GenAI storage
        for path in img_paths:
            uf = client.files.upload(file=path)
            uploaded_remote_files.append(uf)

        # 4. Hardened 4-Page Auditor Prompt with Strict Negative Constraints & Format Boundaries
        prompt = f"""
You are an expert research administration auditor inspecting up to {len(img_paths)} sequential page image(s) from a grant award PDF.

CRITICAL EXTRACTION RULES & NEGATIVE CONSTRAINTS:
1. ORACLE AWARD NUMBER: Must be STRICTLY 6 numeric digits (e.g., 238025, 166016). NEVER extract Cayuse project IDs, subcontract numbers, status strings (such as "Active"), or vendor supplier codes here. If a true 6-digit Oracle ID is not explicitly printed, return null.
2. CAYUSE PROJECT NUMBER: Must strictly follow the format YY-XXXX where YY is a 2-digit year and XXXX is a 4-digit sequence (e.g., 22-0996, 26-0802). NEVER extract vendor ID strings, internal subaward tracking codes, or prime sponsor grant numbers here.
3. FINANCIAL VECTORS: If this is a non-financial action (such as an NCE without funds, a PI change, or a technical progress report), ALL financial vector amounts MUST be explicitly returned as null, NOT 0.0.

Inspect the document pages (headers, notice title blocks, financial summary boxes, and budget tables) and extract the following:

1. DOCUMENT CLASSIFICATION & METADATA:
   - detected_sponsor_template: Choose the BEST match from: ["NIH_PHS", "NSF_STANDARD", "DOE_HQ", "DOD_GENERIC", "SUBCONTRACT_IN", "OSR_INTERNAL", "GENERIC_NOA"].
     * Use "SUBCONTRACT_IN" if this is an incoming subaward from another university/institution (e.g. UCSF, BCM, UTHSC).
     * Use "NIH_PHS" for direct NIH Notices of Award (Box 20/23/27).
     * Use "NSF_STANDARD" for direct National Science Foundation Award Notices / Amendments.
   - document_title_header: The exact primary title or heading (e.g., "Notice of Award", "Subaward Agreement", "Amendment 002").
   - is_official_sponsor_notice: true if this is a binding sponsor Notice of Award or Subaward contract; false if progress report, technical proposal, or internal draft.
   - is_administrative_non_financial_action: true if this is strictly a non-financial admin change (PI change, address change, NCE without funds); false if it obligates or restates funds.

2. IDENTIFIERS:
   - oracle_award_number: 6 numeric digits or null.
   - cayuse_project_number: YY-XXXX format or null.
   - aln_cfda_number: Federal CFDA / ALN Number if listed (e.g., 93.286, 47.070).
   - principal_investigator: Lead Principal Investigator name.

3. DATES (Format strictly as YYYY-MM-DD or null):
   - execution_date: Official award date / notice issue date for THIS action.
   - budget_period_start: Start date for THIS active budget period.
   - budget_period_end: End date for THIS active budget period.
   - project_period_start: Overall project start date.
   - project_period_end: Overall project max end date.

4. FINANCIAL VECTOR (Extract exact numeric dollar values or null):
   - action_obligation_amount: Incremental dollar obligation by THIS action (e.g. NIH Box 20, NSF "Amount of This Action", Subaward "Funded This Action").
   - current_budget_period_total: Approved total for current budget period (e.g. NIH Box 23).
   - cumulative_obligated_to_date: Total running funds obligated to date across all periods (e.g. NIH Box 27, NSF "Cumulative Obligated Amount").
   - total_project_ceiling: Total anticipated multi-year project ceiling or estimated project total.
   - local_subrecipient_direct_cost: Local subrecipient direct cost for this action (if subaward pass-through).
   - local_subrecipient_indirect_cost: Local subrecipient indirect/F&A cost for this action (if subaward pass-through).

5. CONFIDENCE METRICS:
   - overall_score: Confidence float from 0.0 to 1.0.
   - financial_confidence: Confidence float specifically in the extracted numbers from 0.0 to 1.0.
   - extraction_notes: Brief note on which page/box the main financial numbers were found.

Respond STRICTLY with a valid JSON object matching this structure:
{{
  "document_metadata": {{
    "detected_sponsor_template": "NIH_PHS",
    "document_title_header": "Notice of Award",
    "is_official_sponsor_notice": true,
    "is_administrative_non_financial_action": false,
    "pages_inspected": {len(img_paths)}
  }},
  "identifiers": {{
    "oracle_award_number": null,
    "cayuse_project_number": null,
    "aln_cfda_number": null,
    "principal_investigator": null
  }},
  "dates": {{
    "execution_date": null,
    "budget_period_start": null,
    "budget_period_end": null,
    "project_period_start": null,
    "project_period_end": null
  }},
  "financial_vector": {{
    "action_obligation_amount": null,
    "current_budget_period_total": null,
    "cumulative_obligated_to_date": null,
    "total_project_ceiling": null,
    "local_subrecipient_direct_cost": null,
    "local_subrecipient_indirect_cost": null
  }},
  "ai_confidence": {{
    "overall_score": 0.0,
    "financial_confidence": 0.0,
    "extraction_notes": ""
  }}
}}
"""

        # 5. Execute Gemini Flash query with 45-second timeout safeguard
        contents = uploaded_remote_files + [prompt]
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                http_options={'timeout': 45000}
            )
        )

        raw_json_dict = json.loads(response.text)

        # 6. Sanitize Identifiers via Defense-in-Depth Programmatic Guardrail
        raw_idents = raw_json_dict.get("identifiers") or {}
        sanitized_idents = sanitize_ai_identifiers(raw_idents)

        # 7. Normalize and sanitize JSON fields
        fin = raw_json_dict.get("financial_vector") or {}
        a_act = parse_dollar_amount(fin.get("action_obligation_amount"))
        b_tot = parse_dollar_amount(fin.get("current_budget_period_total"))
        c_cum = parse_dollar_amount(fin.get("cumulative_obligated_to_date"))
        m_ceil = parse_dollar_amount(fin.get("total_project_ceiling"))
        sub_dir = parse_dollar_amount(fin.get("local_subrecipient_direct_cost"))
        sub_ind = parse_dollar_amount(fin.get("local_subrecipient_indirect_cost"))

        conf = raw_json_dict.get("ai_confidence") or {}
        overall_conf = float(conf.get("overall_score", 0.95))
        fin_conf = float(conf.get("financial_confidence", 0.95))

        normalized_payload = {
            "EXTRACTION_METHOD": "Multimodal_AI_Reader",
            "EVAL_STATUS": "PASSED",
            "CONFIDENCE_SCORE": overall_conf,
            "FINANCIAL_CONFIDENCE_SCORE": fin_conf,
            "DOCUMENT_METADATA": raw_json_dict.get("document_metadata") or {},
            "IDENTIFIERS": sanitized_idents,
            "DATES": raw_json_dict.get("dates") or {},
            "FINANCIAL_VECTOR": {
                "action_obligation_amount": a_act,
                "current_budget_period_total": b_tot,
                "cumulative_obligated_to_date": c_cum,
                "total_project_ceiling": m_ceil,
                "local_subrecipient_direct_cost": sub_dir,
                "local_subrecipient_indirect_cost": sub_ind
            },
            "EXTRACTION_NOTES": conf.get("extraction_notes", "")
        }

        # Save to disk cache
        with open(json_cache_path, "w", encoding="utf-8") as jf:
            json.dump(normalized_payload, jf, indent=2)

        return normalized_payload

    except Exception as e:
        fallback_result["EVAL_STATUS"] = "FAILED"
        fallback_result["EXTRACTION_NOTES"] = f"AI Reader execution error: {str(e)}"
        return fallback_result

    finally:
        # Clean up temporary remote files
        for uf in uploaded_remote_files:
            try:
                client.files.delete(name=uf.name)
            except Exception:
                pass