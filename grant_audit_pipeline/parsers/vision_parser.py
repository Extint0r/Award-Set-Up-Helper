import os
import json
import fitz  # PyMuPDF
from pathlib import Path
from typing import Dict, Any, Optional

VISION_CACHE_DIR = Path(r"D:\0-Batch-AWARDS\processed_files\Review\Vision_Cache")

# Import configuration constants and sanitizers
from config import VISION_CACHE_DIR, parse_dollar_amount, ENABLE_VISION_FALLBACK

# SDK Check for Google GenAI (New SDK)
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False


def render_page_to_image(pdf_path: Path, page_number: int, target_dpi: int = 200) -> Path:
    """Renders a PDF page to a high-resolution PNG for Vision AI inspection."""
    VISION_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    out_img_path = VISION_CACHE_DIR / f"{pdf_path.stem}_p{page_number}.png"
    
    if out_img_path.exists():
        return out_img_path

    doc = fitz.open(pdf_path)
    page = doc.load_page(page_number - 1)  # 0-indexed in PyMuPDF

    zoom = target_dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    
    print(f"Processing image {out_img_path.name} (Dimensions: {pix.width}x{pix.height}px)...")
    pix.save(str(out_img_path))
    doc.close()
    
    return out_img_path


def parse_budget_table_with_vision(
    pdf_path: Path, 
    page_number: int, 
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    Targeted Multimodal Vision Parser using Gemini Flash via the new google-genai SDK.
    Explicitly handles NIH Notice of Award boxes (Box 20, 20a, 20b, 27) and generic award tables.
    Includes instant local disk caching and socket timeout safeguards.
    """
    fallback_result = {
        "EXTRACTION_METHOD": "Fallback_Rules",
        "VISION_EVAL_STATUS": "SKIPPED",
        "VISION_CONFIDENCE_SCORE": 0.0,
        "VISION_PARSED_DIRECT": 0.0,
        "VISION_PARSED_INDIRECT": 0.0,
        "VISION_PARSED_TOTAL": 0.0,
        "VISION_PARSED_CEILING": 0.0,
        "VISION_EXECUTION_DATE": None,
        "VISION_BUDGET_START": None,
        "VISION_BUDGET_END": None,
        "VISION_PROJECT_START": None,
        "VISION_PROJECT_END": None,
        "VISION_NOTES": "Vision fallback disabled or API unavailable."
    }

    if not ENABLE_VISION_FALLBACK:
        return fallback_result

    effective_key = api_key or os.getenv("GEMINI_API_KEY") or os.getenv("VISION_API_KEY")
    if not effective_key:
        fallback_result["VISION_EVAL_STATUS"] = "FAILED"
        fallback_result["VISION_NOTES"] = "Missing GEMINI_API_KEY or VISION_API_KEY in environment/.env file."
        return fallback_result

    if not GENAI_AVAILABLE:
        fallback_result["VISION_EVAL_STATUS"] = "FAILED"
        fallback_result["VISION_NOTES"] = "google-genai SDK not installed (run `pip install google-genai`)."
        return fallback_result

    try:
        # 1. Render budget page to image (or return existing cached image path)
        img_path = render_page_to_image(pdf_path, page_number)
        json_cache_path = img_path.with_suffix(".json")

        # 2. FAST PATH: Check if audit JSON already exists locally before hitting the network
        if json_cache_path.exists():
            print(f"[Cache Hit] Loading audit JSON directly from disk for {pdf_path.name}")
            with open(json_cache_path, "r", encoding="utf-8") as jf:
                res_data = json.load(jf)
        else:
            # 3. Initialize modern GenAI client
            client = genai.Client(api_key=effective_key)

            # 4. Prompt specifically tailored for NIH Notice of Award headers & standard budget tables
            prompt = """
        You are an expert research administration auditor inspecting an award document page image.

        1. DATE CLASSIFICATION (Look at header text, Section I, or signature blocks):
           - execution_date: The official issue date or sponsor signature date for THIS action (e.g., "Award Date", "Issue Date").
           - budget_period_start: Start date for THIS active budget period (e.g., NIH Box 19 Start Date).
           - budget_period_end: End date for THIS active budget period (e.g., NIH Box 19 End Date).
           - project_period_start: Overall cumulative project start date (e.g., NIH Box 26 Start Date).
           - project_period_end: Overall cumulative project max end date (e.g., NIH Box 26 End Date).

        2. FINANCIAL BREAKDOWN FOR THIS SPECIFIC ACTION:
           - direct_cost: Direct costs for THIS action (numeric float or null).
           - indirect_cost: Indirect / F&A costs for THIS action (numeric float or null).
           - total_action_amount: Total obligated by THIS action (numeric float or null, e.g., NIH Box 20).

        3. CUMULATIVE AWARD / PROJECT CEILING:
           - cumulative_award_to_date: Total running federal funds obligated to date (numeric float or null, e.g., NIH Box 27).
           - total_project_ceiling: Total anticipated multi-year project ceiling if listed.

        Respond STRICTLY with JSON matching this structure:
        {
          "execution_date": null,
          "budget_period_start": null,
          "budget_period_end": null,
          "project_period_start": null,
          "project_period_end": null,
          "direct_cost": null,
          "indirect_cost": null,
          "total_action_amount": null,
          "cumulative_award_to_date": null,
          "total_project_ceiling": null,
          "confidence_score": 0.0,
          "notes": ""
        }
        """

            # 5. Upload file and query model with a 30-second timeout safeguard
            uploaded_file = client.files.upload(file=img_path)
            
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[uploaded_file, prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    http_options={'timeout': 30000}  # 30-second timeout safeguard
                )
            )

            res_data = json.loads(response.text)

            # Save local JSON cache alongside the PNG image for auditing
            with open(json_cache_path, "w", encoding="utf-8") as jf:
                json.dump(res_data, jf, indent=2)

            # Clean up temporary uploaded file from remote storage
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass

        # 6. Parse JSON data & sanitize outputs
        direct = parse_dollar_amount(res_data.get("direct_cost", 0.0))
        indirect = parse_dollar_amount(res_data.get("indirect_cost", 0.0))
        total = parse_dollar_amount(res_data.get("total_action_amount", 0.0))
        
        # Cumulative/Ceiling resolution
        cum_to_date = parse_dollar_amount(res_data.get("cumulative_award_to_date", 0.0))
        proj_ceiling = parse_dollar_amount(res_data.get("total_project_ceiling", 0.0))
        ceiling = cum_to_date if cum_to_date > 0 else proj_ceiling

        confidence = float(res_data.get("confidence_score", 0.0))
        notes = str(res_data.get("notes", ""))

        return {
            "EXTRACTION_METHOD": "Multimodal_Vision_AI",
            "VISION_EVAL_STATUS": "PASSED",
            "VISION_CONFIDENCE_SCORE": confidence,
            "VISION_PARSED_DIRECT": direct,
            "VISION_PARSED_INDIRECT": indirect,
            "VISION_PARSED_TOTAL": total,
            "VISION_PARSED_CEILING": ceiling,
            "VISION_EXECUTION_DATE": res_data.get("execution_date"),
            "VISION_BUDGET_START": res_data.get("budget_period_start"),
            "VISION_BUDGET_END": res_data.get("budget_period_end"),
            "VISION_PROJECT_START": res_data.get("project_period_start"),
            "VISION_PROJECT_END": res_data.get("project_period_end"),
            "VISION_NOTES": notes
        }

    except Exception as e:
        return {
            "EXTRACTION_METHOD": "Multimodal_Vision_AI",
            "VISION_EVAL_STATUS": "FAILED",
            "VISION_CONFIDENCE_SCORE": 0.0,
            "VISION_PARSED_DIRECT": 0.0,
            "VISION_PARSED_INDIRECT": 0.0,
            "VISION_PARSED_TOTAL": 0.0,
            "VISION_PARSED_CEILING": 0.0,
            "VISION_EXECUTION_DATE": None,
            "VISION_BUDGET_START": None,
            "VISION_BUDGET_END": None,
            "VISION_PROJECT_START": None,
            "VISION_PROJECT_END": None,
            "VISION_NOTES": f"Vision API execution error: {str(e)}"
        }