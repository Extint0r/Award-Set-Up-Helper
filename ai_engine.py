import os
import json
import re
import io
import pandas as pd
from google import genai
from google.genai import types as genai_types
from core_parser import load_all_registries
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("Critical Configuration Deficit: 'GEMINI_API_KEY' was not found.")

client = genai.Client(api_key=API_KEY)


def clean_binary_csv_local(filename):
    """Executes the mandatory 5-step binary parsing protocol on local reference registries."""
    if not os.path.exists(filename):
        return pd.DataFrame()
    raw_bytes = open(filename, 'rb').read()
    bom_index = raw_bytes.find(b'\xef\xbb\xbf')
    if bom_index != -1:
        cleaned_bytes = raw_bytes[bom_index:]
        text_stream = cleaned_bytes.decode('utf-8')
    else:
        text_stream = raw_bytes.decode('utf-8', errors='ignore')
    df = pd.read_csv(io.StringIO(text_stream))
    df.columns = [c.strip() for c in df.columns]
    return df


def clean_json_text(text):
    """Safely removes markdown tag fences from the output payload stream."""
    clean = text.strip()
    markdown_fence_mask = "\x60\x60\x60"
    if clean.startswith(markdown_fence_mask):
        clean = re.sub(r'^' + markdown_fence_mask + r'(?:json)?\s*', '', clean)
        clean = re.sub(r'\s*' + markdown_fence_mask + r'$', '', clean)
    return clean.strip()


## Define the Standalone Single-Document Ingestion Schema with Safe Arrays
DATA_SCHEMA = genai_types.Schema(
    type=genai_types.Type.OBJECT,
    properties={
        # Standalone File Structural Identifiers
        "parent_proposal_number": genai_types.Schema(type=genai_types.Type.STRING, description="Format: YY-XXXX Unique Anchor Key"),
        "sponsor_award_number": genai_types.Schema(type=genai_types.Type.STRING),
        "modification_number": genai_types.Schema(type=genai_types.Type.STRING, description="e.g., Base Award, Mod 01, Amendment 2, NCE"),
        "document_execution_date": genai_types.Schema(type=genai_types.Type.STRING, description="The signature or execution date in YYYY-MM-DD format"),
        
        # Core Contract Attributes Stated on the Page Face
        "award_name": genai_types.Schema(type=genai_types.Type.STRING, description="Format: [Funder Acronym. Sponsor Award Number. Award Acronym. Lead PI Family Name]"),
        "primary_sponsor": genai_types.Schema(type=genai_types.Type.STRING, description="Legal agreement partner validated against Sponsors.csv"),
        "start_date": genai_types.Schema(type=genai_types.Type.STRING, description="YYYY-MM-DD"),
        "end_date": genai_types.Schema(type=genai_types.Type.STRING, description="YYYY-MM-DD Project Period End Date established or modified by this specific file"),
        "principal_investigator": genai_types.Schema(type=genai_types.Type.STRING, description="Lead PI name aligned with Investigators.csv"),
        "award_owning_organization": genai_types.Schema(type=genai_types.Type.STRING, description="Format: XXXXX-Department Name from Rice University Orgs.csv"),
        
        # System Enums Bound to Destination Oracle/Grants Module Constraints
        "award_purpose": genai_types.Schema(type=genai_types.Type.STRING, enum=["Research", "Instruction", "Training", "Public Service", "Construction", "Student Aid", "Agency Funds", "Other", "Undetermined"]),
        "award_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Federal Gov", "State/Local Gov", "State/Local Gov - Fed Prime", "Business", "Business - Fed Prime", "Non-Profit Org", "Non-Profit Org - Fed Prime", "Non-Profit Org - State/Local Prime", "Higher Ed", "Higher Ed - Fed Prime", "Other", "Undetermined"]),
        "bill_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["CRM", "CRO", "CRQ", "CRY", "FBM", "FBO", "FBQ", "FBY", "SCHM", "SCHO", "SCHP", "SCHQ", "SCHY", "Other", "Undetermined"]),
        "flow_through_sponsor": genai_types.Schema(type=genai_types.Type.STRING, enum=["Yes", "No", "Other", "Undetermined"]),
        "originating_sponsor_and_prime_num": genai_types.Schema(type=genai_types.Type.STRING, description="Format: [Funder Name/Agency Acronym]: [Prime Award Number]"),
        "aln_number": genai_types.Schema(type=genai_types.Type.STRING, description="Format: XX.XXX Code"),
        "fa_rate": genai_types.Schema(type=genai_types.Type.STRING, description="Numerical percentage and base, e.g., 61% MTDC"),
        "close_date_days": genai_types.Schema(type=genai_types.Type.STRING, description="Number of closeout days as a string envelope, e.g., '90'"),
        "funding_mechanism": genai_types.Schema(type=genai_types.Type.STRING, enum=["Contract", "Cooperative Agreement", "Grant", "OTA", "Other", "Undetermined"]),
        
        "subject_to_terms_and_conditions": genai_types.Schema(
            type=genai_types.Type.ARRAY, 
            items=genai_types.Schema(type=genai_types.Type.STRING),
            description="List of all applicable regulatory terms selected from the protocol instructions"
        ),
        "category_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Applied Research", "Basic Research", "Experimental Development", "Non Research", "Other", "Undetermined"]),
        
        "far_clauses": genai_types.Schema(
            type=genai_types.Type.ARRAY, 
            items=genai_types.Schema(type=genai_types.Type.STRING),
            description="List of all applicable FAR clauses selected from the protocol instructions"
        ),
        
        "field_of_science_fos": genai_types.Schema(type=genai_types.Type.STRING, description="Extracted categories totaling 100%"),
        "special_interest_si": genai_types.Schema(type=genai_types.Type.STRING),
        "special_interest_thecb": genai_types.Schema(type=genai_types.Type.STRING),
        "subject_to_single_audit": genai_types.Schema(type=genai_types.Type.STRING),
        
        # ──────────────────────────────────────────────────────────────────
        # FINANCIAL REALIGNMENT PASS: CONTEXTUAL RESEARCH SETUP GUIDES
        # ──────────────────────────────────────────────────────────────────
        "direct_funding_delta": genai_types.Schema(
            type=genai_types.Type.STRING, 
            description="The incremental change (+/-) in direct costs introduced solely by this individual file action. Look for labels like 'Direct Costs' or 'Current Action Direct'."
        ),
        "indirect_funding_delta": genai_types.Schema(
            type=genai_types.Type.STRING, 
            description="The incremental change (+/-) in indirect/F&A costs introduced solely by this individual file action. Look for labels like 'Indirect Costs' or 'Current Action F&A'."
        ),
        "total_funding_delta": genai_types.Schema(
            type=genai_types.Type.STRING, 
            description="The explicit incremental total funding amount obligated or authorized by THIS individual notice action. Look for labels like 'Amount of This Action', 'Total Federal Share Obligated', 'Obligated Amount', 'This Action Total', or 'Funding Delta'. For No-Cost Extensions (NCE), always return '0.0'."
        ),
        "total_awarded_ceiling": genai_types.Schema(
            type=genai_types.Type.STRING, 
            description="The overall maximum cumulative anticipated or approved project ceiling value stated on the face of THIS specific notice page. Look for labels like 'Total Project Period Approved Amount', 'Cumulative Estimated Total Cost', 'Anticipated Project Total', or 'Ceiling Envelope'. If unstated, return 'N/A'."
        ),
        
        # Multi-Dimensional Arrays
        "deliverables": genai_types.Schema(
            type=genai_types.Type.ARRAY,
            items=genai_types.Schema(
                type=genai_types.Type.OBJECT,
                properties={
                    "name": genai_types.Schema(type=genai_types.Type.STRING),
                    "report_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Quarterly", "Semi-Annual", "Annual", "Final", "Other", "Undetermined"]),
                    "due_date": genai_types.Schema(type=genai_types.Type.STRING),
                    "description": genai_types.Schema(type=genai_types.Type.STRING)
                },
                required=["name", "report_type", "due_date", "description"]
            )
        ),
        "budget_ledger": genai_types.Schema(
            type=genai_types.Type.ARRAY,
            items=genai_types.Schema(
                type=genai_types.Type.OBJECT,
                properties={
                    "period": genai_types.Schema(type=genai_types.Type.STRING),
                    "investigator": genai_types.Schema(type=genai_types.Type.STRING),
                    "category": genai_types.Schema(type=genai_types.Type.STRING, enum=["Equipment", "F&A Cost", "Fringe Benefits", "Other Direct Costs", "Salaries & Wages", "Subawards", "Travel", "Tuition Remission", "Other", "Undetermined"]),
                    "amount": genai_types.Schema(type=genai_types.Type.STRING)
                },
                required=["period", "investigator", "category", "amount"]
            )
        )
    },
    # 🌟 LOOSENED VALIDATION GATE: Reduced to bare minimum fields to allow polymorphic parameter extraction
    required=["parent_proposal_number", "total_funding_delta"]
)


def extract_award_data(pdf_path, protocol_path="1.71.md"):
    """Ingests behavior registries and runs standalone document mapping serialization passes."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Target document missing: {pdf_path}")
        
    with open(protocol_path, "r", encoding="utf-8") as f:
        raw_rules = f.read()
    system_instruction_rules = raw_rules.split("Final Output Styling & Layout Architecture")[0] if "Final Output Styling & Layout Architecture" in raw_rules else raw_rules

    system_instruction_rules += (
        "\n\nSINGLE-FILE EXTRACTION FOCUS: You are analyzing ONE individual contract document from a timeline history. "
        "Extract only the financial deltas, overall stated awarded ceiling, target dates, and deliverables introduced specifically by this single attached file.\n"
        "HIGH-RIGOR CERTAINTY PROTOCOL: Cross-reference the extracted 'aln_number' hierarchically against the provided ALN PROGRAM HIERARCHY dataset. "
        "Use the specific sub-program numerical codes to confidently resolve 'award_purpose' and 'category_type' strings based on federal agency mandates. "
        "If the ALN code or text indicators are ambiguous, missing, or unstated, you are ORDERED to select the 'Other' or 'Undetermined' enum choice. "
        "Do not guess or leave fields empty.\n"
        "CRITICAL JSON STRING FORMATTING RULE: Every string field value must be a single continuous line of text. "
        "Do NOT output raw literal line breaks or physical newlines inside any string value. "
        "Semicolon-separate internal items on a single line. Escape inner double quotes as \\\" or convert them to single quotes."
    )

    sponsors_df, investigators_df, orgs_df = load_all_registries(data_folder="data")
    aln_hierarchy_df = clean_binary_csv_local("aln_hierarchy.csv")
    
    uploaded_file = client.files.upload(file=pdf_path)
    
    reference_payload = f"""
    Process the attached standalone document according to Protocol 1.71 single-file extraction constraints.
    Cross-reference your extraction strings programmatically against the literal text entries provided below.
    --- REFERENCE REGISTER: SPONSORS ---
    {sponsors_df.to_csv(index=False)}
    --- REFERENCE REGISTER: INVESTIGATORS ---
    {investigators_df.to_csv(index=False)}
    --- REFERENCE REGISTER: RICE ORGS ---
    {orgs_df.to_csv(index=False)}
    --- REFERENCE REGISTER: ALN PROGRAM HIERARCHY ---
    {aln_hierarchy_df.to_csv(index=False) if not aln_hierarchy_df.empty else 'No local ALN program csv registry currently staged on disk.'}
    """
    
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[uploaded_file, reference_payload],
        config=genai_types.GenerateContentConfig(
            system_instruction=system_instruction_rules,
            response_mime_type="application/json",
            response_schema=DATA_SCHEMA,
            temperature=0.1,
            max_output_tokens=8192
        ),
    )
    client.files.delete(name=uploaded_file.name)
    
    cleaned_json_str = clean_json_text(response.text)
    
    # ──────────────────────────────────────────────────────────────────
    # TWO-PASS AUTO-REPAIR GATEWAY ARCHITECTURE
    # ──────────────────────────────────────────────────────────────────
    try:
        return json.loads(cleaned_json_str)
    except json.JSONDecodeError as decode_error:
        print(f"⚠️ Formatting anomaly detected ({decode_error.msg}). Activating automated AI auto-repair gateway...")
        try:
            repair_prompt = f"""
            The following structured JSON payload possesses a formatting or text-boundary defect (such as an unterminated string quote, an unescaped character, or a missing trailing bracket) that prevents standard JSON parsers from initializing it.
            
            Review the payload string text, isolate the syntax fracture, and output a completely pristine, flawless version of the identical JSON data structure. 
            Do NOT truncate the data, summarize fields, or delete items from arrays. Keep all extracted information fully intact, but ensure all string fields are valid single lines.
            
            Return ONLY the clean raw JSON text payload without markdown fences.
            
            --- MALFORMED JSON PAYLOAD ---
            {cleaned_json_str}
            """
            repair_response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=repair_prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction="You are an expert JSON lint repair utility. Output pure, valid JSON code strings only.",
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            repaired_text = clean_json_text(repair_response.text)
            return json.loads(repaired_text)
        except Exception as repair_failure:
            print(f"❌ Auto-repair gateway exhausted: {repair_failure}")
            raise decode_error