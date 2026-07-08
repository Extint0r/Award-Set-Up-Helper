import os
import json
from google import genai
from google.genai import types
from core_parser import load_all_registries
from dotenv import load_dotenv # Ingest our modern environment variable loader

# 1. Automatically locate and read the hidden .env file on your hard drive
load_dotenv()

# 2. Extract the key securely from your system environment memory layer
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("Critical Configuration Deficit: 'GEMINI_API_KEY' was not found. "
                     "Please check that your local .env file contains this variable.")

# 3. Initialize the client securely
client = genai.Client(api_key=API_KEY)

##  Define the Rigid Data Schema using the modern Type structure
DATA_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        # --- Profile A: Flat Metadata (01-31) ---
        "award_name": types.Schema(type=types.Type.STRING, description="Format: [Funder Acronym. Sponsor Award Number. Award Acronym. Lead PI Family Name]"),
        "primary_sponsor": types.Schema(type=types.Type.STRING, description="Legal agreement partner validated against Sponsors.csv"),
        "start_date": types.Schema(type=types.Type.STRING, description="YYYY-MM-DD"),
        "end_date": types.Schema(type=types.Type.STRING, description="YYYY-MM-DD"),
        "principal_investigator": types.Schema(type=types.Type.STRING, description="Lead PI name aligned with Investigators.csv"),
        "award_owning_organization": types.Schema(type=types.Type.STRING, description="Format: XXXXX-Department Name from Rice University Orgs.csv"),
        "sponsor_award_number": types.Schema(type=types.Type.STRING),
        "award_purpose": types.Schema(type=types.Type.STRING, enum=["Research", "Instruction", "Training", "Public Service", "Construction", "Student Aid", "Agency Funds"]),
        "award_type": types.Schema(type=types.Type.STRING, enum=["Federal Gov", "State/Local Gov", "State/Local Gov - Fed Prime", "Business", "Business - Fed Prime", "Non-Profit Org", "Non-Profit Org - Fed Prime", "Non-Profit Org - State/Local Prime", "Higher Ed", "Higher Ed - Fed Prime"]),
        "bill_type": types.Schema(type=types.Type.STRING, enum=["CRM", "CRO", "CRQ", "CRY", "FBM", "FBO", "FBQ", "FBY", "SCHM", "SCHO", "SCHP", "SCHQ", "SCHY"]),
        "flow_through_sponsor": types.Schema(type=types.Type.STRING, enum=["Yes", "No"]),
        "originating_sponsor_and_prime_num": types.Schema(type=types.Type.STRING, description="Format: [Funder Name/Agency Acronym]: [Prime Award Number]"),
        "parent_proposal_number": types.Schema(type=types.Type.STRING, description="Format: YY-XXXX Unique Anchor Key"),
        "aln_number": types.Schema(type=types.Type.STRING, description="Format: XX.XXX"),
        "fa_rate": types.Schema(type=types.Type.STRING, description="Numerical percentage and base, e.g., 61% MTDC"),
        "close_date_days": types.Schema(type=types.Type.INTEGER, description="Number of days after end date for closeout"),
        "funding_mechanism": types.Schema(type=types.Type.STRING, enum=["Contract", "Cooperative Agreement", "Grant", "OTA"]),
        "subject_to_terms_and_conditions": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "category_type": types.Schema(type=types.Type.STRING, enum=["Applied Research", "Basic Research", "Experimental Development", "Non Research"]),
        "far_clauses": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "field_of_science_fos": types.Schema(type=types.Type.STRING, description="Extracted categories totaling 100%"),
        "special_interest_si": types.Schema(type=types.Type.STRING),
        "special_interest_thecb": types.Schema(type=types.Type.STRING),
        "subject_to_single_audit": types.Schema(type=types.Type.STRING),
        "direct_funding_obligated": types.Schema(type=types.Type.NUMBER),
        "indirect_funding_obligated": types.Schema(type=types.Type.NUMBER),
        "total_funding_obligated": types.Schema(type=types.Type.NUMBER),
        "direct_funding_total_awarded": types.Schema(type=types.Type.NUMBER),
        "indirect_funding_total_awarded": types.Schema(type=types.Type.NUMBER),
        "total_funding_total_awarded": types.Schema(type=types.Type.NUMBER),
        "letter_of_credit_payment_id": types.Schema(type=types.Type.STRING),
        
        # --- Profile B: Multi-Dimensional Child Arrays (32-36) ---
        "deliverables": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "name": types.Schema(type=types.Type.STRING),
                    "report_type": types.Schema(type=types.Type.STRING, enum=["Quarterly", "Semi-Annual", "Annual", "Final"]),
                    "due_date": types.Schema(type=types.Type.STRING, description="YYYY-MM-DD"),
                    "description": types.Schema(type=types.Type.STRING, description="Max 200 character summary")
                },
                required=["name", "report_type", "due_date", "description"]
            )
        ),
        "budget_ledger": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "period": types.Schema(type=types.Type.STRING, description="e.g., Year 1, Year 2"),
                    "investigator": types.Schema(type=types.Type.STRING),
                    "category": types.Schema(type=types.Type.STRING, enum=["Equipment", "F&A Cost", "Fringe Benefits", "Other Direct Costs", "Salaries & Wages", "Subawards", "Travel", "Tuition Remission"]),
                    "amount": types.Schema(type=types.Type.NUMBER),
                    "logic_validation_notes": types.Schema(type=types.Type.STRING)
                },
                required=["period", "investigator", "category", "amount"]
            )
        ),
        "labor_distribution": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "period": types.Schema(type=types.Type.STRING),
                    "rice_investigator": types.Schema(type=types.Type.STRING),
                    "salary_type": types.Schema(type=types.Type.STRING, enum=["AY", "Summer", "CY"]),
                    "effort_percentage": types.Schema(type=types.Type.NUMBER, description="Numeric percentage, e.g., 25.0"),
                    "salary_amount": types.Schema(type=types.Type.NUMBER)
                },
                required=["period", "rice_investigator", "salary_type", "effort_percentage", "salary_amount"]
            )
        ),
        "subaward_detailed_budget": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "period": types.Schema(type=types.Type.STRING),
                    "institution": types.Schema(type=types.Type.STRING),
                    "external_co_pi": types.Schema(type=types.Type.STRING),
                    "category": types.Schema(type=types.Type.STRING),
                    "amount": types.Schema(type=types.Type.NUMBER)
                },
                required=["period", "institution", "category", "amount"]
            )
        ),
        # --- Audit Flags & Logs ---
        "sequence_gap_detected": types.Schema(type=types.Type.BOOLEAN),
        "validation_flags": types.Schema(type=types.Type.STRING),
        "discrepancy_summary": types.Schema(type=types.Type.STRING, description="Consolidated summary of financial ambiguities"),

        # --- Chronology Matrix (Moved safely INSIDE the properties dictionary) ---
        "amendment_history": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "modification_number": types.Schema(type=types.Type.STRING, description="Base, Mod 01, Mod 02, etc."),
                    "effective_date": types.Schema(type=types.Type.STRING, description="YYYY-MM-DD"),
                    "action_type": types.Schema(type=types.Type.STRING, enum=["Base Award", "Funded Extension", "No-Cost Extension", "De-obligation", "Admin Change"]),
                    "funding_delta_obligated": types.Schema(type=types.Type.NUMBER, description="The specific financial change (+/-) of this document"),
                    "cumulative_obligated_total": types.Schema(type=types.Type.NUMBER, description="The rolling total obligated amount after this action"),
                    "funding_delta_anticipated": types.Schema(type=types.Type.NUMBER, description="The anticipated change (+/-) of this document"),
                    "cumulative_anticipated_total": types.Schema(type=types.Type.NUMBER, description="The rolling total anticipated amount after this action"),
                    "scope_or_terms_summary": types.Schema(type=types.Type.STRING, description="Concise summary of changes introduced")
                },
                required=["modification_number", "effective_date", "action_type", "funding_delta_obligated", "cumulative_obligated_total"]
            )
        )
    },
    
    # The 'required' list goes at the end, as a parameter of the main DATA_SCHEMA
    required=["parent_proposal_number", "award_name", "primary_sponsor", "total_funding_obligated", "deliverables", "budget_ledger"]
)

def extract_award_data(pdf_path, protocol_path="1.71.md"):
    """
    Ingests 1.71.md rules, programmatically strips out visual markdown 
    presentation layouts, and enforces pure analytical data extraction logic.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Target document not found at: {pdf_path}")
    if not os.path.exists(protocol_path):
        raise FileNotFoundError(f"Mandatory protocol ruleset missing at: {protocol_path}")
        
    # Step A: Load and clean behavioral rules dynamically
    print(f"切割 Ingesting behavioral rules from {protocol_path}...")
    with open(protocol_path, "r", encoding="utf-8") as f:
        raw_rules = f.read()
        
    # ⚡️ THE MASTER LOGIC FILTER: Separate Extraction Rules from UI Presentation Layouts
    # This keeps your 1.71 file intact for your Gem, but strips it for the API
    split_marker = "Final Output Styling & Layout Architecture"
    if split_marker in raw_rules:
        print("✂️ Visual markdown layout section detected and stripped for API safety.")
        system_instruction_rules = raw_rules.split(split_marker)[0]
    else:
        system_instruction_rules = raw_rules

    # Step B: Load cleaned local validations
    sponsors_df, investigators_df, orgs_df = load_all_registries(data_folder="data")
    
    # Step C: Stream file to API sandbox
    print(f"📦 Uploading file via modern Files API: {pdf_path}...")
    uploaded_file = client.files.upload(file=pdf_path)
    
    # Step D: Construct context data stream
    reference_payload = f"""
    You are an institutional ingestion router. Process the attached document according to Protocol 1.71 constraints.
    Cross-reference your extraction strings programmatically against the literal text entries provided below.
    If a soft match is achieved via punctuation/overrides instead of a literal match, append an asterisk (*) to the field value.
    
    --- REFERENCE REGISTER: SPONSORS ---
    {sponsors_df.to_csv(index=False)}
    
    --- REFERENCE REGISTER: INVESTIGATORS ---
    {investigators_df.to_csv(index=False)}
    
    --- REFERENCE REGISTER: RICE ORGS ---
    {orgs_df.to_csv(index=False)}
    """
    
    print("🤖 Querying gemini-2.5-flash with structured schema controls...")
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[uploaded_file, reference_payload],
        config=types.GenerateContentConfig(
            system_instruction=system_instruction_rules,
            response_mime_type="application/json",
            response_schema=DATA_SCHEMA,
            temperature=0.1
        ),
    )
    
    # Clean up the cloud file allocation instantly
    client.files.delete(name=uploaded_file.name)
    
    try:
        return json.loads(response.text)
    except json.JSONDecodeError as decode_error:
        print("\n⚠️ Formatting error caught! Saving raw response to 'debug_raw_response.txt' for review...")
        with open("debug_raw_response.txt", "w", encoding="utf-8") as debug_file:
            debug_file.write(response.text)
        raise decode_error