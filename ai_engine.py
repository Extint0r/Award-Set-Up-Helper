import os
import json
import re
from google import genai
from google.genai import types as genai_types  # Aliased to eliminate Pylance collision with standard 'types'
from core_parser import load_all_registries
from dotenv import load_dotenv

# 1. Securely load environmental memory keys
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("Critical Configuration Deficit: 'GEMINI_API_KEY' was not found.")

client = genai.Client(api_key=API_KEY)


def clean_json_text(text):
    """
    Safely removes markdown wraps if present. Uses explicit string replaces 
    to eliminate backtick serialization collisions during code deployments.
    """
    clean = text.strip()
    # Masking out markdown tags safely
    if clean.startswith("xxx".replace('x', '\x60')):
        clean = clean.replace("xxxjson".replace('x', '\x60'), "")
        clean = clean.replace("xxx".replace('x', '\x60'), "")
    return clean.strip()


## Define the Standalone Single-Document Extraction Schema using Explicit genai_types
DATA_SCHEMA = genai_types.Schema(
    type=genai_types.Type.OBJECT,
    properties={
        # Standalone File Operational Trackers
        "parent_proposal_number": genai_types.Schema(type=genai_types.Type.STRING, description="Format: YY-XXXX Unique Anchor Key"),
        "sponsor_award_number": genai_types.Schema(type=genai_types.Type.STRING),
        "modification_number": genai_types.Schema(type=genai_types.Type.STRING, description="The identifier of this specific document, e.g., Base Award, Mod 01, Amendment 2, NCE"),
        "document_execution_date": genai_types.Schema(type=genai_types.Type.STRING, description="The formal signature or execution date of this specific file in YYYY-MM-DD format"),
        
        # Core Contract Attributes Stated in this Specific Document
        "award_name": genai_types.Schema(type=genai_types.Type.STRING, description="Format: [Funder Acronym. Sponsor Award Number. Award Acronym. Lead PI Family Name]"),
        "primary_sponsor": genai_types.Schema(type=genai_types.Type.STRING, description="Legal agreement partner validated against Sponsors.csv"),
        "start_date": genai_types.Schema(type=genai_types.Type.STRING, description="YYYY-MM-DD"),
        "end_date": genai_types.Schema(type=genai_types.Type.STRING, description="YYYY-MM-DD Project Period End Date established or modified by this specific file"),
        "principal_investigator": genai_types.Schema(type=genai_types.Type.STRING, description="Lead PI name aligned with Investigators.csv"),
        "award_owning_organization": genai_types.Schema(type=genai_types.Type.STRING, description="Format: XXXXX-Department Name from Rice University Orgs.csv"),
        
        "award_purpose": genai_types.Schema(type=genai_types.Type.STRING, enum=["Research", "Instruction", "Training", "Public Service", "Construction", "Student Aid", "Agency Funds"]),
        "award_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Federal Gov", "State/Local Gov", "State/Local Gov - Fed Prime", "Business", "Business - Fed Prime", "Non-Profit Org", "Non-Profit Org - Fed Prime", "Non-Profit Org - State/Local Prime", "Higher Ed", "Higher Ed - Fed Prime"]),
        "bill_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["CRM", "CRO", "CRQ", "CRY", "FBM", "FBO", "FBQ", "FBY", "SCHM", "SCHO", "SCHP", "SCHQ", "SCHY"]),
        "flow_through_sponsor": genai_types.Schema(type=genai_types.Type.STRING, enum=["Yes", "No"]),
        "originating_sponsor_and_prime_num": genai_types.Schema(type=genai_types.Type.STRING, description="Format: [Funder Name/Agency Acronym]: [Prime Award Number]"),
        "aln_number": genai_types.Schema(type=genai_types.Type.STRING, description="Format: XX.XXX"),
        "fa_rate": genai_types.Schema(type=genai_types.Type.STRING, description="Numerical percentage and base, e.g., 61% MTDC"),
        "close_date_days": genai_types.Schema(type=genai_types.Type.STRING, description="Number of closeout days as a string envelope, e.g., '90'"),
        "funding_mechanism": genai_types.Schema(type=genai_types.Type.STRING, enum=["Contract", "Cooperative Agreement", "Grant", "OTA"]),
        
        "subject_to_terms_and_conditions": genai_types.Schema(type=genai_types.Type.STRING, description="Continuous flat string of terms stated in this document, separated by semicolons"),
        "category_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Applied Research", "Basic Research", "Experimental Development", "Non Research"]),
        "far_clauses": genai_types.Schema(type=genai_types.Type.STRING, description="Continuous flat string of FAR clauses stated in this document, separated by semicolons"),
        
        "field_of_science_fos": genai_types.Schema(type=genai_types.Type.STRING, description="Extracted categories totaling 100%"),
        "special_interest_si": genai_types.Schema(type=genai_types.Type.STRING),
        "special_interest_thecb": genai_types.Schema(type=genai_types.Type.STRING),
        "subject_to_single_audit": genai_types.Schema(type=genai_types.Type.STRING),
        
        # Standalone Financial Changes Introduced by *This Specific File*
        "direct_funding_delta": genai_types.Schema(type=genai_types.Type.STRING, description="The incremental change (+/-) in direct funding introduced by this individual file"),
        "indirect_funding_delta": genai_types.Schema(type=genai_types.Type.STRING, description="The incremental change (+/-) in indirect funding introduced by this individual file"),
        "total_funding_delta": genai_types.Schema(type=genai_types.Type.STRING, description="The total incremental change introduced by this individual file (Use '0.0' for No-Cost Extensions)"),
        
        # Multi-Dimensional Sub-Arrays Stated in *This Standalone File*
        "deliverables": genai_types.Schema(
            type=genai_types.Type.ARRAY,
            items=genai_types.Schema(
                type=genai_types.Type.OBJECT,
                properties={
                    "name": genai_types.Schema(type=genai_types.Type.STRING),
                    "report_type": genai_types.Schema(type=genai_types.Type.STRING, enum=["Quarterly", "Semi-Annual", "Annual", "Final"]),
                    "due_date": genai_types.Schema(type=genai_types.Type.STRING, description="YYYY-MM-DD"),
                    "description": genai_types.Schema(type=genai_types.Type.STRING, description="Max 200 character summary")
                },
                required=["name", "report_type", "due_date", "description"]
            )
        ),
        "budget_ledger": genai_types.Schema(
            type=genai_types.Type.ARRAY,
            items=genai_types.Schema(
                type=genai_types.Type.OBJECT,
                properties={
                    "period": genai_types.Schema(type=genai_types.Type.STRING, description="e.g., Year 1, Year 2"),
                    "investigator": genai_types.Schema(type=genai_types.Type.STRING),
                    "category": genai_types.Schema(type=genai_types.Type.STRING, enum=["Equipment", "F&A Cost", "Fringe Benefits", "Other Direct Costs", "Salaries & Wages", "Subawards", "Travel", "Tuition Remission"]),
                    "amount": genai_types.Schema(type=genai_types.Type.STRING, description="monetary string value envelope")
                },
                required=["period", "investigator", "category", "amount"]
            )
        )
    },
    required=["parent_proposal_number", "sponsor_award_number", "modification_number", "document_execution_date", "total_funding_delta"]
)

def extract_award_data(pdf_path, protocol_path="1.71.md"):
    """Processes a single contract document with standalone structural layout stability."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Target document not found at: {pdf_path}")
    if not os.path.exists(protocol_path):
        raise FileNotFoundError(f"Mandatory protocol ruleset missing at: {protocol_path}")
        
    print(f"Ingesting standalone extraction constraints from {protocol_path}...")
    with open(protocol_path, "r", encoding="utf-8") as f:
        raw_rules = f.read()
        
    split_marker = "Final Output Styling & Layout Architecture"
    system_instruction_rules = raw_rules.split(split_marker)[0] if split_marker in raw_rules else raw_rules

    system_instruction_rules += (
        "\n\nSINGLE-FILE EXTRACTION FOCUS: You are analyzing ONE individual contract document from a timeline history. "
        "Extract only the financial deltas, target dates, and deliverables introduced specifically by this single attached file. "
        "All monetary values, efforts, and numbers must be output as standard double-quoted text strings."
    )

    sponsors_df, investigators_df, orgs_df = load_all_registries(data_folder="data")
    
    print(f"Uploading file via modern Files API: {os.path.basename(pdf_path)}...")
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
    """
    
    print("🤖 Querying gemini-2.5-flash with structured schema controls...")
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
    try:
        return json.loads(cleaned_json_str)
    except json.JSONDecodeError:
        return json.loads(cleaned_json_str, strict=False)