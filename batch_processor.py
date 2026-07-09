import os
import glob
import re
import json
import csv
import time
import hashlib
from datetime import datetime
from collections import defaultdict
from google import genai
from google.genai import types
from pypdf import PdfReader

# Ingest underlying execution blocks and structural definitions
from ai_engine import client, DATA_SCHEMA
from sheets_interface import initialize_workbook_tabs, push_extracted_data_to_sheets

# --- CONFIGURATION ---
TARGET_SPREADSHEET_ID = "15UPiZespLn_r2VhmCSOzlZSW1D9piwjdPHsijaqZKEU"
PDF_INTAKE_DIRECTORY = "legacy_pdf_vault"
ORACLE_MAPPING_CSV = "cayuse_oracle.csv"


def load_oracle_mapping(csv_path):
    """
    Reads the local 1-1 cross-walk CSV file dynamically. Normalizes column headers 
    to protect lookups from unexpected whitespaces or case-matching variance.
    """
    mapping = {}
    if not os.path.exists(csv_path):
        print(f"ℹ️ Optional mapping cross-walk '{csv_path}' not found. Moving forward without manual Oracle lookups.")
        return mapping
        
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames:
                return mapping
                
            headers_map = {name.strip().upper(): name for name in reader.fieldnames}
            cayuse_col = None
            oracle_col = None
            for normalized_name, original_name in headers_map.items():
                if "CAYUSE" in normalized_name:
                    cayuse_col = original_name
                if "ORACLE" in normalized_name:
                    oracle_col = original_name
                    
            if not cayuse_col or not oracle_col:
                print("⚠️ Structural Warning: 'cayuse_oracle.csv' must contain columns representing 'CAYUSE' and 'Oracle'. Lookups bypassed.")
                return mapping
                
            for row in reader:
                cayuse_key = str(row.get(cayuse_col, "")).strip()
                oracle_val = str(row.get(oracle_col, "")).strip()
                if cayuse_key:
                    mapping[cayuse_key] = oracle_val
                    
        print(f"🎯 Successfully loaded {len(mapping)} ERP identifiers from mapping cross-walk data file.")
    except Exception as err:
        print(f"⚠️ Non-critical failure parsing tracking cross-walk file: {err}")
    return mapping


def generate_deterministic_uid(sponsor_award_num, parent_proposal_num):
    """
    Generates a standardized, immutable system identifier by combining the 
    base core award token with the institutional proposal key.
    Strips leading agency application lifecycle type codes (e.g., 5R01 -> r01).
    """
    clean_proposal = re.sub(r'[^a-zA-Z0-9]', '', str(parent_proposal_num)).lower()
    clean_sponsor = str(sponsor_award_num).split('-')[0].split(' ')[0]
    clean_sponsor = re.sub(r'[^a-zA-Z0-9]', '', clean_sponsor).lower()
    clean_sponsor = re.sub(r'^\d(?=[a-zA-Z])', '', clean_sponsor)
    return f"{clean_sponsor}_{clean_proposal}"


def log_administrative_error(filepath, error_type, detail_message):
    """
    Isolates corrupted or unparseable files from the main database 
    and logs tracking data for human review.
    """
    log_file = "admin_error_log.csv"
    file_exists = os.path.exists(log_file)
    filename = os.path.basename(filepath)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    headers = ["Timestamp", "Filename", "Error Type", "Diagnostic Details", "Resolution Status"]
    row_data = [timestamp, filename, error_type, detail_message, "Pending Manual Review"]
    
    with open(log_file, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(headers)
        writer.writerow(row_data)
    print(f"⚠️ Administrative Error Captured: '{filename}' isolated. Details logged to {log_file}.")


def group_files_by_proposal(directory):
    """
    Pre-scans the staging directory to sort and cluster historical portfolios.
    Deploys SHA-256 byte hashing for duplication containment and local 3-page text extraction fallbacks.
    """
    packet_groups = defaultdict(list)
    search_pattern = os.path.join(directory, "*.pdf")
    all_files = glob.glob(search_pattern)
    
    seen_hashes = set()
    
    for filepath in all_files:
        filename = os.path.basename(filepath)
        
        # 🛡️ FAILSAFE 1: BYTE-LEVEL CONTENT DEDUPLICATION
        try:
            with open(filepath, "rb") as f:
                file_bytes = f.read()
                file_hash = hashlib.sha256(file_bytes).hexdigest()
                
            if file_hash in seen_hashes:
                print(f"🧹 Deduplication Filter: Skipping '{filename}' (Identical file content already staged in this batch).")
                continue
            seen_hashes.add(file_hash)
        except Exception as hash_err:
            print(f"⚠️ Warning: Hash validation skipped for '{filename}': {hash_err}")
            
        # 🛡️ IDENTITY DETECTOR
        # Strategy A: Scan Filename Mask
        match = re.search(r'(\d{2}-\d{4})', filename)
        if match:
            proposal_id = match.group(1)
            packet_groups[proposal_id].append(filepath)
            continue
            
        # Strategy B: Deep Content Fallback Scan (First 3 Pages)
        print(f"🔍 Filename mask missing for '{filename}'. Opening document for local deep page scan...")
        try:
            reader = PdfReader(filepath)
            max_pages = min(3, len(reader.pages))
            found_id = None
            
            for page_num in range(max_pages):
                page_text = reader.pages[page_num].extract_text() or ""
                content_match = re.search(r'(\d{2}-\d{4})', page_text)
                if content_match:
                    found_id = content_match.group(1)
                    break
                    
            if found_id:
                print(f"🎯 Verified Cayuse ID '{found_id}' inside text layers of '{filename}'. Aligning to packet.")
                packet_groups[found_id].append(filepath)
            else:
                packet_groups["UNASSIGNED_PACKET"].append(filepath)
                
        except Exception as pdf_err:
            log_administrative_error(filepath, "Local Extraction Interruption", f"Could not read internal text layer: {pdf_err}")
            packet_groups["UNASSIGNED_PACKET"].append(filepath)
            
    return packet_groups


def process_stateful_pipeline():
    """
    Executes the multi-document stateful ingestion stream under Protocol 1.71 constraints.
    """
    print("🚀 Initializing Multi-Document Stateful Chronology Pipeline...")
    
    if not os.path.exists(PDF_INTAKE_DIRECTORY):
        os.makedirs(PDF_INTAKE_DIRECTORY)
        print(f"📁 Created empty intake directory: '{PDF_INTAKE_DIRECTORY}'")
        return

    try:
        initialize_workbook_tabs(TARGET_SPREADSHEET_ID)
    except Exception as connection_err:
        print(f"❌ Spreadsheet connection initialization aborted: {connection_err}")
        return

    oracle_lookup_cache = load_oracle_mapping(ORACLE_MAPPING_CSV)
    groups = group_files_by_proposal(PDF_INTAKE_DIRECTORY)
    
    if not groups or (len(groups) == 1 and "UNASSIGNED_PACKET" in groups and not groups["UNASSIGNED_PACKET"]):
        print("ℹ️ Ingestion vault is empty. No files found to process.")
        return

    for proposal_id, file_packet in groups.items():
        if proposal_id == "UNASSIGNED_PACKET":
            for raw_file in file_packet:
                log_administrative_error(
                    filepath=raw_file, 
                    error_type="Naming Rule Violation", 
                    detail_message="Filename and first 3 pages lack a valid tracking code mask (YY-XXXX)."
                )
            continue
            
        print(f"\n📦 Processing Unified Historical Packet for Proposal: {proposal_id}")
        print(f"  Found {len(file_packet)} tied chronological documents.")
        
        try:
            uploaded_handles = []
            for filepath in file_packet:
                print(f"  📎 Uploading to context channel: {os.path.basename(filepath)}...")
                uploaded_handles.append(client.files.upload(file=filepath))
                
            with open("1.71.md", "r", encoding="utf-8") as f:
                raw_rules = f.read()
            system_rules = raw_rules.split("Final Output Styling & Layout Architecture")[0]
            
            prompt_instruction = f"""
            You are looking at a complete, unified historical packet for Proposal {proposal_id}.
            The attached files contain the baseline award notice and subsequent amendments/modifications.
            
            Execute Protocol 1.71 Part 4 and Part 5 across this combined collection:
            1. Identify which file establishes the base contract award.
            2. Sort all modification files in strict chronological order based on execution date.
            3. Trace the mathematical chain of custody. Determine if an amendment introduces a 'Funded Extension' 
               (increasing obligated funds) or a 'No-Cost Extension' (changing dates only).
            4. Populate the 'amendment_history' array showing the delta entries and the rolling cumulative totals.
            5. For your primary top-level metadata values, report the final, most recent state of the award.
            
            ⚠️ CRITICAL CROSS-VALIDATION MATHEMATICAL AUDIT:
            Before outputting the JSON response, you must execute an internal balance check:
            - Mathematically sum the 'amount' field across every row in your compiled 'budget_ledger' array.
            - Mathematically sum the 'salary_amount' field across every row in your 'labor_distribution' array.
            - Mathematically sum the 'amount' field across every row in your 'subaward_detailed_budget' array.
            
            Compare these computed sums against your extracted 'total_funding_obligated' and 'total_funding_total_awarded' metrics. These independent numbers extracted from the NOA face sheets MUST cleanly align and cross-validate each other. If there is a calculation discrepancy, a ledger misalignment, or a structural variance between the sub-tables and the master stated totals, you must calculate the exact dollar variance and state it explicitly inside the 'discrepancy_summary' field.
            """
            contents_payload = uploaded_handles + [prompt_instruction]
            
            print("  🤖 Running deep packet matrix synthesis on Gemini 2.5 Flash...")
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=contents_payload,
                config=types.GenerateContentConfig(
                    system_instruction=system_rules,
                    response_mime_type="application/json",
                    response_schema=DATA_SCHEMA,
                    temperature=0.1
                ),
            )
            
            for handle in uploaded_handles:
                client.files.delete(name=handle.name)
                
            extracted_json = json.loads(response.text)
            
            sponsor_num = extracted_json.get("sponsor_award_number")
            prop_num = extracted_json.get("parent_proposal_number")
            
            if not prop_num or str(prop_num).strip().lower() == "not found":
                raise ValueError("Critical ID Deficit: Gemini engine returned an empty or unverified Parent Proposal Number.")
            if not sponsor_num or str(sponsor_num).strip().lower() == "not found":
                raise ValueError("Critical ID Deficit: Gemini engine returned an empty or unverified Sponsor Award Number.")
                
            system_uid = generate_deterministic_uid(sponsor_num, prop_num)
            
            oracle_award_number = oracle_lookup_cache.get(str(proposal_id), "") or oracle_lookup_cache.get(str(prop_num), "")
            if oracle_award_number:
                print(f"  🔗 Oracle ERP Reference Linked: {oracle_award_number}")
            else:
                print("  🔗 Oracle ERP Reference: None Found in local cross-walk file.")
            
            push_extracted_data_to_sheets(TARGET_SPREADSHEET_ID, extracted_json, system_uid, oracle_award_number)
            print(f"  ✅ Chronological mapping successfully committed for System UID: {system_uid}")
            
            print("  ⏳ Cooling down API connection buffers for 3 seconds...")
            time.sleep(3)
            
        except json.JSONDecodeError as json_err:
            log_administrative_error(file_packet[0], "JSON Structural Corruption", str(json_err))
        except ValueError as val_err:
            log_administrative_error(file_packet[0], "Metadata Validation Failure", str(val_err))
        except Exception as general_err:
            log_administrative_error(file_packet[0], "Pipeline Runtime Interruption", str(general_err))


if __name__ == "__main__":
    process_stateful_pipeline()

