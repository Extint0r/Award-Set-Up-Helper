import os
import glob
import re
import json
import csv
from datetime import datetime
from collections import defaultdict
from google import genai
from google.genai import types

# Ingest underlying execution blocks and structural definitions
from ai_engine import client, DATA_SCHEMA
from sheets_interface import initialize_workbook_tabs, push_extracted_data_to_sheets

# ─── CONFIGURATION ───
# Replace with your personal Google Sheet ID (Verified in Test 1 & 2)
TARGET_SPREADSHEET_ID = "PASTE_YOUR_PERSONAL_TEST_SPREADSHEET_ID_HERE"
PDF_INTAKE_DIRECTORY = "legacy_pdf_vault"


def generate_deterministic_uid(sponsor_award_num, parent_proposal_num):
    """
    Generates a standardized, immutable system identifier by combining the 
    base core award token with the institutional proposal key.
    """
    # 1. Clean the Proposal Number (remove non-alphanumeric characters)
    clean_proposal = re.sub(r'[^a-zA-Z0-9]', '', str(parent_proposal_num)).lower()
    
    # 2. Clean the Sponsor Number (isolate core sequence, strip extensions/amendments)
    clean_sponsor = str(sponsor_award_num).split('-')[0].split(' ')[0]
    clean_sponsor = re.sub(r'[^a-zA-Z0-9]', '', clean_sponsor).lower()
    
    # 3. Concatenate into our absolute pipeline primary management key
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
    Pre-scans the folder files to cluster historical amendments together 
    by searching for the mandatory internal institutional tracking mask (YY-XXXX).
    """
    packet_groups = defaultdict(list)
    search_pattern = os.path.join(directory, "*.pdf")
    all_files = glob.glob(search_pattern)
    
    for filepath in all_files:
        filename = os.path.basename(filepath)
        # Look for the institutional proposal regex pattern: 2 digits + dash + 4 digits
        match = re.search(r'(\d{2}-\d{4})', filename)
        if match:
            proposal_id = match.group(1)
            packet_groups[proposal_id].append(filepath)
        else:
            # Fallback grouping for items without clear naming architecture
            packet_groups["UNASSIGNED_PACKET"].append(filepath)
            
    return packet_groups


def process_stateful_pipeline():
    """
    Executes the multi-document stateful ingestion stream under Protocol 1.71 constraints.
    """
    print("🚀 Initializing Multi-Document Stateful Chronology Pipeline...")
    
    # Ensure intake workspace folder layout exists
    if not os.path.exists(PDF_INTAKE_DIRECTORY):
        os.makedirs(PDF_INTAKE_DIRECTORY)
        print(f"📁 Created empty intake directory: '{PDF_INTAKE_DIRECTORY}'")
        return

    # Verify tracking matrix sheets are prepared
    try:
        initialize_workbook_tabs(TARGET_SPREADSHEET_ID)
    except Exception as connection_err:
        print(f"❌ Spreadsheet connection initialization aborted: {connection_err}")
        return

    # Group physical assets dynamically by structural anchor tags
    groups = group_files_by_proposal(PDF_INTAKE_DIRECTORY)
    
    if not groups or (len(groups) == 1 and "UNASSIGNED_PACKET" in groups and not groups["UNASSIGNED_PACKET"]):
        print("ℹ️ Ingestion vault is empty. No files found to process.")
        return

    for proposal_id, file_packet in groups.items():
        # Handle files that completely missed naming format restrictions
        if proposal_id == "UNASSIGNED_PACKET":
            for raw_file in file_packet:
                log_administrative_error(
                    filepath=raw_file, 
                    error_type="Naming Rule Violation", 
                    detail_message="Filename does not contain a valid matching institutional tracking code mask (YY-XXXX)."
                )
            continue
            
        print(f"\n📦 Processing Unified Historical Packet for Proposal: {proposal_id}")
        print(f"  Found {len(file_packet)} tied chronological documents.")
        
        try:
            # Step 1: Upload matching files concurrently into cloud processing memory channel
            uploaded_handles = []
            for filepath in file_packet:
                print(f"  📎 Uploading to context channel: {os.path.basename(filepath)}...")
                uploaded_handles.append(client.files.upload(file=filepath))
                
            # Step 2: Read rules, dynamically stripping out UI layout models
            with open("1.71.md", "r", encoding="utf-8") as f:
                raw_rules = f.read()
            system_rules = raw_rules.split("Final Output Styling & Layout Architecture")[0]
            
            # Step 3: Inject targeted delta tracking guidelines
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
            """
            contents_payload = uploaded_handles + [prompt_instruction]
            
            # Step 4: Stream the compiled token array to Gemini 2.5 Flash
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
            
            # Step 5: Clean up cloud allocations instantly
            for handle in uploaded_handles:
                client.files.delete(name=handle.name)
                
            # Step 6: Parse the text response structure
            extracted_json = json.loads(response.text)
            
            # Step 7: Identity Gate — Validate required metadata keys exist before sheet commitment
            sponsor_num = extracted_json.get("sponsor_award_number")
            prop_num = extracted_json.get("parent_proposal_number")
            
            if not prop_num or str(prop_num).strip().lower() == "not found":
                raise ValueError("Critical ID Deficit: Gemini engine returned an empty or unverified Parent Proposal Number.")
            if not sponsor_num or str(sponsor_num).strip().lower() == "not found":
                raise ValueError("Critical ID Deficit: Gemini engine returned an empty or unverified Sponsor Award Number.")
                
            # Step 8: Generate our calculated management identifier
            system_uid = generate_deterministic_uid(sponsor_num, prop_num)
            
            # Step 9: Commit data vectors across our multi-tab workbook
            push_extracted_data_to_sheets(TARGET_SPREADSHEET_ID, extracted_json, system_uid)
            print(f"  ✅ Chronological mapping successfully committed for System UID: {system_uid}")
            
        except json.JSONDecodeError as json_err:
            log_administrative_error(file_packet[0], "JSON Structural Corruption", str(json_err))
            
        except ValueError as val_err:
            log_administrative_error(file_packet[0], "Metadata Validation Failure", str(val_err))
            
        except Exception as general_err:
            log_administrative_error(file_packet[0], "Pipeline Runtime Interruption", str(general_err))


if __name__ == "__main__":
    process_stateful_pipeline()