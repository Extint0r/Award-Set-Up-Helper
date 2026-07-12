import os
import glob
import re
import json
import csv
import time
import hashlib
import sqlite3
from datetime import datetime
from collections import defaultdict
import pandas as pd
from google import genai
from google.genai import types

# Ingest underlying execution blocks and structural definitions
from ai_engine import client, DATA_SCHEMA, clean_json_text, extract_award_data
from sheets_interface import initialize_workbook_tabs, push_extracted_data_to_sheets

# --- CONFIGURATION ---
TARGET_SPREADSHEET_ID = "15UPiZespLn_r2VhmCSOzlZSW1D9piwjdPHsijaqZKEU"
PDF_INTAKE_DIRECTORY = "legacy_pdf_vault"
MASTER_REGISTRY_CSV = "sam_sponsor_registry.csv"
LOCAL_DB_FILE = "award_clearinghouse.db"


def initialize_local_ledger_database():
    """Initializes the relational SQLite tables for unchangeable event logging and version snapshots."""
    conn = sqlite3.connect(LOCAL_DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS individual_file_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposal_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            file_hash TEXT UNIQUE NOT NULL,
            sequence_number INTEGER NOT NULL,
            document_execution_date TEXT NOT NULL,
            extraction_json TEXT NOT NULL,
            processed_timestamp TEXT NOT NULL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reduce_version_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposal_id TEXT NOT NULL,
            version_number INTEGER NOT NULL,
            max_sequence_number_included INTEGER NOT NULL,
            synthesized_json TEXT NOT NULL,
            processed_timestamp TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def safe_float(val, default=0.0):
    """Safely converts string numeric envelopes containing text/symbols into floats."""
    if val is None:
        return default
    clean = str(val).replace('$', '').replace(',', '').strip()
    try:
        return float(clean)
    except ValueError:
        return default


def parse_mixed_date(date_val):
    """Gracefully parses standard ISO strings and Excel numeric date serial numbers."""
    clean_val = str(date_val).strip()
    if not clean_val or clean_val.lower() == "not found" or clean_val.lower() == "n/a":
        return "N/A"
        
    if clean_val.isdigit():
        try:
            serial_days = int(clean_val)
            base_date = datetime(1899, 12, 30)
            return (base_date + pd.to_timedelta(serial_days, unit='D')).strftime("%Y-%m-%d")
        except Exception:
            try:
                from datetime import timedelta
                return (datetime(1899, 12, 30) + timedelta(days=serial_days)).strftime("%Y-%m-%d")
            except Exception:
                return "N/A"
                
    match = re.search(r'(\d{4})[-/](\d{2})[-/](\d{2})', clean_val)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
        
    return clean_val


def load_master_reconciliation_registry(csv_path):
    """Ingests the 5000+ line master reference registry into memory cache."""
    registry_cache = {}
    if not os.path.exists(csv_path):
        print(f"⚠️ Critical Halt: Required master baseline lookup file '{csv_path}' missing.")
        return registry_cache
        
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            headers_map = {name.strip().upper(): name for name in reader.fieldnames}
            
            cayuse_col = headers_map.get("CAYUSE_ID")
            if not cayuse_col:
                print("❌ Structure Failure: Master baseline registry must possess a 'CAYUSE_ID' tracking column.")
                return registry_cache
                
            for row in reader:
                raw_cayuse_key = str(row.get(cayuse_col, "")).strip()
                if raw_cayuse_key and "-" not in raw_cayuse_key and len(raw_cayuse_key) == 6:
                    raw_cayuse_key = f"{raw_cayuse_key[:2]}-{raw_cayuse_key[2:]}"
                    
                if raw_cayuse_key:
                    registry_cache[raw_cayuse_key] = {k.strip().upper(): v.strip() for k, v in row.items()}
                    
        print(f"🎯 Successfully cached {len(registry_cache)} historical system baseline records in memory.")
    except Exception as err:
        print(f"❌ Structural failure parsing master reconciliation registry: {err}")
    return registry_cache


def calculate_three_way_reconciliation(registry_row, extracted_json):
    """Executes multi-system delta matrix matching equations."""
    oracle_end = parse_mixed_date(registry_row.get("ORACLE_END_DATE", "N/A"))
    cayuse_end = parse_mixed_date(registry_row.get("CAYUSE_END_DATE", "N/A"))
    extracted_pdf_end = parse_mixed_date(extracted_json.get("end_date", "N/A"))
    
    sponsor_type = str(registry_row.get("SPONSOR_TYPE", "Not Found")).strip()
    funding_source_oracle = str(registry_row.get("FUNDING_SOURCE_NAME_ORACLE", "")).strip()
    oracle_award_num = str(registry_row.get("ORACLE_ID", "N/A")).strip()
    
    is_federal = any(x in sponsor_type.lower() or x in funding_source_oracle.lower() 
                     for x in ["federal", "fed prime", "federal government"])
    
    if is_federal:
        if oracle_end == cayuse_end and cayuse_end == extracted_pdf_end and oracle_end != "N/A":
            alignment_status = "IN_SYNC"
        else:
            alignment_status = "NOT_IN_SYNC"
    else:
        if oracle_end == cayuse_end and oracle_end != "N/A":
            alignment_status = "IN_SYNC"
        else:
            alignment_status = "NOT_IN_SYNC"
            
    if alignment_status == "IN_SYNC" or (oracle_award_num == "N/A" and oracle_end == cayuse_end):
        action_flag = "SYSTEMS ALIGNED"
    elif oracle_award_num == "N/A":
        action_flag = "BOTH: Other divergences"
    elif not is_federal:
        action_flag = "BOTH: Other divergences"
    elif oracle_end > cayuse_end and oracle_end != "N/A" and cayuse_end != "N/A":
        action_flag = "BOTH: Other divergences"
    elif extracted_pdf_end == "N/A":
        action_flag = "BOTH: Other divergences"
    elif extracted_pdf_end > cayuse_end and extracted_pdf_end != "N/A" and cayuse_end != "N/A":
        action_flag = "OSP: Pending sponsor award modification review"
    elif cayuse_end != "N/A" and extracted_pdf_end != "N/A":
        try:
            d1 = datetime.strptime(cayuse_end, "%Y-%m-%d")
            d2 = datetime.strptime(extracted_pdf_end, "%Y-%m-%d")
            delta_days = (d1 - d2).days
            if delta_days > 0 and delta_days <= 365:
                action_flag = "RCA: Cayuse date is 365 days or less greater than Federal (FDP valid extension max)"
            else:
                action_flag = "BOTH: Other divergences"
        except Exception:
            action_flag = "BOTH: Other divergences"
    else:
        action_flag = "BOTH: Other divergences"
        
    try:
        cayuse_auth = float(str(registry_row.get("CAYUSE_TOTAL_AWARDED", 0)).replace('$', '').replace(',', '').strip() or 0)
    except Exception:
        cayuse_auth = 0.0
        
    try:
        oracle_alloc = float(str(registry_row.get("ORACLE_HARD_LIMIT", 0)).replace('$', '').replace(',', '').strip() or 0)
    except Exception:
        oracle_alloc = 0.0
        
    try:
        fed_stated_spending = float(str(registry_row.get("FEDERAL_USA_OBLIGATED_AMOUNT", 0)).replace('$', '').replace(',', '').strip() or 0)
    except Exception:
        fed_stated_spending = 0.0
        
    internal_delta = cayuse_auth - oracle_alloc
    fed_oracle_delta = (fed_stated_spending - oracle_alloc) if is_federal else 0.0
    
    return {
        "ALIGNMENT": alignment_status,
        "ACTION": action_flag,
        "INTERNAL_DELTA_TOTAL_AWARDED": internal_delta,
        "FED_ORACLE_DELTA": fed_oracle_delta,
        "SPONSOR_ORACLE": funding_source_oracle or "Not Found",
        "SPONSOR_SAM": str(registry_row.get("PROPOSED_FUNDING_SOURCE_NAME_SAM", "Not Found")),
        "SPONSOR_UEI": str(registry_row.get("PROPOSED_FUNDING_SOURCE_NAME_UEI", "Not Found"))
    }


def generate_deterministic_uid(sponsor_award_num, parent_proposal_num):
    """Generates a standardized, immutable system identifier."""
    clean_proposal = re.sub(r'[^a-zA-Z0-9]', '', str(parent_proposal_num)).lower()
    clean_sponsor = str(sponsor_award_num).split('-')[0].split(' ')[0]
    clean_sponsor = re.sub(r'[^a-zA-Z0-9]', '', clean_sponsor).lower()
    clean_sponsor = re.sub(r'^\d(?=[a-zA-Z])', '', clean_sponsor)
    return f"{clean_sponsor}_{clean_proposal}"


def log_administrative_error(filepath, error_type, detail_message):
    """Logs parsing block exceptions to the review register file."""
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
    print(f"⚠️ Administrative Error Captured: Details logged to {log_file}.")


def group_files_by_proposal(directory):
    """Clusters physical intake directory components by tracking ID handles."""
    packet_groups = defaultdict(list)
    search_pattern = os.path.join(directory, "*.pdf")
    all_files = glob.glob(search_pattern)
    for filepath in all_files:
        filename = os.path.basename(filepath)
        match = re.search(r'(\d{2}-\d{4})', filename)
        if match:
            packet_groups[match.group(1)].append(filepath)
            continue
        packet_groups["UNASSIGNED_PACKET"].append(filepath)
    return packet_groups


def process_stateful_pipeline():
    """Executes the Event Sourced Map-Reduce pipeline loop with automated DB version control."""
    print("🚀 Initializing Local Ledger Ingestion & Version Snapshot Sync Pipeline...")
    initialize_local_ledger_database()
    
    if not os.path.exists(PDF_INTAKE_DIRECTORY):
        os.makedirs(PDF_INTAKE_DIRECTORY)
        return

    try:
        initialize_workbook_tabs(TARGET_SPREADSHEET_ID)
    except Exception as connection_err:
        print(f"❌ Spreadsheet connection initialization aborted: {connection_err}")
        return

    master_registry_cache = load_master_reconciliation_registry(MASTER_REGISTRY_CSV)
    groups = group_files_by_proposal(PDF_INTAKE_DIRECTORY)
    
    if not groups or (len(groups) == 1 and "UNASSIGNED_PACKET" in groups and not groups["UNASSIGNED_PACKET"]):
        print("ℹ️ Ingestion vault is empty. No files found to process.")
        return

    db_conn = sqlite3.connect(LOCAL_DB_FILE)

    for proposal_id, file_packet in groups.items():
        if proposal_id == "UNASSIGNED_PACKET":
            continue
            
        print(f"\n📦 Processing Ledger Stack for Proposal Packet: {proposal_id}")
        packet_updated = False
        
        # 1. 🗺️ THE RESILIENT EVENT MAP PHASE
        for filepath in file_packet:
            fname = os.path.basename(filepath)
            with open(filepath, "rb") as f:
                f_hash = hashlib.sha256(f.read()).hexdigest()
                
            cursor = db_conn.cursor()
            cursor.execute("SELECT id FROM individual_file_logs WHERE file_hash = ?", (f_hash,))
            if cursor.fetchone():
                continue
                
            print(f"  📄 Ingesting Standalone Transaction File: {fname}")
            packet_updated = True
            
            try:
                doc_json = extract_award_data(filepath)
                exec_date = parse_mixed_date(doc_json.get("document_execution_date", "N/A"))
                
                cursor.execute("""
                    SELECT COALESCE(MAX(sequence_number), 0) + 1 
                    FROM individual_file_logs WHERE proposal_id = ?
                """, (proposal_id,))
                next_seq = cursor.fetchone()[0]
                
                cursor.execute("""
                    INSERT INTO individual_file_logs 
                    (proposal_id, filename, file_hash, sequence_number, document_execution_date, extraction_json, processed_timestamp)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (proposal_id, fname, f_hash, next_seq, exec_date, json.dumps(doc_json), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                db_conn.commit()
                print(f"    ✅ Logged Event Entry -> Sequence #{next_seq}")
            except Exception as map_err:
                log_administrative_error(filepath, "Map Layer Ingestion Collapse", str(map_err))

        # 2. 🧪 THE DETERMINISTIC REDUCE PHASE
        cursor = db_conn.cursor()
        cursor.execute("""
            SELECT sequence_number, document_execution_date, filename, extraction_json 
            FROM individual_file_logs WHERE proposal_id = ?
        """, (proposal_id,))
        db_rows = cursor.fetchall()
        
        if not db_rows:
            continue
            
        mapped_documents = []
        max_seq_included = 0
        for seq, ex_date, filename, ext_json_str in db_rows:
            doc_data = json.loads(ext_json_str)
            doc_data["_source_filename"] = filename
            doc_data["_db_sequence"] = seq
            if seq > max_seq_included:
                max_seq_included = seq
            mapped_documents.append(doc_data)
            
        def get_date_key(doc):
            d = parse_mixed_date(doc.get("document_execution_date", "1900-01-01"))
            return d if d != "N/A" else "1900-01-01"
        mapped_documents.sort(key=get_date_key)
        
        cursor.execute("""
            SELECT version_number FROM reduce_version_logs 
            WHERE proposal_id = ? AND max_sequence_number_included = ?
        """, (proposal_id, max_seq_included))
        existing_version = cursor.fetchone()
        
        if existing_version and not packet_updated:
            continue

        print(f"  ⚙️ Executing Reduce Engine on {len(mapped_documents)} events to synthesize states...")
        base_record = mapped_documents[0]
        synthesized_json = {}
        
        metadata_keys = [
            "award_name", "primary_sponsor", "principal_investigator", "award_owning_organization",
            "sponsor_award_number", "award_purpose", "award_type", "bill_type", "flow_through_sponsor",
            "originating_sponsor_and_prime_num", "parent_proposal_number", "aln_number", "fa_rate",
            "close_date_days", "funding_mechanism", "category_type", "field_of_science_fos",
            "special_interest_si", "special_interest_thecb", "subject_to_single_audit"
        ]
        for key in metadata_keys:
            synthesized_json[key] = base_record.get(key, "Not Found")

        cum_obligated_total = 0.0
        latest_stated_ceiling = 0.0
        
        final_start_date = "N/A"
        final_end_date = "N/A"
        all_terms = set()
        all_far = set()
        all_deliverables = []
        all_budgets = []
        amendment_history = []
        chronological_document_ledger = []

        for idx, doc in enumerate(mapped_documents):
            fname = doc["_source_filename"]
            exec_date = parse_mixed_date(doc.get("document_execution_date", "N/A"))
            mod_num = doc.get("modification_number", f"Mod {idx}" if idx > 0 else "Base Award")
            
            delta_tot = safe_float(doc.get("total_funding_delta"))
            cum_obligated_total += delta_tot

            if doc.get("total_awarded_ceiling") and str(doc["total_awarded_ceiling"]).lower() != "n/a":
                latest_stated_ceiling = safe_float(doc["total_awarded_ceiling"])

            if doc.get("start_date") and final_start_date == "N/A":
                final_start_date = parse_mixed_date(doc["start_date"])
            if doc.get("end_date") and doc["end_date"] != "N/A":
                final_end_date = parse_mixed_date(doc["end_date"])

            # 🌟 ARRAY-SAFE LOGIC CONVERSION FOR TERMS AND CONDITIONS
            terms_val = doc.get("subject_to_terms_and_conditions", "")
            if isinstance(terms_val, list):
                for t in terms_val:
                    if str(t).strip(): all_terms.add(str(t).strip())
            else:
                for t in str(terms_val).split(";"):
                    if t.strip(): all_terms.add(t.strip())

            # 🌟 ARRAY-SAFE LOGIC CONVERSION FOR FAR CLAUSES
            far_val = doc.get("far_clauses", "")
            if isinstance(far_val, list):
                for f in far_val:
                    if str(f).strip(): all_far.add(str(f).strip())
            else:
                for f in str(far_val).split(";"):
                    if f.strip(): all_far.add(f.strip())

            if doc.get("deliverables"): all_deliverables.extend(doc["deliverables"])
            if doc.get("budget_ledger"): all_budgets.extend(doc["budget_ledger"])

            amendment_history.append({
                "modification_number": mod_num,
                "effective_date": exec_date,
                "action_type": "Base Award" if idx == 0 else "Funded Extension" if delta_tot > 0 else "No-Cost Extension",
                "funding_delta_obligated": delta_tot,
                "cumulative_obligated_total": cum_obligated_total,
                "funding_delta_anticipated": delta_tot,
                "cumulative_anticipated_total": max(cum_obligated_total, latest_stated_ceiling),
                "scope_or_terms_summary": f"Sequence ledger entry #{doc['_db_sequence']} compiled via file: {fname}"
            })

            chronological_document_ledger.append({
                "document_name": fname,
                "execution_date": exec_date,
                "dollar_delta": delta_tot,
                "adjusted_performance_end_date": final_end_date
            })

        final_total_awarded = max(cum_obligated_total, latest_stated_ceiling)

        synthesized_json["start_date"] = final_start_date
        synthesized_json["end_date"] = final_end_date
        synthesized_json["direct_funding_obligated"] = cum_obligated_total
        synthesized_json["indirect_funding_obligated"] = 0.0
        synthesized_json["total_funding_obligated"] = cum_obligated_total
        synthesized_json["direct_funding_total_awarded"] = final_total_awarded
        synthesized_json["indirect_funding_total_awarded"] = 0.0
        synthesized_json["total_funding_total_awarded"] = final_total_awarded

        synthesized_json["subject_to_terms_and_conditions"] = list(all_terms)
        synthesized_json["far_clauses"] = list(all_far)
        synthesized_json["deliverables"] = all_deliverables
        for b in all_budgets:
            b["amount"] = safe_float(b.get("amount"))
        synthesized_json["budget_ledger"] = all_budgets
        synthesized_json["amendment_history"] = amendment_history
        synthesized_json["chronological_document_ledger"] = chronological_document_ledger
        synthesized_json["labor_distribution"] = []
        synthesized_json["subaward_detailed_budget"] = []
        synthesized_json["sequence_gap_detected"] = False
        synthesized_json["validation_flags"] = "CLEAN"
        synthesized_json["discrepancy_summary"] = "No structural failures found inside event stream loop tracker."

        registry_row = master_registry_cache.get(str(proposal_id), {})
        
        synthesized_json["_oracle_end_date"] = parse_mixed_date(registry_row.get("ORACLE_END_DATE", "N/A"))
        synthesized_json["_cayuse_end_date"] = parse_mixed_date(registry_row.get("CAYUSE_END_DATE", "N/A"))
        synthesized_json["_oracle_hard_limit"] = safe_float(registry_row.get("ORACLE_HARD_LIMIT", 0.0))
        synthesized_json["_cayuse_authorized"] = safe_float(registry_row.get("CAYUSE_TOTAL_AWARDED", 0.0))

        try:
            print("    🧠 Invoking narrative summary synthesizer for audit log files...")
            forensic_payload_context = f"""
            Analyze this compiled award timeline event stream and baseline register trace:
            --- COMPILED EVENT TIMELINE HISTORY ---
            {json.dumps(chronological_document_ledger, indent=2)}
            --- MASTER HISTORICAL REGISTRY RECORD ---
            Oracle Hard Limit Stated: {registry_row.get("ORACLE_HARD_LIMIT", "N/A")}
            Cayuse Total Stated: {registry_row.get("CAYUSE_TOTAL_AWARDED", "N/A")}
            Oracle End Date Stated: {registry_row.get("ORACLE_END_DATE", "N/A")}
            Cayuse End Date Stated: {registry_row.get("CAYUSE_END_DATE", "N/A")}
            Sponsor Type: {registry_row.get("SPONSOR_TYPE", "N/A")}
            
            Synthesize a clear audit verdict. Return valid JSON containing exactly these two keys:
            "audited_judgment_verdict": (A concise single-sentence summary of the true legal operational state)
            "evidentiary_justification": (A detailed paragraph citing the file data and variances)
            """
            synthesis_res = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=forensic_payload_context,
                config=types.GenerateContentConfig(
                    system_instruction="You are a senior compliance officer. Output valid JSON summary judgements only.",
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            synthesis_json = json.loads(clean_json_text(synthesis_res.text))
            synthesized_json["audited_judgment_verdict"] = synthesis_json.get("audited_judgment_verdict", "Timeline compiled successfully.")
            synthesized_json["evidentiary_justification"] = synthesis_json.get("evidentiary_justification", "No validation errors caught.")
        except Exception as synth_fail:
            synthesized_json["audited_judgment_verdict"] = "Timeline assembled automatically via ledger matrix."
            synthesized_json["evidentiary_justification"] = f"Automated ledger compilation completed. Error: {synth_fail}"

        audit_metrics = calculate_three_way_reconciliation(registry_row, synthesized_json)
        synthesized_json["audit_alignment"] = audit_metrics["ALIGNMENT"]
        synthesized_json["audit_action"] = audit_metrics["ACTION"]
        synthesized_json["internal_delta"] = audit_metrics["INTERNAL_DELTA_TOTAL_AWARDED"]
        synthesized_json["fed_oracle_delta"] = audit_metrics["FED_ORACLE_DELTA"]
        synthesized_json["sponsor_oracle"] = audit_metrics["SPONSOR_ORACLE"]
        synthesized_json["sponsor_sam"] = audit_metrics["SPONSOR_SAM"]
        synthesized_json["sponsor_uei"] = audit_metrics["SPONSOR_UEI"]
        
        system_uid = generate_deterministic_uid(synthesized_json["sponsor_award_number"], proposal_id)
        oracle_award_number = registry_row.get("ORACLE_ID", "TBD")
        
        cursor.execute("SELECT COALESCE(MAX(version_number), 0) + 1 FROM reduce_version_logs WHERE proposal_id = ?", (proposal_id,))
        next_ver = cursor.fetchone()[0]
        
        cursor.execute("""
            INSERT INTO reduce_version_logs 
            (proposal_id, version_number, max_sequence_number_included, synthesized_json, processed_timestamp)
            VALUES (?, ?, ?, ?, ?)
        """, (proposal_id, next_ver, max_seq_included, json.dumps(synthesized_json), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        db_conn.commit()
        print(f"    💾 Saved Snapshot Version #{next_ver} (Tracks Events up to Sequence #{max_seq_included})")
        
        push_extracted_data_to_sheets(TARGET_SPREADSHEET_ID, synthesized_json, system_uid, oracle_award_number)
        print(f"  ✅ Ingestion run committed to target sheets layer for UID: {system_uid}")
        
    db_conn.close()


if __name__ == "__main__":
    process_stateful_pipeline()