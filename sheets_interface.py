import gspread
from google.oauth2.service_account import Credentials

def get_sheets_client():
    """
    Bypasses organizational service account key blocks by prompting a local 
    browser login window to authenticate securely as your user profile.
    """
    return gspread.oauth(
        credentials_filename="credentials.json",
        authorized_user_filename="authorized_user.json"
    )

def delete_existing_uid_records(spreadsheet, system_uid):
    """
    Scans Column A across all 5 operational tabs. If the incoming System UID 
    already exists, it purges those specific rows to make room for the updated 
    chronological data package, preventing duplicate accumulation.
    """
    tabs_to_clean = ["Awards_Review", "Deliverables_Detail", "Budget_Ledger", "Labor_Distribution", "Subaward_Budgets"]
    for tab_name in tabs_to_clean:
        try:
            sheet = spreadsheet.worksheet(tab_name)
            matching_cells = sheet.findall(system_uid, in_column=1)
            if matching_cells:
                rows_to_delete = sorted([cell.row for cell in matching_cells], reverse=True)
                print(f"    🧹 Purging {len(rows_to_delete)} obsolete rows from tab: '{tab_name}'...")
                for row_num in rows_to_delete:
                    sheet.delete_rows(row_num)
        except Exception as clean_err:
            print(f"    ⚠️ Non-critical cleaning interruption on tab '{tab_name}': {clean_err}")

def initialize_workbook_tabs(spreadsheet_id):
    """
    Ensures the target workbook contains all 5 required relational tabs
    governed by our absolute generated System UID.
    """
    client = get_sheets_client()
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # 🌟 UPGRADED SCHEMA: Added "Oracle Award Number" as Column C
    required_tabs = {
        "Awards_Review": [
            "System UID", "Parent Proposal Number", "Oracle Award Number", "Award Name", "Primary Sponsor", 
            "Start Date", "End Date", "Principal Investigator", "Award Owning Organization", 
            "Sponsor Award Number", "Award Purpose", "Award Type", "Bill Type", 
            "Direct Funding Obligated", "Indirect Funding Obligated", "Total Funding Obligated", 
            "Direct Stated Awarded Budget", "Indirect Stated Awarded Budget", "Total Stated Awarded Budget", 
            "Review Status", "Discrepancy Summary"
        ],
        "Deliverables_Detail": ["System UID", "Parent Proposal Number", "Name", "Report Type", "Due Date", "Description"],
        "Budget_Ledger": ["System UID", "Parent Proposal Number", "Period", "Investigator", "Category", "Amount", "Logic Validation Notes"],
        "Labor_Distribution": ["System UID", "Parent Proposal Number", "Period", "Rice Investigator", "Salary Type", "Effort %", "Salary Amount"],
        "Subaward_Budgets": ["System UID", "Parent Proposal Number", "Period", "Institution", "External Co-PI", "Category", "Amount"]
    }
    
    for tab_name, headers in required_tabs.items():
        try:
            spreadsheet.worksheet(tab_name)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(title=tab_name, rows="1000", cols=len(headers) + 2)
            sheet.append_row(headers)

def push_extracted_data_to_sheets(spreadsheet_id, data, system_uid, oracle_award_number=""):
    """
    Pushes data matrices using the structural System UID, Cayuse Proposal Number,
    and the cross-referenced Oracle Award Number (if available).
    """
    client = get_sheets_client()
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Purge any old timeline rows for this specific UID to make room for updated cumulative packet calculations
    delete_existing_uid_records(spreadsheet, system_uid)
    
    proposal_id = data.get("parent_proposal_number", "UNKNOWN_PROPOSAL")
    
    # 1. Append to Tab 1: Awards_Review (Now tracking explicit double-entry sub-totals)
    awards_sheet = spreadsheet.worksheet("Awards_Review")
    flat_metadata_row = [
        system_uid,                             # Column A: System UID
        proposal_id,                            # Column B: Parent Proposal Number
        oracle_award_number or "",              # Column C: Oracle Award Number (🌟 NEW CROSS-REFERENCE COLUMN)
        data.get("award_name"),                 # Column D: Award Name
        data.get("primary_sponsor"),            # Column E: Primary Sponsor
        data.get("start_date"),                 # Column F: Start Date
        data.get("end_date"),                   # Column G: End Date
        data.get("principal_investigator"),     # Column H: PI Name
        data.get("award_owning_organization"),  # Column I: Award Owning Organization
        data.get("sponsor_award_number"),       # Column J: Sponsor Award Number
        data.get("award_purpose"),              # Column K: Award Purpose
        data.get("award_type"),                 # Column L: Award Type
        data.get("bill_type"),                  # Column M: Bill Type
        
        # OBLIGATED BREAKDOWN (Current Funding Matrix)
        data.get("direct_funding_obligated"),   # Column N: Direct Funding (Obligated)
        data.get("indirect_funding_obligated"), # Column O: Indirect Funding (Obligated)
        data.get("total_funding_obligated"),    # Column P: Total Funding (Obligated)
        
        # TOTAL AWARDED BREAKDOWN (Long-Term Project Anticipated Matrix)
        data.get("direct_funding_total_awarded"),   # Column Q: Direct Funding (Total Awarded)
        data.get("indirect_funding_total_awarded"), # Column R: Indirect Funding (Total Awarded)
        data.get("total_funding_total_awarded"),    # Column S: Total Funding (Total Awarded)
        
        "Pending Review",                       # Column T: Review Status
        data.get("discrepancy_summary")           # Column U: Discrepancy Summary
    ]
    awards_sheet.append_row(flat_metadata_row)
    
    # 2. Append to Tab 2: Deliverables_Detail
    if "deliverables" in data and data["deliverables"]:
        deliv_sheet = spreadsheet.worksheet("Deliverables_Detail")
        rows = [[system_uid, proposal_id, d.get("name"), d.get("report_type"), d.get("due_date"), d.get("description")] for d in data["deliverables"]]
        deliv_sheet.append_rows(rows)
            
    # 3. Append to Tab 3: Budget_Ledger
    if "budget_ledger" in data and data["budget_ledger"]:
        budget_sheet = spreadsheet.worksheet("Budget_Ledger")
        rows = [[system_uid, proposal_id, b.get("period"), b.get("investigator"), b.get("category"), b.get("amount"), b.get("logic_validation_notes")] for b in data["budget_ledger"]]
        budget_sheet.append_rows(rows)

    # 4. Append to Tab 4: Labor_Distribution
    if "labor_distribution" in data and data["labor_distribution"]:
        labor_sheet = spreadsheet.worksheet("Labor_Distribution")
        rows = [[system_uid, proposal_id, l.get("period"), l.get("rice_investigator"), l.get("salary_type"), l.get("effort_percentage"), l.get("salary_amount")] for l in data["labor_distribution"]]
        labor_sheet.append_rows(rows)

    # 5. Append to Tab 5: Subaward_Budgets
    if "subaward_detailed_budget" in data and data["subaward_detailed_budget"]:
        sub_sheet = spreadsheet.worksheet("Subaward_Budgets")
        rows = [[system_uid, proposal_id, s.get("period"), s.get("institution"), s.get("external_co_pi"), s.get("category"), s.get("amount")] for s in data["subaward_detailed_budget"]]
        sub_sheet.append_rows(rows)

def delete_existing_uid_records(spreadsheet, system_uid):
    """
    Scans Column A across all 5 operational tabs. If the incoming System UID 
    already exists, it purges those specific rows to make room for the updated 
    chronological data package, preventing duplicate accumulation.
    """
    tabs_to_clean = ["Awards_Review", "Deliverables_Detail", "Budget_Ledger", "Labor_Distribution", "Subaward_Budgets"]
    
    for tab_name in tabs_to_clean:
        try:
            sheet = spreadsheet.worksheet(tab_name)
            # Find all cell matches exactly in Column 1 (Column A)
            matching_cells = sheet.findall(system_uid, in_column=1)
            
            if matching_cells:
                # Sort matching row numbers in descending order (bottom-to-top)
                # This ensures deleting a row doesn't shift the indexes of rows above it
                rows_to_delete = sorted([cell.row for cell in matching_cells], reverse=True)
                
                print(f"    🧹 Purging {len(rows_to_delete)} obsolete rows from tab: '{tab_name}'...")
                for row_num in rows_to_delete:
                    sheet.delete_rows(row_num)
        except Exception as clean_err:
            print(f"    ⚠️ Non-critical cleaning interruption on tab '{tab_name}': {clean_err}")