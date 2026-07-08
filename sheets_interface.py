import json
import os
import gspread
from google.oauth2.service_account import Credentials

# Define scopes required to read/write sheets and access drive files
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_sheets_client():
    """
    Bypasses organizational service account key blocks by prompting a local 
    browser login window to authenticate securely as your user profile.
    """
    # Fixed argument syntax: credentials_filename and authorized_user_filename
    return gspread.oauth(
        credentials_filename="credentials.json",
        authorized_user_filename="authorized_user.json"
    )

def initialize_workbook_tabs(spreadsheet_id):
    """
    Ensures the target workbook contains all 5 required relational tabs.
    Creates them with standard database headers if they are missing.
    """
    client = get_sheets_client()
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Define the mandatory sheets and their relational headers
    required_tabs = {
        "Awards_Review": [
            "Parent Proposal Number", "Award Name", "Primary Sponsor", "Start Date", 
            "End Date", "Principal Investigator", "Award Owning Organization", 
            "Sponsor Award Number", "Award Purpose", "Award Type", "Bill Type", 
            "Total Funding Obligated", "Review Status", "Discrepancy Summary", "Raw AI JSON"
        ],
        "Deliverables_Detail": ["Parent Proposal Number", "Name", "Report Type", "Due Date", "Description"],
        "Budget_Ledger": ["Parent Proposal Number", "Period", "Investigator", "Category", "Amount", "Logic Validation Notes"],
        "Labor_Distribution": ["Parent Proposal Number", "Period", "Rice Investigator", "Salary Type", "Effort %", "Salary Amount"],
        "Subaward_Budgets": ["Parent Proposal Number", "Period", "Institution", "External Co-PI", "Category", "Amount"]
    }
    
    for tab_name, headers in required_tabs.items():
        try:
            sheet = spreadsheet.worksheet(tab_name)
            print(f"✅ Found existing operational tab: {tab_name}")
        except gspread.exceptions.WorksheetNotFound:
            # Create the worksheet if it doesn't exist
            sheet = spreadsheet.add_worksheet(title=tab_name, rows="1000", cols=len(headers) + 2)
            sheet.append_row(headers)
            print(f"✨ Created missing relational workbook tab: {tab_name}")

def push_extracted_data_to_sheets(spreadsheet_id, data):
    """
    Breaks apart the nested multidimensional arrays and appends them 
    sequentially across the 5 structural workspace tabs.
    """
    client = get_sheets_client()
    spreadsheet = client.open_by_key(spreadsheet_id)
    
    # Extract our unique relational link key
    proposal_id = data.get("parent_proposal_number", "UNKNOWN_PROPOSAL")
    
    # 1. Append to Tab 1: Awards_Review (Flat Executive Dashboard)
    awards_sheet = spreadsheet.worksheet("Awards_Review")
    flat_metadata_row = [
        proposal_id,
        data.get("award_name"),
        data.get("primary_sponsor"),
        data.get("start_date"),
        data.get("end_date"),
        data.get("principal_investigator"),
        data.get("award_owning_organization"),
        data.get("sponsor_award_number"),
        data.get("award_purpose"),
        data.get("award_type"),
        data.get("bill_type"),
        data.get("total_funding_obligated"),
        "Pending Review",  # Default Review Status for HITL flow
        data.get("discrepancy_summary"),
        json.dumps(data)    # The Hidden Column Insurance Policy containing the raw payload
    ]
    awards_sheet.append_row(flat_metadata_row)
    
    # 2. Append to Tab 2: Deliverables_Detail (Field 32 Arrays)
    if "deliverables" in data and data["deliverables"]:
        deliv_sheet = spreadsheet.worksheet("Deliverables_Detail")
        rows = []
        for d in data["deliverables"]:
            rows.append([proposal_id, d.get("name"), d.get("report_type"), d.get("due_date"), d.get("description")])
        if rows:
            deliv_sheet.append_rows(rows)
            
    # 3. Append to Tab 3: Budget_Ledger (Field 34 Arrays)
    if "budget_ledger" in data and data["budget_ledger"]:
        budget_sheet = spreadsheet.worksheet("Budget_Ledger")
        rows = []
        for b in data["budget_ledger"]:
            rows.append([proposal_id, b.get("period"), b.get("investigator"), b.get("category"), b.get("amount"), b.get("logic_validation_notes")])
        if rows:
            budget_sheet.append_rows(rows)

    # 4. Append to Tab 4: Labor_Distribution (Field 35 Arrays)
    if "labor_distribution" in data and data["labor_distribution"]:
        labor_sheet = spreadsheet.worksheet("Labor_Distribution")
        rows = []
        for l in data["labor_distribution"]:
            rows.append([proposal_id, l.get("period"), l.get("rice_investigator"), l.get("salary_type"), l.get("effort_percentage"), l.get("salary_amount")])
        if rows:
            labor_sheet.append_rows(rows)

    # 5. Append to Tab 5: Subaward_Budgets (Field 36 Arrays)
    if "subaward_detailed_budget" in data and data["subaward_detailed_budget"]:
        sub_sheet = spreadsheet.worksheet("Subaward_Budgets")
        rows = []
        for s in data["subaward_detailed_budget"]:
            rows.append([proposal_id, s.get("period"), s.get("institution"), s.get("external_co_pi"), s.get("category"), s.get("amount")])
        if rows:
            sub_sheet.append_rows(rows)

    print(f"🎉 Relational mapping complete! Records successfully committed for Proposal: {proposal_id}")

if __name__ == "__main__":
    print("Sheets data-mapping interface successfully compiled.")