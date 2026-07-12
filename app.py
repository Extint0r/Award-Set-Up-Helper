import os
import pandas as pd
import streamlit as st

# Import core functional execution engines
from batch_processor import process_stateful_pipeline, PDF_INTAKE_DIRECTORY, TARGET_SPREADSHEET_ID
from sheets_interface import get_sheets_client

# Corporate Application Configuration
st.set_page_config(
    page_title="Rice Contract Reviewer",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Deep Institutional Theme Custom Styling
st.markdown("""
    <style>
    .metric-card { background-color: #f8f9fa; border-left: 5px solid #002b49; padding: 15px; border-radius: 5px; margin-bottom: 10px; }
    .status-badge { background-color: #e3f2fd; color: #0d47a1; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .sync-badge-green { background-color: #d4edda; color: #155724; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .sync-badge-red { background-color: #f8d7da; color: #721c24; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .alert-badge { background-color: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; border: 1px solid #ffeeba; font-family: monospace; }
    .section-header { color: #002b49; border-bottom: 2px solid #002b49; padding-bottom: 5px; margin-top: 25px; margin-bottom: 15px; }
    .judgment-box { background-color: #e8f4fd; border-left: 5px solid #2196f3; padding: 15px; border-radius: 5px; margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=10)
def fetch_all_sheets_data(spreadsheet_id):
    """
    Connects securely to Google Sheets and caches data matrices 
    for 10 seconds to maintain UI responsiveness.
    """
    try:
        client = get_sheets_client()
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        data_frames = {}
        tabs = ["Awards_Review", "Document_Ledger", "Deliverables_Detail", "Budget_Ledger", "Labor_Distribution", "Subaward_Budgets"]
        for tab in tabs:
            try:
                worksheet = spreadsheet.worksheet(tab)
                records = worksheet.get_all_records()
                data_frames[tab] = pd.DataFrame(records)
            except Exception:
                data_frames[tab] = pd.DataFrame()
        return data_frames
    except Exception as e:
        st.error(f"Spreadsheet Synchronization Failure: {e}")
        return None


# SIDEBAR CONTROL CONSOLE
with st.sidebar:
    st.title("Rice Contract Reviewer")
    st.markdown("Database Processing Control Console")
    st.markdown("---")
    
    st.subheader("Staging Vault Upload")
    uploaded_files = st.file_uploader(
        "Upload contract PDF notices or amendment modifications:",
        type=["pdf"],
        accept_multiple_files=True,
        help="Assets are automatically copied to the local staging workspace."
    )
    
    if uploaded_files:
        if not os.path.exists(PDF_INTAKE_DIRECTORY):
            os.makedirs(PDF_INTAKE_DIRECTORY)
        
        saved_count = 0
        for f in uploaded_files:
            target_path = os.path.join(PDF_INTAKE_DIRECTORY, f.name)
            if not os.path.exists(target_path):
                with open(target_path, "wb") as out_file:
                    out_file.write(f.read())
                saved_count += 1
        if saved_count > 0:
            st.sidebar.success(f"Successfully staged {saved_count} file(s).")

    st.markdown("---")
    
    st.subheader("Processing Engine")
    if st.button("Run Ingestion Pipeline", width="stretch"):
        with st.spinner("Processing documents under Protocol 1.71 constraints..."):
            with st.status("Executing Portfolio Ingestion...", expanded=True) as status:
                st.write("Initializing file structural scan and deduplication arrays...")
                process_stateful_pipeline()
                status.update(label="Ingestion processing cycle finished.", state="complete", expanded=False)
            st.toast("Google Sheets database successfully updated.")
            st.rerun()

# MAIN WORKSPACE BOARD AREA
st.title("Cayuse-Oracle Award Review Dashboard")
st.markdown("---")

data_pools = fetch_all_sheets_data(TARGET_SPREADSHEET_ID)

# 🌟 MODIFIED: Brought lookups outside the conditional gate to protect scannability
st.subheader("Portfolio Cross-Reference Lookup")
search_query = st.text_input(
    "Search records instantly by System UID, Parent Proposal (Cayuse), or Oracle ID:",
    placeholder="Enter identification token string..."
).strip().lower()

if data_pools and not data_pools["Awards_Review"].empty:
    df_awards = data_pools["Awards_Review"]
    
    if "selected_uid" not in st.session_state:
        st.session_state.selected_uid = df_awards.iloc[0]['System UID']
    
    if search_query:
        matched_rows = df_awards[
            df_awards['System UID'].astype(str).str.lower().str.contains(search_query) |
            df_awards['Parent Proposal Number'].astype(str).str.lower().str.contains(search_query) |
            df_awards['Oracle Award Number'].astype(str).str.lower().str.contains(search_query)
        ]
    else:
        matched_rows = df_awards

    if matched_rows.empty:
        st.warning("No records discovered matching that search attribute.")
    else:
        # DEFENSIVE GUARD: Automatically inject missing columns if the Google Sheet has old Row 1 headers
        required_audit_columns = {
            "Date Alignment Status": "NOT_IN_SYNC",
            "Reconciliation Action Flag": "Review Divergence",
            "Internal Budget Delta": 0.0,
            "Federal Oracle Delta": 0.0,
            "Audited Judgment Verdict": "Pending System Sync",
            "Evidentiary Justification": "Run ingestion pipeline with clean sheet tabs.",
            "Discrepancy Summary": "No structural discrepancies logged.",
            "Sponsor (Cayuse)": "Not Found",
            "Sponsor (Oracle)": "Not Found",
            "Sponsor (SAM.gov)": "Not Found",
            "Sponsor UEI": "N/A"
        }
        
        for col_name, default_value in required_audit_columns.items():
            if col_name not in matched_rows.columns:
                matched_rows[col_name] = default_value

        # Guardrail: If search parameters isolate out the current selection, shift safely onto first available row
        if st.session_state.selected_uid not in matched_rows['System UID'].values:
            st.session_state.selected_uid = matched_rows.iloc[0]['System UID']

        # ──────────────────────────────────────────────────────────────────
        # PANEL 1: SYSTEM DIRECTORY PROMPT MATRIX
        # ──────────────────────────────────────────────────────────────────
        st.markdown("<div class='section-header'>### Contract Directory Panel</div>", unsafe_allow_html=True)
        
        list_display_df = matched_rows[[
            "Oracle Award Number",
            "Parent Proposal Number",
            "End Date",
            "Total Stated Awarded Budget",
            "Date Alignment Status",      
            "Reconciliation Action Flag",  
            "System UID"
        ]].copy()
        
        list_display_df.columns = [
            "Award Number (Oracle)",
            "Project Number (Cayuse)",
            "Award End Date",
            "Total Awarded Budget",
            "Sync Status",
            "Workflow Action Owner",
            "System UID"
        ]
        
        st.dataframe(
            list_display_df,
            width="stretch",
            hide_index=True,
            selection_mode="single-row",
            key="directory_grid",
            on_select="rerun",
            column_config={
                "Total Awarded Budget": st.column_config.NumberColumn(format="$%,d"),
                "System UID": None
            }
        )
        
        if "directory_grid" in st.session_state and st.session_state.directory_grid.get("selection", {}).get("rows"):
            clicked_row_pos = st.session_state.directory_grid["selection"]["rows"][0]
            if clicked_row_pos < len(list_display_df):
                st.session_state.selected_uid = list_display_df.iloc[clicked_row_pos]["System UID"]

        # ──────────────────────────────────────────────────────────────────
        # PANEL 2: FOCUSED CONTRACT ANALYSIS CARD
        # ──────────────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-header'>### Focused Analysis Profile Card</div>", unsafe_allow_html=True)
        
        if st.session_state.selected_uid in matched_rows['System UID'].values:
            row = matched_rows[matched_rows['System UID'] == st.session_state.selected_uid].iloc[0]
        else:
            row = matched_rows.iloc[0]
            st.session_state.selected_uid = row['System UID']
            
        uid = row['System UID']
        cayuse_id = row['Parent Proposal Number']
        oracle_id = row['Oracle Award Number']
        
        st.markdown(f"#### Profile Target: {row['Award Name']}")
        
        # Upper Identification Parameters Grid
        col1, col2, col3, col4 = st.columns(4)
        col1.markdown(f"**System UID:** `{uid}`")
        col2.markdown(f"**Parent Proposal (Cayuse):** `{cayuse_id}`")
        col3.markdown(f"**Oracle Award ID:** `{oracle_id if oracle_id else 'TBD'}`")
        
        # Color code the triage sync status badges
        sync_html = f"<span class='sync-badge-green'>In Sync</span>" if row['Date Alignment Status'] == "IN_SYNC" else f"<span class='sync-badge-red'>Out of Sync ({row['Reconciliation Action Flag']})</span>"
        col4.markdown(f"**Audit Alignment:** {sync_html}", unsafe_allow_html=True)
        
        # Lower Financial Metrics Grid
        st.markdown("<br>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric(label="Total Obligated Funding (PDF)", value=f"${row['Total Funding Obligated']:,}")
        with m2:
            st.metric(label="Total Stated Awarded Budget (PDF)", value=f"${row['Total Stated Awarded Budget']:,}")
        with m3:
            st.metric(label="Internal Budget Delta (Cayuse vs Oracle)", value=f"${row['Internal Budget Delta']:,}")
        with m4:
            st.metric(label="Federal Oracle Delta (USA vs Oracle)", value=f"${row['Federal Oracle Delta']:,}")
            
        if str(row['Discrepancy Summary']).strip() and str(row['Discrepancy Summary']).lower() != "not found":
            st.markdown(f"<div class='alert-badge'><b>System Audit Notice:</b><br>{row['Discrepancy Summary']}</div>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # Relational Sub-Table Data Views
        with st.expander("View Complete Relational Data Sheets", expanded=True):
            tab_profile, tab_ledger, tab_labor, tab_subawards, tab_compliance = st.tabs([
                "General Profile", "Master Budget Ledger", "Labor Distribution", "Subrecipients", "Compliance & Audited Judgments"
            ])
            
            with tab_profile:
                p1, p2, p3 = st.columns(3)
                p1.write(f"**Principal Investigator:** {row['Principal Investigator']}")
                p2.write(f"**Award Owning Org:** {row['Award Owning Organization']}")
                p3.write(f"**Sponsor Award Number:** `{row['Sponsor Award Number']}`")
                
                st.markdown("---")
                st.markdown("**Triple-Name Sponsor Integration Matrix:**")
                s1, s2, s3 = st.columns(3)
                s1.write(f"**Sponsor Name (Cayuse):** {row['Sponsor (Cayuse)']}")
                s2.write(f"**Sponsor Name (Oracle):** {row['Sponsor (Oracle)']}")
                s3.write(f"**Sponsor Name (SAM.gov):** {row['Sponsor (SAM.gov)']} `({row['Sponsor UEI']})`")
                st.markdown("---")

                p4, p5, p6 = st.columns(3)
                p4.write(f"**Project Start Date:** {row['Start Date']}")
                p5.write(f"**Project End Date:** {row['End Date']}")
                p6.write(f"**Billing Method:** {row['Bill Type']}")

            with tab_ledger:
                df_ledge = data_pools["Budget_Ledger"]
                if not df_ledge.empty and 'System UID' in df_ledge.columns:
                    matched_ledge = df_ledge[df_ledge['System UID'] == uid].drop(columns=['System UID', 'Parent Proposal Number'], errors='ignore')
                    st.dataframe(matched_ledge, width="stretch", hide_index=True)
                else:
                    st.info("No budget sub-ledger items found.")

            with tab_labor:
                df_labor = data_pools["Labor_Distribution"]
                if not df_labor.empty and 'System UID' in df_labor.columns:
                    matched_labor = df_labor[df_labor['System UID'] == uid].drop(columns=['System UID', 'Parent Proposal Number'], errors='ignore')
                    st.dataframe(matched_labor, width="stretch", hide_index=True)
                else:
                    st.info("No personnel labor allocation records found.")

            with tab_subawards:
                df_sub = data_pools["Subaward_Budgets"]
                if not df_sub.empty and 'System UID' in df_sub.columns:
                    matched_sub = df_sub[df_sub['System UID'] == uid].drop(columns=['System UID', 'Parent Proposal Number'], errors='ignore')
                    st.dataframe(matched_sub, width="stretch", hide_index=True)
                else:
                    st.info("No active subrecipient institutional lines logged.")

            with tab_compliance:
                st.markdown("#### Chronological Document Ledger (Stated PDF Truth)")
                df_docs = data_pools["Document_Ledger"]
                if not df_docs.empty and 'System UID' in df_docs.columns:
                    matched_docs = df_docs[df_docs['System UID'] == uid].drop(columns=['System UID', 'Parent Proposal Number'], errors='ignore')
                    st.dataframe(matched_docs, width="stretch", hide_index=True)
                else:
                    st.info("No historical document entries tracked for this portfolio transaction packet.")
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("#### Independent Forensic AI Audited Judgment")
                
                st.markdown(f"""
                <div class='judgment-box'>
                    <b>⚖️ Audited Verdict Declaration:</b><br>{row['Audited Judgment Verdict']}
                </div>
                <div class='metric-card' style='margin-top: 15px;'>
                    <b>🔍 Forensic Evidentiary Rationale:</b><br>{row['Evidentiary Justification']}
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")
else:
    st.info("The application portfolio database is currently unpopulated. Stage award timeline files in the control panel to initialize the reviewer screen.")