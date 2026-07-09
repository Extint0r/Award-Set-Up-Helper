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
    .alert-badge { background-color: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; border: 1px solid #ffeeba; font-family: monospace; }
    .section-header { color: #002b49; border-bottom: 2px solid #002b49; padding-bottom: 5px; margin-top: 25px; margin-bottom: 15px; }
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
        tabs = ["Awards_Review", "Deliverables_Detail", "Budget_Ledger", "Labor_Distribution", "Subaward_Budgets"]
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
    st.markdown("")
    st.markdown("---")
    
    # Section A: File Staging Ingestion Vault
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
    
    # Section B: Core Engine Execution Trigger
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

# Sync relational data layers from cloud spreadsheet
data_pools = fetch_all_sheets_data(TARGET_SPREADSHEET_ID)

if data_pools and not data_pools["Awards_Review"].empty:
    df_awards = data_pools["Awards_Review"]
    
    # Initialize Persistent Selected Identifier Pointer
    if "selected_uid" not in st.session_state:
        st.session_state.selected_uid = df_awards.iloc[0]['System UID']
    
    # SYSTEM INTERFACE PORTFOLIO LOOKUP BAR
    st.subheader("System Cross-Reference Lookup")
    search_query = st.text_input(
        "Search records instantly by Generated Database UID, Parent Proposal Number (Cayuse Project), or Oracle ID (Award Number):",
        placeholder="Enter identification token string..."
    ).strip().lower()
    
    # Comprehensive triple-attribute lookup matching
    if search_query:
        matched_rows = df_awards[
            df_awards['System UID'].astype(str).str.lower().str.contains(search_query) |
            df_awards['Parent Proposal Number'].astype(str).str.lower().str.contains(search_query) |
            df_awards['Oracle Award Number'].astype(str).str.lower().str.contains(search_query)
        ]
    else:
        matched_rows = df_awards

    # Present Filtered Queue
    if matched_rows.empty:
        st.warning("No records discovered matching that search attribute.")
    else:
        # Guardrail: If search parameters isolate out the current selection, shift safely onto first available row
        if st.session_state.selected_uid not in matched_rows['System UID'].values:
            st.session_state.selected_uid = matched_rows.iloc[0]['System UID']

        # ──────────────────────────────────────────────────────────────────
        # PANEL 1: SYSTEM DIRECTORY PROMPT MATRIX (PERMANENT TOP SECTION)
        # ──────────────────────────────────────────────────────────────────
        st.markdown("<div class='section-header'>### Contract Directory Panel</div>", unsafe_allow_html=True)
        st.caption("Click the selection circle next to any row below to instantly call its data layers into the profile card below.")
        
        # Slice down to the exact requested five-attribute display matrix
        list_display_df = matched_rows[[
            "Oracle Award Number",
            "Parent Proposal Number",
            "End Date",
            "Total Stated Awarded Budget",
            "Total Funding Obligated",
            "System UID" # Passed in background for selection tracking
        ]].copy()
        
        # Formally rename column variables to deployment standard names
        list_display_df.columns = [
            "Award Number (Oracle)",
            "Project Number (Cayuse)",
            "Award End Date",
            "Total Awarded Budget",
            "Total Obligated Budget",
            "System UID"
        ]
        
        # Render clean interactive grid layout using a persistent state storage key
        st.dataframe(
            list_display_df,
            width="stretch",
            hide_index=True,
            selection_mode="single-row",
            key="directory_grid", # 🌟 FIXED: Added stable state synchronization key
            on_select="rerun",
            column_config={
                "Total Awarded Budget": st.column_config.NumberColumn(format="$%,d"),
                "Total Obligated Budget": st.column_config.NumberColumn(format="$%,d"),
                "System UID": None # Keeps background relation key hidden from user view
            }
        )
        
        # Intercept persistent memory selections before rendering downstream components
        if "directory_grid" in st.session_state and st.session_state.directory_grid.get("selection", {}).get("rows"):
            clicked_row_pos = st.session_state.directory_grid["selection"]["rows"][0]
            if clicked_row_pos < len(list_display_df):
                # Update absolute identification token state on the fly
                st.session_state.selected_uid = list_display_df.iloc[clicked_row_pos]["System UID"]

        # ──────────────────────────────────────────────────────────────────
        # PANEL 2: FOCUSED CONTRACT ANALYSIS CARD (PERMANENT BOTTOM SECTION)
        # ──────────────────────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div class='section-header'>### Focused Analysis Profile Card</div>", unsafe_allow_html=True)
        
        # Double-check guardrail to ensure the active selection is contained in currently filtered matched rows
        if st.session_state.selected_uid in matched_rows['System UID'].values:
            row = matched_rows[matched_rows['System UID'] == st.session_state.selected_uid].iloc[0]
        else:
            row = matched_rows.iloc[0]
            st.session_state.selected_uid = row['System UID']
            
        uid = row['System UID']
        cayuse_id = row['Parent Proposal Number']
        oracle_id = row['Oracle Award Number']
        
        st.markdown(f"#### Profile Target: {row['Award Name']}")
        
        # Identification Parameters Grid
        col1, col2, col3, col4 = st.columns(4)
        col1.markdown(f"**System UID:** `{uid}`")
        col2.markdown(f"**Parent Proposal (Cayuse):** `{cayuse_id}`")
        col3.markdown(f"**Oracle Award ID:** `{oracle_id if oracle_id else 'TBD'}`")
        col4.markdown(f"**Review Status:** <span class='status-badge'>{row['Review Status']}</span>", unsafe_allow_html=True)
        
        # Financial Double-Entry Accounting Metrics Grid
        st.markdown("<br>", unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric(label="Total Obligated Funding", value=f"${row['Total Funding Obligated']:,}")
        with m2:
            st.metric(label="Total Stated Awarded Budget", value=f"${row['Total Stated Awarded Budget']:,}")
        with m3:
            st.metric(label="Direct Cost (Obligated)", value=f"${row['Direct Funding Obligated']:,}")
        with m4:
            st.metric(label="Indirect Cost (Obligated)", value=f"${row['Indirect Funding Obligated']:,}")
            
        # Programmatic Balance Sheet Discrepancy Callout
        if str(row['Discrepancy Summary']).strip() and str(row['Discrepancy Summary']).lower() != "not found":
            st.markdown(f"<div class='alert-badge'><b>System Audit Notice:</b><br>{row['Discrepancy Summary']}</div>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # Relational Sub-Table Data Views
        with st.expander("View Complete Relational Data Sheets", expanded=True):
            tab_profile, tab_deliverables, tab_ledger, tab_labor, tab_subawards = st.tabs([
                "General Profile", "Tracking Deliverables", "Master Budget Ledger", "Labor Distribution", "Subrecipients"
            ])
            
            with tab_profile:
                p1, p2, p3 = st.columns(3)
                p1.write(f"**Principal Investigator:** {row['Principal Investigator']}")
                p2.write(f"**Sponsor:** {row['Primary Sponsor']}")
                p3.write(f"**Award Owning Org:** {row['Award Owning Organization']}")
                
                p4, p5, p6 = st.columns(3)
                p4.write(f"**Project Start Date:** {row['Start Date']}")
                p5.write(f"**Project End Date:** {row['End Date']}")
                p6.write(f"**Sponsor Award Number:** `{row['Sponsor Award Number']}`")
                
                p7, p8, p9 = st.columns(3)
                p7.write(f"**Award Purpose:** {row['Award Purpose']}")
                p8.write(f"**Category Type:** {row['Award Type']}")
                p9.write(f"**Billing Method:** {row['Bill Type']}")

            with tab_deliverables:
                df_deliv = data_pools["Deliverables_Detail"]
                if not df_deliv.empty and 'System UID' in df_deliv.columns:
                    matched_deliv = df_deliv[df_deliv['System UID'] == uid].drop(columns=['System UID', 'Parent Proposal Number'], errors='ignore')
                    st.dataframe(matched_deliv, width="stretch", hide_index=True)
                else:
                    st.info("No tracking milestones logged for this award lineage.")

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
        st.markdown("---")
else:
    st.info("The application portfolio database is currently unpopulated. Stage award timeline files in the control panel to initialize the reviewer screen.")