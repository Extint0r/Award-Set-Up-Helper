import re
import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Optional
from config import TRIAGE_EXCEL_PATH

# Global Lookup Maps
LOOKUP_PROJ_TO_ORACLE: Dict[str, str] = {}
LOOKUP_PROP_TO_ORACLE: Dict[str, str] = {}
LOOKUP_BANNER_TO_ORACLE: Dict[str, str] = {}
LOOKUP_BANNER_TO_PROJ: Dict[str, str] = {}
LOOKUP_ORACLE_TO_PROJ: Dict[str, str] = {}
TRIAGE_METRICS: Dict[str, Dict[str, float]] = {}


def clean_id_str(val: Any) -> str:
    """Helper to stringify, clean whitespace, and strip float trailing '.0' from Excel values."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "" if s.lower() in ("nan", "none", "null") else s


def is_valid_oracle_id(val: str) -> bool:
    """Oracle Award Numbers must strictly be 6 contiguous numeric digits (100000 - 999999)."""
    if not val:
        return False
    return bool(re.match(r'^[1-9]\d{5}$', val.strip()))


def find_triage_columns(cols: List[str]) -> Dict[str, Optional[str]]:
    """Flexibly maps column names from Triage sheet regardless of casing/formatting."""
    cols_upper = [str(c).strip().upper() for c in cols]
    
    col_orc_num, col_cay_proj, col_banner_uid = None, None, None
    col_cay_ceil, col_cay_ob, col_orc_ceil, col_orc_ob = None, None, None, None

    # Step 1: Look for exact standard names first
    for original, c in zip(cols, cols_upper):
        if c in ["ORACLE AWARD NUMBER", "ORACLE_AWARD_NUMBER", "ORACLE AWARD ID", "ORACLE_AWARD_ID", "ORACLE ID", "ORACLE_ID"]:
            col_orc_num = original
        elif c in ["CAYUSE PROJECT NUMBER", "CAYUSE_PROJECT_NUMBER", "CAYUSE PROJECT ID", "CAYUSE_PROJECT_ID", "CAYUSE PROJECT", "CAYUSE_PROJECT"]:
            col_cay_proj = original
        elif c in ["BANNER AWARD NUMBER", "BANNER_AWARD_NUMBER", "BANNER FUND NUMBER", "BANNER_FUND_NUMBER", "BANNER ID", "BANNER_ID", "BANNER UID", "BANNER_UID"]:
            col_banner_uid = original

    # Step 2: Fallback matching with explicit exclusions for STATUS / STATE / TYPE
    for original, c in zip(cols, cols_upper):
        if not col_orc_num and "ORACLE" in c and ("AWARD" in c or "NUMBER" in c or "ID" in c) and not any(x in c for x in ["LIMIT", "CEILING", "BURDENED", "COST", "OBLIGAT", "STATUS", "STATE", "TYPE", "TITLE", "NAME"]):
            col_orc_num = original
        elif not col_cay_proj and "CAYUSE" in c and ("PROJECT" in c or "PROJ" in c or "PROPOSAL" in c) and not any(x in c for x in ["TOTAL", "AMOUNT", "OBLIGAT", "CEILING", "STATUS", "STATE", "TYPE"]):
            col_cay_proj = original
        elif not col_banner_uid and ("BANNER" in c or "FUND" in c) and ("NUMBER" in c or "ID" in c or "NO" in c or "UID" in c or "CODE" in c) and not any(x in c for x in ["STATUS", "STATE", "TYPE"]):
            col_banner_uid = original
        elif "CAYUSE" in c and ("TOTAL" in c or "AMOUNT" in c or "CEILING" in c) and "OBLIGAT" not in c and "PROJECT" not in c:
            col_cay_ceil = original
        elif "CAYUSE" in c and "OBLIGAT" in c:
            if "TOTAL" in c or col_cay_ob is None:
                col_cay_ob = original
        elif "ORACLE" in c and ("LIMIT" in c or "HARD" in c or "CEILING" in c or "HEADER" in c) and "COST" not in c and "BURDENED" not in c:
            col_orc_ceil = original
        elif "ORACLE" in c and ("BURDENED" in c or "COST" in c or "OBLIGAT" in c or "BUDGET" in c):
            col_orc_ob = original

    return {
        "ORACLE_AWARD_NUMBER": col_orc_num,
        "CAYUSE_PROJECT_NUMBER": col_cay_proj,
        "BANNER_AWARD_NUMBER": col_banner_uid,
        "CAYUSE_CEILING": col_cay_ceil,
        "CAYUSE_OBLIGATED": col_cay_ob,
        "ORACLE_CEILING": col_orc_ceil,
        "ORACLE_OBLIGATED": col_orc_ob
    }


def build_dynamic_crosswalk(extracted_docs: List[Dict[str, Any]], triage_path: Optional[Path] = None):
    """
    3-Stage Dynamic Crosswalk Engine:
    1. Ingests master baseline linkages from Triage Excel file ('MASTER' tab).
    2. Learns cross-references discovered dynamically from PDF filenames/metadata.
    3. Resolves cluster keys and infers missing Cayuse Project Numbers via Banner IDs.
    """
    global LOOKUP_PROJ_TO_ORACLE, LOOKUP_PROP_TO_ORACLE, LOOKUP_BANNER_TO_ORACLE
    global LOOKUP_BANNER_TO_PROJ, LOOKUP_ORACLE_TO_PROJ, TRIAGE_METRICS

    LOOKUP_PROJ_TO_ORACLE.clear()
    LOOKUP_PROP_TO_ORACLE.clear()
    LOOKUP_BANNER_TO_ORACLE.clear()
    LOOKUP_BANNER_TO_PROJ.clear()
    LOOKUP_ORACLE_TO_PROJ.clear()
    TRIAGE_METRICS.clear()

    active_triage_path = triage_path or TRIAGE_EXCEL_PATH

    # STAGE 1: Ingest Triage Excel File
    if active_triage_path and active_triage_path.exists():
        try:
            excel_file = pd.ExcelFile(active_triage_path)
            sheet_target = "MASTER" if "MASTER" in [s.upper() for s in excel_file.sheet_names] else excel_file.sheet_names[0]
            actual_sheet_name = next(s for s in excel_file.sheet_names if s.upper() == sheet_target.upper())
            df_triage = pd.read_excel(excel_file, sheet_name=actual_sheet_name)

            col_map = find_triage_columns(df_triage.columns.tolist())
            c_orc_id = col_map["ORACLE_AWARD_NUMBER"]
            c_cay_id = col_map["CAYUSE_PROJECT_NUMBER"]
            c_ban_id = col_map["BANNER_AWARD_NUMBER"]

            for _, row in df_triage.iterrows():
                oracle_id   = clean_id_str(row.get(c_orc_id)) if c_orc_id else ""
                cayuse_proj = clean_id_str(row.get(c_cay_id)) if c_cay_id else ""
                banner_id   = clean_id_str(row.get(c_ban_id)).upper() if c_ban_id else ""

                # Reject non-6-digit Oracle values (e.g., 'Active', 'Closed')
                if not is_valid_oracle_id(oracle_id):
                    oracle_id = ""

                if oracle_id and cayuse_proj:
                    LOOKUP_PROJ_TO_ORACLE[cayuse_proj] = oracle_id
                    LOOKUP_ORACLE_TO_PROJ[oracle_id] = cayuse_proj
                if banner_id and cayuse_proj:
                    LOOKUP_BANNER_TO_PROJ[banner_id] = cayuse_proj
                if banner_id and oracle_id:
                    LOOKUP_BANNER_TO_ORACLE[banner_id] = oracle_id

        except Exception as e:
            print(f"[WARNING] Could not read Triage sheet: {e}")

    # STAGE 2: Learn Discovered Links from PDF Filenames
    for doc in extracted_docs:
        r_proj   = clean_id_str(doc.get("RAW_CAYUSE_PROJ"))
        r_prop   = clean_id_str(doc.get("RAW_CAYUSE_PROP"))
        r_oracle = clean_id_str(doc.get("RAW_ORACLE_NUM"))
        r_banner = clean_id_str(doc.get("RAW_BANNER_UID")).upper()

        if not is_valid_oracle_id(r_oracle):
            r_oracle = ""

        if r_banner:
            if r_proj:
                LOOKUP_BANNER_TO_PROJ[r_banner] = r_proj
            if r_oracle:
                LOOKUP_BANNER_TO_ORACLE[r_banner] = r_oracle
        if r_oracle:
            if r_proj:
                LOOKUP_PROJ_TO_ORACLE[r_proj] = r_oracle
                LOOKUP_ORACLE_TO_PROJ[r_oracle] = r_proj
            if r_prop:
                LOOKUP_PROP_TO_ORACLE[r_prop] = r_oracle

    # STAGE 3: Re-resolve Missing Identifiers & Assign Cluster Key
    for doc in extracted_docs:
        r_proj   = clean_id_str(doc.get("RAW_CAYUSE_PROJ"))
        r_prop   = clean_id_str(doc.get("RAW_CAYUSE_PROP"))
        r_oracle = clean_id_str(doc.get("RAW_ORACLE_NUM"))
        r_banner = clean_id_str(doc.get("RAW_BANNER_UID")).upper()

        if not is_valid_oracle_id(r_oracle):
            r_oracle = ""

        resolved_oracle = (
            r_oracle or 
            LOOKUP_PROJ_TO_ORACLE.get(r_proj, "") or 
            LOOKUP_PROP_TO_ORACLE.get(r_prop, "") or
            LOOKUP_BANNER_TO_ORACLE.get(r_banner, "")
        )
        if not is_valid_oracle_id(resolved_oracle):
            resolved_oracle = ""

        resolved_cayuse_proj = (
            r_proj or 
            LOOKUP_ORACLE_TO_PROJ.get(resolved_oracle, "") or
            LOOKUP_BANNER_TO_PROJ.get(r_banner, "")
        )

        doc["ORACLE_AWARD_NUMBER"] = resolved_oracle
        doc["CAYUSE_PROJECT_NUMBER"] = resolved_cayuse_proj
        doc["BANNER_AWARD_NUMBER"] = r_banner

        # Hierarchy: Oracle -> Cayuse Project -> Banner ID -> UNKNOWN
        cluster_key = resolved_oracle or resolved_cayuse_proj or r_banner or "UNKNOWN"
        doc["AWARD_CLUSTER_KEY"] = cluster_key

        if resolved_oracle:
            doc["MATCH_CONFIDENCE"] = "Active Index Match"
        elif resolved_cayuse_proj:
            doc["MATCH_CONFIDENCE"] = "Cayuse Inference" if r_proj else "Banner-to-Cayuse Inference"
        elif r_banner:
            doc["MATCH_CONFIDENCE"] = "Banner ID Cluster"
        else:
            doc["MATCH_CONFIDENCE"] = "Unmatched Baseline"