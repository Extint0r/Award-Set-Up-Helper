import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Optional
from config import TRIAGE_EXCEL_PATH

# Global Lookup Maps for Crosswalk & Ledger Reconciliation
LOOKUP_PROJ_TO_ORACLE: Dict[str, str] = {}
LOOKUP_PROP_TO_ORACLE: Dict[str, str] = {}
LOOKUP_BANNER_TO_ORACLE: Dict[str, str] = {}
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


def find_triage_columns(cols: List[str]) -> Dict[str, Optional[str]]:
    """
    Flexibly maps column names from Triage sheet (MasterBaselineSheet table)
    regardless of exact casing or formatting variations.
    """
    cols_upper = [str(c).strip().upper() for c in cols]
    
    col_orc_num, col_cay_proj, col_cay_ceil, col_cay_ob, col_orc_ceil, col_orc_ob = None, None, None, None, None, None

    for original, c in zip(cols, cols_upper):
        if "ORACLE" in c and ("AWARD" in c or "NUMBER" in c or "ID" in c) and not any(x in c for x in ["LIMIT", "CEILING", "BURDENED", "COST", "OBLIGAT"]):
            col_orc_num = original
        elif "CAYUSE" in c and ("PROJECT" in c or "PROJ" in c or "PROPOSAL" in c) and not any(x in c for x in ["TOTAL", "AMOUNT", "OBLIGAT", "CEILING"]):
            col_cay_proj = original
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
        "CAYUSE_CEILING": col_cay_ceil,
        "CAYUSE_OBLIGATED": col_cay_ob,
        "ORACLE_CEILING": col_orc_ceil,
        "ORACLE_OBLIGATED": col_orc_ob
    }


def build_dynamic_crosswalk(extracted_docs: List[Dict[str, Any]], triage_path: Optional[Path] = None):
    """
    3-Stage Dynamic Crosswalk Engine:
    1. Ingests master crosswalk linkages and ledger balances from Triage Excel file (MASTER tab / MasterBaselineSheet table).
    2. Learns additional cross-references discovered dynamically from PDF filenames/metadata.
    3. Resolves cluster keys and attaches Cayuse/Oracle ledger metrics to all document records.
    """
    global LOOKUP_PROJ_TO_ORACLE, LOOKUP_PROP_TO_ORACLE, LOOKUP_BANNER_TO_ORACLE
    global LOOKUP_ORACLE_TO_PROJ, TRIAGE_METRICS

    # Clear maps on re-run
    LOOKUP_PROJ_TO_ORACLE.clear()
    LOOKUP_PROP_TO_ORACLE.clear()
    LOOKUP_BANNER_TO_ORACLE.clear()
    LOOKUP_ORACLE_TO_PROJ.clear()
    TRIAGE_METRICS.clear()

    active_triage_path = triage_path or TRIAGE_EXCEL_PATH

    # =========================================================================
    # STAGE 1: Ingest Triage Excel File (Master Baseline Sheet)
    # =========================================================================
    if active_triage_path and active_triage_path.exists():
        try:
            excel_file = pd.ExcelFile(active_triage_path)
            sheet_target = "MASTER" if "MASTER" in [s.upper() for s in excel_file.sheet_names] else excel_file.sheet_names[0]
            
            actual_sheet_name = next(s for s in excel_file.sheet_names if s.upper() == sheet_target.upper())
            df_triage = pd.read_excel(excel_file, sheet_name=actual_sheet_name)

            col_map = find_triage_columns(df_triage.columns.tolist())
            print(f"Ingesting Triage Sheet ('{actual_sheet_name}') with mapped columns: {col_map}")

            c_orc_id = col_map["ORACLE_AWARD_NUMBER"]
            c_cay_id = col_map["CAYUSE_PROJECT_NUMBER"]
            c_cay_c  = col_map["CAYUSE_CEILING"]
            c_cay_o  = col_map["CAYUSE_OBLIGATED"]
            c_orc_c  = col_map["ORACLE_CEILING"]
            c_orc_o  = col_map["ORACLE_OBLIGATED"]

            for _, row in df_triage.iterrows():
                oracle_id   = clean_id_str(row.get(c_orc_id)) if c_orc_id else ""
                cayuse_proj = clean_id_str(row.get(c_cay_id)) if c_cay_id else ""

                # Crosswalk mappings
                if oracle_id and cayuse_proj:
                    LOOKUP_PROJ_TO_ORACLE[cayuse_proj] = oracle_id
                    LOOKUP_ORACLE_TO_PROJ[oracle_id] = cayuse_proj

                # Ingest exact mapped ledger columns
                metrics = {
                    "CAYUSE_CEILING": pd.to_numeric(row.get(c_cay_c), errors='coerce') or 0.0 if c_cay_c else 0.0,
                    "CAYUSE_OBLIGATED": pd.to_numeric(row.get(c_cay_o), errors='coerce') or 0.0 if c_cay_o else 0.0,
                    "ORACLE_CEILING": pd.to_numeric(row.get(c_orc_c), errors='coerce') or 0.0 if c_orc_c else 0.0,
                    "ORACLE_OBLIGATED": pd.to_numeric(row.get(c_orc_o), errors='coerce') or 0.0 if c_orc_o else 0.0,
                }

                if oracle_id:
                    TRIAGE_METRICS[oracle_id] = metrics
                if cayuse_proj:
                    TRIAGE_METRICS[cayuse_proj] = metrics

            print(f"Successfully ingested {len(df_triage)} master baseline records from {active_triage_path.name}")
        except Exception as e:
            print(f"[WARNING] Could not read Triage sheet from {active_triage_path}: {e}")
    else:
        print(f"[WARNING] Triage file not found at: {active_triage_path}")

    # =========================================================================
    # STAGE 2: Learn Discovered Links from PDF Filenames / Pass 1 Metadata
    # =========================================================================
    for doc in extracted_docs:
        r_proj   = clean_id_str(doc.get("RAW_CAYUSE_PROJ"))
        r_prop   = clean_id_str(doc.get("RAW_CAYUSE_PROP"))
        r_oracle = clean_id_str(doc.get("RAW_ORACLE_NUM"))
        r_banner = clean_id_str(doc.get("RAW_BANNER_UID"))

        if r_oracle:
            if r_proj:
                LOOKUP_PROJ_TO_ORACLE[r_proj] = r_oracle
                LOOKUP_ORACLE_TO_PROJ[r_oracle] = r_proj
            if r_prop:
                LOOKUP_PROP_TO_ORACLE[r_prop] = r_oracle
            if r_banner:
                LOOKUP_BANNER_TO_ORACLE[r_banner] = r_oracle

    # =========================================================================
    # STAGE 3: Re-resolve Missing Identifiers & Attach Ledger Metrics
    # =========================================================================
    for doc in extracted_docs:
        r_proj   = clean_id_str(doc.get("RAW_CAYUSE_PROJ"))
        r_prop   = clean_id_str(doc.get("RAW_CAYUSE_PROP"))
        r_oracle = clean_id_str(doc.get("RAW_ORACLE_NUM"))
        r_banner = clean_id_str(doc.get("RAW_BANNER_UID"))

        resolved_oracle = (
            r_oracle or 
            LOOKUP_PROJ_TO_ORACLE.get(r_proj, "") or 
            LOOKUP_PROP_TO_ORACLE.get(r_prop, "") or
            LOOKUP_BANNER_TO_ORACLE.get(r_banner, "")
        )

        resolved_cayuse_proj = (
            r_proj or 
            LOOKUP_ORACLE_TO_PROJ.get(resolved_oracle, "")
        )

        doc["ORACLE_AWARD_NUMBER"] = resolved_oracle
        doc["CAYUSE_PROJECT_NUMBER"] = resolved_cayuse_proj
        doc["CAYUSE_PROPOSAL_NUMBER"] = r_prop
        doc["BANNER_AWARD_UID"] = r_banner

        cluster_key = resolved_oracle or resolved_cayuse_proj or r_banner or "UNKNOWN"
        doc["AWARD_CLUSTER_KEY"] = cluster_key

        # Attach Triage Ledger Balances
        ledger_data = TRIAGE_METRICS.get(resolved_oracle) or TRIAGE_METRICS.get(resolved_cayuse_proj) or {}
        doc["CAYUSE_CEILING"]   = float(ledger_data.get("CAYUSE_CEILING", 0.0))
        doc["CAYUSE_OBLIGATED"] = float(ledger_data.get("CAYUSE_OBLIGATED", 0.0))
        doc["ORACLE_CEILING"]   = float(ledger_data.get("ORACLE_CEILING", 0.0))
        doc["ORACLE_OBLIGATED"] = float(ledger_data.get("ORACLE_OBLIGATED", 0.0))

        if resolved_oracle:
            doc["MATCH_CONFIDENCE"] = "Active Index Match"
            doc["MATCH_REASON"] = f"Matched Oracle Award Number ({resolved_oracle})"
        elif resolved_cayuse_proj:
            doc["MATCH_CONFIDENCE"] = "Cayuse Inference"
            doc["MATCH_REASON"] = f"Matched Cayuse Project Number ({resolved_cayuse_proj})"
        elif r_banner:
            doc["MATCH_CONFIDENCE"] = "Banner UID Inference"
            doc["MATCH_REASON"] = f"Matched Banner UID ({r_banner})"
        else:
            doc["MATCH_CONFIDENCE"] = "Unmatched Baseline"
            doc["MATCH_REASON"] = "No direct Oracle, Cayuse, or Banner ID in filename"