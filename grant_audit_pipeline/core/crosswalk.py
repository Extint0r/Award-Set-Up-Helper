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


def build_dynamic_crosswalk(extracted_docs: List[Dict[str, Any]], triage_path: Optional[Path] = None):
    """
    3-Stage Dynamic Crosswalk Engine:
    1. Ingests master crosswalk linkages and ledger balances from MASTER tab of Triage file.
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
    # STAGE 1: Ingest MASTER Tab from Triage Excel File
    # =========================================================================
    if active_triage_path and active_triage_path.exists():
        try:
            df_triage = pd.read_excel(active_triage_path, sheet_name="MASTER")
            df_triage.columns = [str(c).strip().upper() for c in df_triage.columns]

            for _, row in df_triage.iterrows():
                oracle_id   = clean_id_str(row.get("ORACLE_AWARD_NUMBER"))
                cayuse_proj = clean_id_str(row.get("CAYUSE_PROJECT_NUMBER"))

                # Crosswalk mappings
                if oracle_id and cayuse_proj:
                    LOOKUP_PROJ_TO_ORACLE[cayuse_proj] = oracle_id
                    LOOKUP_ORACLE_TO_PROJ[oracle_id] = cayuse_proj

                # Ingest exact mapped ledger columns from MASTER tab
                metrics = {
                    "CAYUSE_CEILING": pd.to_numeric(row.get("CAYUSE_TOTAL_AMOUNT"), errors='coerce') or 0.0,
                    "CAYUSE_OBLIGATED": pd.to_numeric(row.get("CAYUSE_OBLIGATED_TOTAL"), errors='coerce') or 0.0,
                    "ORACLE_CEILING": pd.to_numeric(row.get("ORACLE_HEADER_HARD_LIMIT"), errors='coerce') or 0.0,
                    "ORACLE_OBLIGATED": pd.to_numeric(row.get("ORACLE_BUDGET_BURDENED_COST"), errors='coerce') or 0.0,
                }

                if oracle_id:
                    TRIAGE_METRICS[oracle_id] = metrics
                if cayuse_proj:
                    TRIAGE_METRICS[cayuse_proj] = metrics

            print(f"Ingested {len(df_triage)} master records from Triage sheet: {active_triage_path.name}")
        except Exception as e:
            print(f"[WARNING] Could not read MASTER tab from {active_triage_path}: {e}")
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
        doc["CAYUSE_CEILING"]   = ledger_data.get("CAYUSE_CEILING", 0.0)
        doc["CAYUSE_OBLIGATED"] = ledger_data.get("CAYUSE_OBLIGATED", 0.0)
        doc["ORACLE_CEILING"]   = ledger_data.get("ORACLE_CEILING", 0.0)
        doc["ORACLE_OBLIGATED"] = ledger_data.get("ORACLE_OBLIGATED", 0.0)

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