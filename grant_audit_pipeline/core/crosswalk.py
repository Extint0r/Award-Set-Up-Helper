from typing import List, Dict, Any

LOOKUP_PROJ_TO_ORACLE: Dict[str, str] = {}
LOOKUP_PROP_TO_ORACLE: Dict[str, str] = {}
LOOKUP_BANNER_TO_ORACLE: Dict[str, str] = {}
LOOKUP_ORACLE_TO_PROJ: Dict[str, str] = {}

def build_dynamic_crosswalk(extracted_docs: List[Dict[str, Any]]):
    """
    2-Pass Crosswalk Engine: Unifies split award clusters across the entire batch
    using Oracle IDs, Cayuse Projects, Proposal IDs, and Banner UIDs.
    """
    global LOOKUP_PROJ_TO_ORACLE, LOOKUP_PROP_TO_ORACLE, LOOKUP_BANNER_TO_ORACLE, LOOKUP_ORACLE_TO_PROJ
    
    # Pass 1: Learn all inter-identifier links from files containing multiple IDs
    for doc in extracted_docs:
        r_proj   = doc.get("RAW_CAYUSE_PROJ", "")
        r_prop   = doc.get("RAW_CAYUSE_PROP", "")
        r_oracle = doc.get("RAW_ORACLE_NUM", "")
        r_banner = doc.get("RAW_BANNER_UID", "")
        
        if r_oracle:
            if r_proj:
                LOOKUP_PROJ_TO_ORACLE[r_proj] = r_oracle
                LOOKUP_ORACLE_TO_PROJ[r_oracle] = r_proj
            if r_prop:
                LOOKUP_PROP_TO_ORACLE[r_prop] = r_oracle
            if r_banner:
                LOOKUP_BANNER_TO_ORACLE[r_banner] = r_oracle

    # Pass 2: Re-resolve missing identifiers across all records
    for doc in extracted_docs:
        r_proj   = doc.get("RAW_CAYUSE_PROJ", "")
        r_prop   = doc.get("RAW_CAYUSE_PROP", "")
        r_oracle = doc.get("RAW_ORACLE_NUM", "")
        r_banner = doc.get("RAW_BANNER_UID", "")
        
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
        doc["AWARD_CLUSTER_KEY"] = resolved_oracle or resolved_cayuse_proj or r_banner or "UNKNOWN"

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