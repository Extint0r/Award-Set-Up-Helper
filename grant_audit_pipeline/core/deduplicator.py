from typing import List, Dict, Any, Tuple, Set

def deduplicate_and_merge_sources(extracted_docs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Merges OSP and Oracle Post-Award document instances, assigning ORIGIN_STATUS and CROSS_REF_PATH.
    """
    merged_docs: List[Dict[str, Any]] = []
    
    # Tier 1: Group by Exact MD5 Binary Hash
    hash_groups: Dict[str, List[Dict[str, Any]]] = {}
    unhashed_docs = []
    
    for doc in extracted_docs:
        h = doc.get("FILE_HASH", "")
        if h:
            hash_groups.setdefault(h, []).append(doc)
        else:
            unhashed_docs.append(doc)

    processed_hashes: Set[str] = set()

    for h, group in hash_groups.items():
        processed_hashes.add(h)
        sources = {d["SOURCE_TAG"] for d in group}
        
        primary_doc = group[0].copy()
        
        if len(sources) > 1 or len(group) > 1:
            primary_doc["ORIGIN_STATUS"] = "PRESENT_IN_BOTH" if len(sources) > 1 else f"{list(sources)[0]}_DUPLICATE"
            secondary_doc = next((d for d in group if d["PDF_Path"] != primary_doc["PDF_Path"]), None)
            primary_doc["CROSS_REF_PATH"] = secondary_doc["PDF_Path"] if secondary_doc else "N/A"
        else:
            primary_doc["ORIGIN_STATUS"] = f"{primary_doc['SOURCE_TAG']}_ONLY"
            primary_doc["CROSS_REF_PATH"] = "N/A"
            
        merged_docs.append(primary_doc)

    # Tier 2: Metadata Fingerprint Alignment
    fingerprints: Dict[Tuple, List[Dict[str, Any]]] = {}
    final_docs = []
    
    for doc in merged_docs:
        if doc["ORIGIN_STATUS"].endswith("_ONLY"):
            fp = (
                doc["AWARD_CLUSTER_KEY"], 
                round(doc["DELTA_OBLIGATED"], 2), 
                doc["EXECUTION_DATE"], 
                doc["DOC_START_DATE"]
            )
            if doc["AWARD_CLUSTER_KEY"] != "UNKNOWN" and doc["DELTA_OBLIGATED"] > 0:
                fingerprints.setdefault(fp, []).append(doc)
            else:
                final_docs.append(doc)
        else:
            final_docs.append(doc)

    for fp, group in fingerprints.items():
        sources = {d["SOURCE_TAG"] for d in group}
        if len(sources) > 1:
            primary_doc = group[0].copy()
            primary_doc["ORIGIN_STATUS"] = "PRESENT_IN_BOTH"
            secondary_doc = next((d for d in group if d["PDF_Path"] != primary_doc["PDF_Path"]), None)
            primary_doc["CROSS_REF_PATH"] = secondary_doc["PDF_Path"] if secondary_doc else "N/A"
            final_docs.append(primary_doc)
        else:
            final_docs.extend(group)

    return final_docs