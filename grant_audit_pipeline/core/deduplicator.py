import hashlib
from pathlib import Path
from typing import List, Dict, Any, Tuple, Set


def compute_file_hash(pdf_path: Path) -> str:
    """Computes MD5 hash for exact binary duplicate matching."""
    hasher = hashlib.md5()
    try:
        with open(pdf_path, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception:
        return ""


def stage0_binary_preflight(
    discovered_pdfs: List[Tuple[Path, str]]
) -> Tuple[List[Tuple[Path, str]], Dict[str, List[Tuple[Path, str]]]]:
    """
    STAGE 0: Binary Pre-Flight Optimization
    Calculates MD5 hashes across all discovered PDFs before Pass 1 parsing.
    
    Returns:
        unique_pdfs: List of (pdf_path, source_tag) for unique files to parse
        binary_duplicate_map: Dict mapping MD5 hash -> List of duplicate (pdf_path, source_tag)
    """
    seen_hashes: Dict[str, Tuple[Path, str]] = {}
    unique_pdfs: List[Tuple[Path, str]] = []
    binary_duplicate_map: Dict[str, List[Tuple[Path, str]]] = {}

    for pdf_path, source_tag in discovered_pdfs:
        file_hash = compute_file_hash(pdf_path)
        if not file_hash:
            # Fallback if unhashable
            unique_pdfs.append((pdf_path, source_tag))
            continue

        if file_hash not in seen_hashes:
            seen_hashes[file_hash] = (pdf_path, source_tag)
            unique_pdfs.append((pdf_path, source_tag))
        else:
            # Record duplicate instance for metadata cloning
            binary_duplicate_map.setdefault(file_hash, []).append((pdf_path, source_tag))

    return unique_pdfs, binary_duplicate_map


def clone_binary_duplicate_metadata(
    extracted_unique_docs: List[Dict[str, Any]], 
    binary_duplicate_map: Dict[str, List[Tuple[Path, str]]]
) -> List[Dict[str, Any]]:
    """
    Expands binary duplicate instances without re-running OCR/Vision parsing.
    Clones parsed metadata from primary instance and attaches ORIGIN_STATUS and CROSS_REF_PATH.
    """
    all_docs: List[Dict[str, Any]] = []

    for doc in extracted_unique_docs:
        h = doc.get("FILE_HASH", "")
        duplicates = binary_duplicate_map.get(h, [])

        if not duplicates:
            doc["ORIGIN_STATUS"] = f"{doc.get('SOURCE_TAG', 'UNKNOWN')}_ONLY"
            doc["CROSS_REF_PATH"] = "N/A"
            all_docs.append(doc)
        else:
            primary_source = doc.get("SOURCE_TAG", "UNKNOWN")
            dup_sources = {d[1] for d in duplicates}
            all_sources = {primary_source}.union(dup_sources)

            if len(all_sources) > 1:
                doc["ORIGIN_STATUS"] = "PRESENT_IN_BOTH"
            else:
                doc["ORIGIN_STATUS"] = f"{primary_source}_DUPLICATE"

            sec_path = str(duplicates[0][0])
            doc["CROSS_REF_PATH"] = sec_path
            all_docs.append(doc)

            # Clone secondary duplicate records for full audit ledger tracking
            for dup_path, dup_source in duplicates:
                dup_doc = doc.copy()
                dup_doc["PDF_Path"] = str(dup_path)
                dup_doc["Filename"] = dup_path.name
                dup_doc["SOURCE_TAG"] = dup_source
                dup_doc["ORIGIN_STATUS"] = "PRESENT_IN_BOTH" if len(all_sources) > 1 else f"{dup_source}_DUPLICATE"
                dup_doc["CROSS_REF_PATH"] = str(doc.get("PDF_Path", "N/A"))
                all_docs.append(dup_doc)

    return all_docs


def deduplicate_and_merge_sources(extracted_docs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Backwards-compatible wrapper that handles Stage 0 binary hashing and Tier 1 deduplication cleanly.
    """
    hash_groups: Dict[str, List[Dict[str, Any]]] = {}
    for doc in extracted_docs:
        h = doc.get("FILE_HASH", "")
        if h:
            hash_groups.setdefault(h, []).append(doc)
        else:
            hash_groups.setdefault(doc.get("Filename", ""), []).append(doc)

    merged_docs: List[Dict[str, Any]] = []
    for h, group in hash_groups.items():
        primary_doc = group[0].copy()
        sources = {d["SOURCE_TAG"] for d in group}
        
        if len(sources) > 1 or len(group) > 1:
            primary_doc["ORIGIN_STATUS"] = "PRESENT_IN_BOTH" if len(sources) > 1 else f"{primary_doc['SOURCE_TAG']}_DUPLICATE"
            sec = next((d for d in group if d.get("PDF_Path") != primary_doc.get("PDF_Path")), None)
            primary_doc["CROSS_REF_PATH"] = sec.get("PDF_Path", "N/A") if sec else "N/A"
        else:
            primary_doc["ORIGIN_STATUS"] = f"{primary_doc['SOURCE_TAG']}_ONLY"
            primary_doc["CROSS_REF_PATH"] = "N/A"
            
        merged_docs.append(primary_doc)

    return merged_docs