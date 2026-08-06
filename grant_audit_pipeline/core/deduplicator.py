import hashlib
import json
from pathlib import Path
from typing import List, Dict, Any, Tuple, Set
from concurrent.futures import ThreadPoolExecutor, as_completed

from config import REVIEW_DIR

CACHE_FILE = REVIEW_DIR / ".md5_file_cache.json"


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


def load_hash_cache() -> Dict[str, str]:
    """Loads cached MD5 hashes from disk to avoid re-reading USB files."""
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_hash_cache(cache: Dict[str, str]):
    """Persists MD5 hash cache to disk."""
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception as e:
        print(f"[Warning] Could not save MD5 cache: {e}")


def get_file_md5_fast(pdf_path: Path, cache: Dict[str, str]) -> str:
    """Computes MD5 hash with caching based on file path, size, and modification time."""
    try:
        stat = pdf_path.stat()
        cache_key = f"{pdf_path.resolve()}|{stat.st_size}|{stat.st_mtime}"
        
        if cache_key in cache:
            return cache[cache_key]

        md5_hash = compute_file_hash(pdf_path)
        if md5_hash:
            cache[cache_key] = md5_hash
        return md5_hash
    except Exception:
        return ""


def stage0_binary_preflight(
    discovered_pdfs: List[Tuple[Path, str]],
    max_workers: int = 16
) -> Tuple[List[Tuple[Path, str]], Dict[str, List[Tuple[Path, str]]]]:
    """
    High-Performance Stage 0 Binary Deduplication (26,000+ File Scale)
    
    1. Pre-groups by exact file size in bytes (skips hashing for unique sizes).
    2. Multi-threaded MD5 hashing for size collisions using persistent disk caching.
    3. Fully backward-compatible with clone_binary_duplicate_metadata().
    """
    print(f"\n--- STAGE 0: Fast Binary Pre-Flight ({len(discovered_pdfs):,} files) ---", flush=True)
    
    hash_cache = load_hash_cache()

    # 1. Fast Size-Based Grouping (No file content reading required)
    print("1/3 Pre-filtering by exact byte size...", flush=True)
    size_groups: Dict[int, List[Tuple[Path, str]]] = {}
    for path, tag in discovered_pdfs:
        try:
            sz = path.stat().st_size
            size_groups.setdefault(sz, []).append((path, tag))
        except Exception:
            continue

    unique_by_size = [items[0] for sz, items in size_groups.items() if len(items) == 1]
    collision_candidates = [item for sz, items in size_groups.items() if len(items) > 1 for item in items]

    print(f" -> Found {len(unique_by_size):,} files with unique byte sizes (MD5 skipped).")
    print(f" -> Found {len(collision_candidates):,} candidate files sharing duplicate byte sizes.")

    # 2. Multi-Threaded Hashing for Size Collisions
    unique_pdfs: List[Tuple[Path, str]] = list(unique_by_size)
    seen_hashes: Dict[str, Tuple[Path, str]] = {}
    binary_duplicate_map: Dict[str, List[Tuple[Path, str]]] = {}

    if collision_candidates:
        print(f"2/3 Hashing {len(collision_candidates):,} collision candidates across {max_workers} threads...", flush=True)
        
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {
                executor.submit(get_file_md5_fast, path, hash_cache): (path, tag)
                for path, tag in collision_candidates
            }
            
            completed = 0
            total_collisions = len(collision_candidates)
            for future in as_completed(future_to_file):
                path, tag = future_to_file[future]
                md5_hash = future.result()
                results.append((path, tag, md5_hash))
                
                completed += 1
                if completed % 1000 == 0 or completed == total_collisions:
                    print(f"    Progress: {completed:,} / {total_collisions:,} processed ({completed/total_collisions*100:.1f}%)", flush=True)

        for path, tag, md5_hash in results:
            if not md5_hash:
                unique_pdfs.append((path, tag))
                continue

            if md5_hash not in seen_hashes:
                seen_hashes[md5_hash] = (path, tag)
                unique_pdfs.append((path, tag))
            else:
                binary_duplicate_map.setdefault(md5_hash, []).append((path, tag))

    save_hash_cache(hash_cache)

    print(f"3/3 Pre-flight complete: {len(unique_pdfs):,} unique files identified ({len(discovered_pdfs) - len(unique_pdfs):,} duplicate instances filtered).\n", flush=True)
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