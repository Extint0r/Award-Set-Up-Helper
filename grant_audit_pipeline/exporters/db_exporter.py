import sqlite3
import json
import pandas as pd
from pathlib import Path
from typing import List, Dict, Any

def sanitize_complex_types_for_sqlite(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts lists, dicts, and complex objects (e.g., DATE_SNIPPETS, BUDGET_PAGES)
    into clean JSON strings for seamless SQLite text storage.
    """
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].apply(lambda x: isinstance(x, (list, dict))).any():
            df_clean[col] = df_clean[col].apply(
                lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else ("" if pd.isna(x) else str(x))
            )
    return df_clean

def export_to_sqlite(merged_docs: List[Dict[str, Any]], recon_results: List[Dict[str, Any]], db_path: Path):
    """Initializes SQLite schema and stages document detail and reconciliation summaries."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    
    doc_df   = pd.DataFrame(merged_docs)
    recon_df = pd.DataFrame(recon_results)

    # 1. Sanitize complex nested structures (DATE_SNIPPETS, BUDGET_PAGES, etc.) to JSON strings
    doc_df   = sanitize_complex_types_for_sqlite(doc_df)
    recon_df = sanitize_complex_types_for_sqlite(recon_df)

    # 2. Format dates to string for SQLite compatibility
    date_cols_doc = [
        "EXECUTION_DATE", "DOC_START_DATE", "DOC_END_DATE", 
        "VISION_EXECUTION_DATE", "VISION_BUDGET_START", "VISION_BUDGET_END", 
        "VISION_PROJECT_START", "VISION_PROJECT_END"
    ]
    for col in date_cols_doc:
        if col in doc_df.columns:
            doc_df[col] = doc_df[col].astype(str)
            
    for col in ["PDF_START_DATE_TRUTH", "PDF_END_DATE_TRUTH"]:
        if col in recon_df.columns:
            recon_df[col] = recon_df[col].astype(str)

    # 3. Drop heavy raw text body and internal tracking hashes before DB export
    cols_to_drop = ["FILE_HASH", "SOURCE_TAG", "RAW_CAYUSE_PROJ", "RAW_CAYUSE_PROP", "RAW_ORACLE_NUM", "RAW_BANNER_UID", "TEXT_BODY"]
    doc_db_df = doc_df.drop(columns=[c for c in cols_to_drop if c in doc_df.columns])

    # 4. Safe connection context manager for database writing
    with sqlite3.connect(db_path) as conn:
        doc_db_df.to_sql("tbl_Documents", conn, if_exists="replace", index=False)
        recon_df.to_sql("tbl_Reconciliation_Summary", conn, if_exists="replace", index=False)