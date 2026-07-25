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
    if df.empty:
        return df
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].apply(lambda x: isinstance(x, (list, dict))).any():
            df_clean[col] = df_clean[col].apply(
                lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else ("" if pd.isna(x) else str(x))
            )
    return df_clean

def export_to_sqlite(
    headers: List[Dict[str, Any]], 
    transactions: List[Dict[str, Any]], 
    recon_results: List[Dict[str, Any]], 
    db_path: Path
):
    """Initializes SQLite schema and stages Award Headers, Transaction Ledger, and Reconciliation Summaries."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    
    header_df = pd.DataFrame(headers)
    tx_df     = pd.DataFrame(transactions)
    recon_df  = pd.DataFrame(recon_results)

    # 1. Sanitize complex nested structures
    header_df = sanitize_complex_types_for_sqlite(header_df)
    tx_df     = sanitize_complex_types_for_sqlite(tx_df)
    recon_df  = sanitize_complex_types_for_sqlite(recon_df)

    # 2. Format dates to string for SQLite compatibility
    date_cols_tx = ["ACTION_DATE", "BUDGET_PERIOD_START", "BUDGET_PERIOD_END"]
    for col in date_cols_tx:
        if col in tx_df.columns:
            tx_df[col] = tx_df[col].astype(str)
            
    for col in ["PDF_START_DATE_TRUTH", "PDF_END_DATE_TRUTH"]:
        if col in recon_df.columns:
            recon_df[col] = recon_df[col].astype(str)
        if col in header_df.columns:
            header_df[col] = header_df[col].astype(str)

    # 3. Write relational tables to SQLite database
    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA foreign_keys = ON;")
        header_df.to_sql("tbl_Award_Header", conn, if_exists="replace", index=False)
        tx_df.to_sql("tbl_Award_Transactions", conn, if_exists="replace", index=False)
        recon_df.to_sql("tbl_Reconciliation_Summary", conn, if_exists="replace", index=False)