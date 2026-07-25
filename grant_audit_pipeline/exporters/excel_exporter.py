import openpyxl
import pandas as pd
import json
from pathlib import Path
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font

def sanitize_complex_types_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts list, dict, and complex object columns (e.g., DATE_SNIPPETS) 
    into clean JSON strings to prevent openpyxl export failures.
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

def export_audit_workbook(
    recon_df: pd.DataFrame, 
    header_df: pd.DataFrame, 
    tx_df: pd.DataFrame, 
    excel_path: Path
):
    """
    Exports audit results into 3 relational sheets:
    1. 3Way_Reconciliation: Compliance & Audit Verdicts
    2. Award_Headers: Macro Award Metadata & Multi-Year Ceilings
    3. Transaction_Ledger: Discrete NOA Increments & Actions with Column Segregation & Deduplication Flags
    """
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 1. Format datetime columns cleanly
    for col in ["PDF_START_DATE_TRUTH", "PDF_END_DATE_TRUTH"]:
        if col in recon_df.columns:
            recon_df[col] = pd.to_datetime(recon_df[col], format='mixed', errors='coerce')
        if col in header_df.columns:
            header_df[col] = pd.to_datetime(header_df[col], format='mixed', errors='coerce')

    date_cols_tx = ["ACTION_DATE", "BUDGET_PERIOD_START", "BUDGET_PERIOD_END"]
    for col in date_cols_tx:
        if col in tx_df.columns:
            tx_df[col] = pd.to_datetime(tx_df[col], format='mixed', errors='coerce')

    # 2. Sanitize complex JSON objects
    recon_df_clean  = sanitize_complex_types_for_excel(recon_df)
    header_df_clean = sanitize_complex_types_for_excel(header_df)
    tx_df_clean     = sanitize_complex_types_for_excel(tx_df)

    # 3. Write dataframes to Excel via openpyxl engine
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        recon_df_clean.to_excel(writer, sheet_name="3Way_Reconciliation", index=False)
        header_df_clean.to_excel(writer, sheet_name="Award_Headers", index=False)
        tx_df_clean.to_excel(writer, sheet_name="Transaction_Ledger", index=False)

    wb = openpyxl.load_workbook(excel_path)
    
    # 4. Apply Excel Tables & Short Date Formats across all 3 sheets
    sheet_tables = [
        ("3Way_Reconciliation", "Table_3Way_Reconciliation"),
        ("Award_Headers", "Table_Award_Headers"),
        ("Transaction_Ledger", "Table_Transaction_Ledger")
    ]

    for sheet_name, table_name in sheet_tables:
        ws = wb[sheet_name]
        tab = Table(displayName=table_name, ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tab)

        headers = [cell.value for cell in ws[1]]
        path_cols = ["PDF_PATH", "PDF_Path", "Markdown_Path"]

        for row in range(2, ws.max_row + 1):
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=row, column=col_idx)
                val_str = str(cell.value) if cell.value else ""

                if ("DATE" in str(h) or "START" in str(h) or "END" in str(h)) and cell.value:
                    cell.number_format = 'm/d/yyyy'

                if str(h) in path_cols and val_str and val_str not in ("N/A", "nan", "None", ""):
                    file_uri = "file:///" + val_str.replace("\\", "/")
                    cell.hyperlink = file_uri
                    cell.font = Font(color="0000FF", underline="single")

    wb.save(excel_path)