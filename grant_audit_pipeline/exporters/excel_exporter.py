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
    df_clean = df.copy()
    for col in df_clean.columns:
        # Check if any element in the column is a list, dict, or complex object
        if df_clean[col].apply(lambda x: isinstance(x, (list, dict))).any():
            df_clean[col] = df_clean[col].apply(
                lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else ("" if pd.isna(x) else str(x))
            )
    return df_clean

def export_audit_workbook(recon_df: pd.DataFrame, doc_df: pd.DataFrame, excel_path: Path):
    """Exports audit results with Excel Tables, Clickable Hyperlinks, Short Dates, and JSON String Sanitization."""
    excel_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 1. Format and parse datetime columns
    for col in ["PDF_START_DATE_TRUTH", "PDF_END_DATE_TRUTH"]:
        if col in recon_df.columns:
            recon_df[col] = pd.to_datetime(recon_df[col], errors='coerce')
            
    date_cols_doc = [
        "EXECUTION_DATE", "DOC_START_DATE", "DOC_END_DATE", 
        "VISION_EXECUTION_DATE", "VISION_BUDGET_START", "VISION_BUDGET_END", 
        "VISION_PROJECT_START", "VISION_PROJECT_END"
    ]
    for col in date_cols_doc:
        if col in doc_df.columns:
            doc_df[col] = pd.to_datetime(doc_df[col], errors='coerce')

    # 2. Drop heavy unneeded body/raw columns for Excel export
    cols_to_drop = ["FILE_HASH", "SOURCE_TAG", "RAW_CAYUSE_PROJ", "RAW_CAYUSE_PROP", "RAW_ORACLE_NUM", "RAW_BANNER_UID", "TEXT_BODY"]
    doc_excel_df = doc_df.drop(columns=[c for c in cols_to_drop if c in doc_df.columns])

    # 3. Convert complex objects (e.g., DATE_SNIPPETS) to clean JSON strings
    recon_df_clean = sanitize_complex_types_for_excel(recon_df)
    doc_excel_df_clean = sanitize_complex_types_for_excel(doc_excel_df)

    # 4. Write dataframes to Excel via openpyxl engine
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        recon_df_clean.to_excel(writer, sheet_name="3Way_Reconciliation", index=False)
        doc_excel_df_clean.to_excel(writer, sheet_name="Document_Level_Detail", index=False)

    wb = openpyxl.load_workbook(excel_path)
    
    # 5. Format 3Way_Reconciliation Sheet
    ws_recon = wb["3Way_Reconciliation"]
    tab_recon = Table(displayName="Table_3Way_Reconciliation", ref=ws_recon.dimensions)
    tab_recon.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_recon.add_table(tab_recon)
    
    header_recon = [cell.value for cell in ws_recon[1]]
    for row in range(2, ws_recon.max_row + 1):
        for col_idx, h in enumerate(header_recon, 1):
            cell = ws_recon.cell(row=row, column=col_idx)
            if "DATE" in str(h) and cell.value:
                cell.number_format = 'm/d/yyyy'

    # 6. Format Document_Level_Detail Sheet
    ws_docs = wb["Document_Level_Detail"]
    tab_docs = Table(displayName="Table_Document_Level_Detail", ref=ws_docs.dimensions)
    tab_docs.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_docs.add_table(tab_docs)

    header_docs = [cell.value for cell in ws_docs[1]]
    path_cols = ["PDF_Path", "Markdown_Path", "CROSS_REF_PATH"]

    for row in range(2, ws_docs.max_row + 1):
        for col_idx, h in enumerate(header_docs, 1):
            cell = ws_docs.cell(row=row, column=col_idx)
            val_str = str(cell.value) if cell.value else ""
            
            # Short Date Format for any date-related column
            if ("DATE" in str(h) or "START" in str(h) or "END" in str(h)) and cell.value:
                cell.number_format = 'm/d/yyyy'

            # Format Clickable Hyperlinks
            if str(h) in path_cols and val_str and val_str not in ("N/A", "nan", "None", ""):
                file_uri = "file:///" + val_str.replace("\\", "/")
                cell.hyperlink = file_uri
                cell.font = Font(color="0000FF", underline="single")

    wb.save(excel_path)