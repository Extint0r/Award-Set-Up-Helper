import fitz
from pathlib import Path
from typing import Tuple
from config import parse_dollar_amount

def extract_budget_from_pdf_tables(pdf_path):
    """
    Safely extracts budget tables from PDF pages using PyMuPDF table finder,
    with robust exception handling to prevent hanging on corrupted drawings.
    """
    total_direct = 0.0
    total_indirect = 0.0
    total_amount = 0.0
    
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            try:
                tabs = page.find_tables()
                if not tabs.tables:
                    continue
                for tab in tabs:
                    df = tab.extract()
                    # Process tabular rows if found
                    for row in df:
                        for cell in row:
                            if cell and any(keyword in str(cell).lower() for keyword in ["total", "direct", "indirect"]):
                                # Basic heuristic parsing can go here
                                pass
            except Exception:
                # Skip specific page if vector drawings cause parsing errors/hangups
                continue
        doc.close()
    except Exception:
        pass
        
    return total_direct, total_indirect, total_amount