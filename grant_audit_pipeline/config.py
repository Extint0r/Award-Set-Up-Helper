import os
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Any, Optional
from dotenv import load_dotenv

# Load secret environment variables from .env file
load_dotenv()

# ==========================================
# PATH & ENVIRONMENT CONFIGURATION
# ==========================================
BASE_DIR = Path(__file__).parent

# Absolute anchor guarantees it finds the file in the script's root directory
TRIAGE_EXCEL_PATH  = BASE_DIR / "2026_7_20_CAYUSE_ORACLE_TRIAGE.xlsx" # (or 2027_7_20 if named 2027)
ALN_CSV_PATH       = BASE_DIR / "ALN.csv"

OSP_SOURCE_DIR     = Path(r"D:\0-Batch-AWARDS\processed_files")
ORACLE_PARENT_DIR  = Path(r"D:\OSR pdf notices")

MD_OUTPUT_DIR      = Path(r"D:\0-Batch-AWARDS\processed_files\Markdown")
REVIEW_DIR         = Path(r"D:\0-Batch-AWARDS\processed_files\Review")

OUTPUT_EXCEL_PATH  = REVIEW_DIR / "AUDIT_INDEX_TABLE.xlsx"
OUTPUT_SQLITE_PATH = REVIEW_DIR / "AUDIT_INDEX_TABLE.db"

VISION_CACHE_DIR   = REVIEW_DIR / "Vision_Cache"

# ==========================================
# SANITY THRESHOLDS & CONSTANTS
# ==========================================
MAX_REASONABLE_TRANSACTION = 200_000_000.0  # $200M Cap per action
ENABLE_VISION_FALLBACK     = True           # Toggle Vision AI API fallback
TEST_RUN_LIMIT             = 200            # Set to integer (e.g. 100) or None for full run
VISION_API_KEY             = os.getenv("GEMINI_API_KEY") or os.getenv("VISION_API_KEY")

# ==========================================
# REGEX PATTERNS
# ==========================================
RE_ORACLE_NUM  = re.compile(r'(?<!\d)([1-9]\d{5})(?!\d)')
RE_CAYUSE_PROJ = re.compile(r'(?<![A-Za-z0-9])(\d{2}-\d{4})(?![A-Za-z0-9])')
RE_CAYUSE_PROP = re.compile(r'(?<![A-Za-z0-9])(A\d{2}-\d{4})(?![A-Za-z0-9])', re.IGNORECASE)
RE_BANNER_UID = re.compile(r'(?<![A-Za-z0-9])([Rr](?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9]{5})(?![A-Za-z0-9])')
RE_ALN         = re.compile(r'(?<!\d)(\d{2}\.\d{3})(?!\d)')

RE_EXECUTION_DATE = re.compile(
    r'(?:Date|Execution\s*Date|Notice\s*Date|Award\s*Date)\s*[:\=]?\s*(\d{1,2}/\d{1,2}/\d{2,4}|[A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}-[A-Za-z]+-\d{2,4})',
    re.IGNORECASE
)

RE_DATE_RANGE = re.compile(
    r'(?:Project\s*Period|Award\s*Period|Performance\s*Period|Grant\s*Period)\s*[:\=]?\s*'
    r'(\d{1,2}/\d{1,2}/\d{2,4}|[A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}-[A-Za-z]+-\d{2,4})\s*[-–—\s+to\s+]+\s*'
    r'(\d{1,2}/\d{1,2}/\d{2,4}|[A-Za-z]+\s+\d{1,2},\s*\d{4}|\d{1,2}-[A-Za-z]+-\d{2,4})',
    re.IGNORECASE
)

RE_CEILING_AMOUNT = re.compile(
    r'(?:Project\s*Total\s*Amount|Total\s*Award\s*Amount|Total\s*Ceiling|Award\s*Ceiling|Total\s*Project\s*Ceiling)\s*(?:\([^)]*\))?\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)\s*(k|m|million|thousand)?',
    re.IGNORECASE
)

RE_OBLIGATED_ACTION = re.compile(
    r'(?:Current\s*Action.*?totaling|Obligated\s*Amount|Amount\s*Awarded\s*This\s*Action|Funding\s*This\s*Action|Action\s*Amount|This\s*Action)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)\s*(k|m|million|thousand)?',
    re.IGNORECASE
)

RE_BUDGET_TOTAL = re.compile(
    r'(?:Total\s*Budget|Total\s*Costs|Total\s*Amount|Total\s*Direct\s*and\s*Indirect)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)\s*(k|m|million|thousand)?',
    re.IGNORECASE
)

RE_INDIRECT_COST = re.compile(
    r'(?:Indirect\s*Costs?|F&A\s*Costs?|Facilities\s*&\s*Admin|Overhead)\s*[:\=]?\s*\$?\s*([\d,]+(?:\.\d{2})?)\s*(k|m|million|thousand)?',
    re.IGNORECASE
)

BUDGET_PAGE_KEYWORDS = [
    "direct cost", "f&a", "facilities and administrative", "f&a rate", 
    "f&a cost", "indirect cost", "total direct", "modified total direct"
]

# ==========================================
# HELPER UTILITIES
# ==========================================
def find_markdown_file(pdf_path: Path, md_dir: Path) -> Optional[Path]:
    """Finds existing Markdown file regardless of prefix variations."""
    exact_path = md_dir / f"{pdf_path.stem}.md"
    if exact_path.exists():
        return exact_path
    matches = list(md_dir.glob(f"*{pdf_path.stem}.md"))
    if matches:
        return matches[0]
    return None

def format_master_yaml_header(metadata: dict) -> str:
    """Formats unified metadata into YAML frontmatter block."""
    lines = ["---"]
    for k, v in metadata.items():
        if isinstance(v, list):
            lines.append(f"{k}: {v}")
        elif isinstance(v, bool):
            lines.append(f"{k}: {str(v).lower()}")
        elif isinstance(v, (int, float)):
            lines.append(f"{k}: {v}")
        elif v is None:
            lines.append(f"{k}: ''")
        else:
            clean_val = str(v).replace("'", "''")
            lines.append(f"{k}: '{clean_val}'")
    lines.append("---\n\n")
    return "\n".join(lines)

def parse_dollar_amount(val_str: Any) -> float:
    """
    Sanitized currency parser that prevents dates (05/01/2015), phone numbers, 
    or file paths from corrupting into astronomical values.
    """
    if val_str is None or str(val_str).strip() in ("", "nan", "None"):
        return 0.0
    
    s = str(val_str).strip()
    
    # Reject paths, filenames, or scientific notation strings
    if re.search(r'\.(pdf|md|txt|xlsx|csv)\b', s, re.IGNORECASE) or '_' in s or 'e+' in s.lower():
        return 0.0

    # Reject if string contains explicit dates
    if re.search(r'\b\d{1,2}/\d{1,2}/\d{2,4}\b', s) or \
       re.search(r'\b\d{4}-\d{2}-\d{2}\b', s) or \
       re.search(r'\b\d{1,2}-[A-Za-z]{3}-\d{2,4}\b', s, re.IGNORECASE):
        return 0.0

    # Reject IDs / CFDA numbers
    if re.search(r'\bR\d{4,6}\b', s, re.IGNORECASE) or \
       re.search(r'\b\d{2}\.\d{3}\b', s) or \
       re.search(r'\b\d{2}-\d{4}\b', s):
        return 0.0

    s_lower = s.lower()

    # 1. Shorthand notation ($500k, $1.8M)
    m_short = re.search(r'\$?\s*([\d,]+(?:\.\d+)?)\s*(k|m|million|thousand|b|billion)\b', s_lower)
    if m_short:
        num = float(m_short.group(1).replace(',', ''))
        unit = m_short.group(2)
        mult = 1_000.0 if unit in ('k', 'thousand') else \
               1_000_000.0 if unit in ('m', 'million') else \
               1_000_000_000.0 if unit in ('b', 'billion') else 1.0
        val = num * mult
        return val if val <= MAX_REASONABLE_TRANSACTION else 0.0

    # 2. Dollar format with $ prefix or clean number
    if '$' in s:
        m_curr = re.search(r'\$\s*([\d,]+(?:\.\d{1,2})?)', s)
        if m_curr:
            try:
                val = float(m_curr.group(1).replace(',', ''))
                return val if val <= MAX_REASONABLE_TRANSACTION else 0.0
            except ValueError:
                return 0.0
    else:
        clean_s = s.replace(',', '').strip()
        if re.match(r'^\d+(?:\.\d{1,2})?$', clean_s):
            try:
                val = float(clean_s)
                return val if val <= MAX_REASONABLE_TRANSACTION else 0.0
            except ValueError:
                return 0.0

    return 0.0

def parse_date(date_str: Any) -> Optional[datetime]:
    if pd.isna(date_str) or not date_str:
        return None
    s = str(date_str).strip()
    formats = (
        "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d",
        "%B %d, %Y", "%b %d, %Y", "%d-%b-%Y", "%d %b %Y", "%d %B %Y"
    )
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None