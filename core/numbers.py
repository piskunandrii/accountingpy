import re
import math
from typing import Optional
import pandas as pd
from config import NORMALIZE_FV_TO_FE

def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)

def clean_doc_number(value) -> str:
    """Normalize document number text for stable joins/export."""
    text = clean_text(value)
    if NORMALIZE_FV_TO_FE:
        text = re.sub(r"^FV\b", "FE", text, flags=re.IGNORECASE)
    return text

def normalize_currency(value, default: Optional[str] = None) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return default or ""
    return str(value).strip().upper()

def to_num(value) -> float:
    """Convert Excel/Polish/European numeric values to float."""
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)

    text = str(value).strip().replace("\u00a0", "")
    text = text.replace(" ", "")

    # Polish/European format: 1.234,56 -> 1234.56
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    return float(text)

def pl_number_to_float(value: str):
    if value is None:
        return None
    value = str(value).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return None

def float_to_pl_text(value):
    if value is None or pd.isna(value):
        return ""
    text = f"{value:,.2f}"
    return text.replace(",", "X").replace(".", ",").replace("X", " ")

def extract_amounts(line: str):
    return re.findall(r"-?\d[\d\s\u00a0]*[,.]\d{2}", str(line))
