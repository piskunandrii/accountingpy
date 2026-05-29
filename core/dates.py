import re
from datetime import date, datetime, timedelta
from typing import Iterable

def to_date(value) -> date:
    if pd.isna(value):
        raise ValueError("Missing Date of issue")
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return pd.to_datetime(value).date()

def previous_calendar_day(d: date) -> date:
    return d - timedelta(days=1)

def date_chunks(start: date, end: date, max_days: int = 360) -> Iterable[tuple[date, date]]:
    """NBP API has range limits, so fetch in chunks."""
    current = start
    while current <= end:
        chunk_end = min(current + timedelta(days=max_days), end)
        yield current, chunk_end
        current = chunk_end + timedelta(days=1)

def normalize_date_text(value: str):
    if not value:
        return ""
    match = re.search(r"(\d{2})[./](\d{2})[./](\d{4})", str(value))
    if not match:
        return ""
    day, month, year = match.groups()
    return f"{day}.{month}.{year}"

def get_first_date(value: str):
    matches = re.findall(r"\d{2}[./]\d{2}[./]\d{4}", str(value))
    return normalize_date_text(matches[0]) if matches else ""

def get_last_date(value: str):
    matches = re.findall(r"\d{2}[./]\d{2}[./]\d{4}", str(value))
    return normalize_date_text(matches[-1]) if matches else ""

def parse_compact_register_date(value: str, *, is_start: bool) -> date:
    """
    Supports:
      MMYYYY   -> first/last day of the month depending on is_start
      DDMMYYYY -> exact day
    """
    value = value.strip()
    if len(value) == 6:
        month = int(value[:2])
        year = int(value[2:])
        if is_start:
            return date(year, month, 1)
        if month == 12:
            return date(year, 12, 31)
        return date(year, month + 1, 1) - timedelta(days=1)
    if len(value) == 8:
        day = int(value[:2])
        month = int(value[2:4])
        year = int(value[4:])
        return date(year, month, day)
    raise ValueError(f"Unsupported compact date: {value}")
