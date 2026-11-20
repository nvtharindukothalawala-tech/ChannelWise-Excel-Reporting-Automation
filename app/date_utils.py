"""Date and time parsing and conversion utilities."""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta

EXCEL_EPOCH = date(1899, 12, 30)
EXCEL_DATETIME_EPOCH = datetime(1899, 12, 30)

DATE_PATTERNS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
)

DATE_COLS = {
    "sales": ("date", "created on"),
    "returns": ("return date", "created on"),
    "stock": ("created date",),
    "summary": ("date",),
}

def to_date(value: any) -> date | None:
    """Parse date from datetime, date, Excel numeric serial, or string."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        num = float(value)
        if 20000 < num < 80000:
            return EXCEL_EPOCH + timedelta(days=int(num))
        return None

    text = str(value).strip()
    if not text:
        return None

    # Check numeric serial as string
    if re.fullmatch(r"\d{5}(?:\.\d+)?", text):
        num = float(text)
        if 20000 < num < 80000:
            return EXCEL_EPOCH + timedelta(days=int(num))

    for pat in DATE_PATTERNS:
        try:
            return datetime.strptime(text, pat).date()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass

    m = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None

    return None

def to_datetime(value: any) -> datetime | None:
    """Convert value to Python datetime."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        num = float(value)
        if 20000 < num < 80000:
            return EXCEL_DATETIME_EPOCH + timedelta(days=num)
        return None

    text = str(value).strip()
    if not text:
        return None

    if re.fullmatch(r"\d{5}(?:\.\d+)?", text):
        num = float(text)
        if 20000 < num < 80000:
            return EXCEL_DATETIME_EPOCH + timedelta(days=num)

    for pat in DATE_PATTERNS:
        try:
            return datetime.strptime(text, pat)
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(text)
    except ValueError:
        pass

    return None

def to_excel_serial(value: date | datetime | any) -> float | None:
    """Convert date or datetime to Excel numeric serial number."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        delta = value - EXCEL_DATETIME_EPOCH
        return delta.total_seconds() / 86400.0
    if isinstance(value, date):
        return float((value - EXCEL_EPOCH).days)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    dt = to_datetime(value)
    if dt is not None:
        return (dt - EXCEL_DATETIME_EPOCH).total_seconds() / 86400.0
    return None

def date_column_index(headers: list[str], kind: str) -> int | None:
    """Find index of the date column for the given dump kind."""
    lower_headers = [str(h).strip().lower() for h in headers]
    for candidate in DATE_COLS.get(kind, ()):
        if candidate in lower_headers:
            return lower_headers.index(candidate)
    return None

def find_reference_day(current_sales_dump: dict) -> int | None:
    """Find the latest day-of-month present in the current-month sales dump."""
    if not current_sales_dump or not current_sales_dump.get("rows"):
        return None
    col = date_column_index(current_sales_dump["headers"], "sales")
    if col is None:
        return None
    latest_day = 0
    for row in current_sales_dump["rows"]:
        if col < len(row):
            d = to_date(row[col])
            if d and d.day > latest_day:
                latest_day = d.day
    return latest_day or None

def extract_dump_month(dump: dict, kind: str) -> tuple[int, int]:
    """Validate that all records in the dump belong to exactly one month, return (year, month)."""
    col = date_column_index(dump["headers"], kind)
    if col is None:
        raise ValueError(f"Date column missing in {dump.get('file', 'dump')}")
    months = set()
    for row in dump.get("rows", []):
        if col < len(row):
            d = to_date(row[col])
            if d is None:
                raise ValueError(f"Invalid date in {dump.get('file', 'dump')}: {row[col]!r}")
            months.add((d.year, d.month))
    if len(months) != 1:
        raise ValueError(f"Expected exactly 1 month in {dump.get('file', 'dump')}; found {months}")
    return next(iter(months))
