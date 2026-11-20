"""Data cleaning, validation, and schema enforcement."""
from __future__ import annotations

import logging
from decimal import Decimal, InvalidOperation
from .date_utils import to_datetime

log = logging.getLogger("reports")

EXCLUDED_SALES_FIELDS = {"vat amount", "total (incl. vat)"}

STANDARD_SALES_SCHEMA_26 = (
    "Bill No",
    "Location",
    "Customer",
    "Address",
    "Created By",
    "City",
    "Sales Channel",
    "Note",
    "Exchange Ref",
    "Contact No1",
    "Contact No2",
    "Sales Person",
    "Invoice Type",
    "Reference",
    "Delivery Amount",
    "Product",
    "Cost Price",
    "Sales Price",
    "Discount",
    "Exchange",
    "Sold Qty",
    "Cost Amount",
    "Sales Amount",
    "Paid By",
    "Date",
    "Serials",
)

def clean_sales_dump(headers: list[str], rows: list[list[any]], filename: str) -> tuple[list[str], list[list[any]]]:
    """
    Remove accidental export-only VAT fields by exact name.
    Verify that exactly 26 standard sales columns remain.
    """
    keep_indices = []
    ignored = []
    for i, h in enumerate(headers):
        norm = str(h or "").strip().casefold()
        if norm in EXCLUDED_SALES_FIELDS:
            ignored.append(str(h).strip())
        else:
            keep_indices.append(i)

    if ignored:
        log.info("%s: ignored accidental export-only VAT fields: %s", filename, ", ".join(ignored))

    cleaned_headers = [str(headers[i]).strip() if headers[i] is not None else "" for i in keep_indices]
    cleaned_rows = [
        [row[i] if i < len(row) else None for i in keep_indices]
        for row in rows
    ]

    if len(cleaned_headers) != 26:
        raise ValueError(
            f"{filename}: expected exactly 26 standard sales columns after filtering VAT fields; "
            f"found {len(cleaned_headers)}. Check dump headers before proceeding."
        )

    return cleaned_headers, cleaned_rows

def clean_returns_dump(headers: list[str], rows: list[list[any]], filename: str) -> tuple[list[str], list[list[any]]]:
    """Clean returns headers and rows."""
    cleaned_headers = [str(h or "").strip() for h in headers]
    cleaned_rows = [list(r) for r in rows]
    return cleaned_headers, cleaned_rows

def to_decimal(value: any, field_label: str = "amount") -> Decimal:
    """Parse numeric or decimal value safely."""
    if value is None or value == "":
        return Decimal("0.00")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        raise ValueError(f"Invalid boolean value for financial field {field_label}")
    try:
        cleaned_str = str(value).replace(",", "").strip()
        result = Decimal(cleaned_str)
        if not result.is_finite():
            raise ValueError(f"Non-finite decimal in {field_label}: {value}")
        return result
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Invalid financial amount in {field_label}: {value!r}") from exc

def map_source_columns(headers: list[str], rows: list[list[any]], target_columns: list[str | None]) -> list[list[any]]:
    """
    Match dump columns by HEADER NAME (case-insensitive).
    None in target_columns results in a blank value.
    Converts Date, Return Date, Created On to Excel date-times.
    """
    header_map = {}
    for i, h in enumerate(headers):
        key = str(h or "").strip().casefold()
        if key in header_map:
            raise ValueError(f"Duplicate source column in dump headers: {h!r}")
        header_map[key] = i

    missing = [col for col in target_columns if col is not None and col.strip().casefold() not in header_map]
    if missing:
        raise ValueError(f"Missing required source columns in dump: {missing}")

    positions = [None if col is None else header_map[col.strip().casefold()] for col in target_columns]

    mapped_rows = []
    for row in rows:
        mapped_row = []
        for name, pos in zip(target_columns, positions):
            val = None if pos is None else (row[pos] if pos < len(row) else None)
            if name in ("Date", "Return Date", "Created On") and val is not None:
                val = to_datetime(val)
            mapped_row.append(val)
        mapped_rows.append(mapped_row)

    return mapped_rows
