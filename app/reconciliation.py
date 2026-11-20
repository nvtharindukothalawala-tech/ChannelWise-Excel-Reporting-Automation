"""Reconciliation and validation utilities."""
from __future__ import annotations

import logging
from decimal import Decimal
from .data_cleaner import to_decimal

log = logging.getLogger("reports")

def reconcile_sales_dump(dump: dict, brand_allocations: dict[str, list[list[any]]]) -> dict:
    """Reconcile sales dump rows and financial amounts across brand allocations."""
    headers = dump["headers"]
    rows = dump["rows"]
    amount_col = headers.index("Sales Amount") if "Sales Amount" in headers else None

    source_rows = len(rows)
    source_amount = sum(
        (to_decimal(r[amount_col], "Sales Amount") for r in rows if amount_col is not None and amount_col < len(r)),
        Decimal("0.00")
    )

    allocated_rows = sum(len(b_rows) for b, b_rows in brand_allocations.items() if b != "__ALL__")
    allocated_amount = sum(
        (to_decimal(r[amount_col], "Sales Amount") for b, b_rows in brand_allocations.items() if b != "__ALL__"
         for r in b_rows if amount_col is not None and amount_col < len(r)),
        Decimal("0.00")
    )

    return {
        "source_rows": source_rows,
        "allocated_rows": allocated_rows,
        "source_amount": source_amount,
        "allocated_amount": allocated_amount,
        "row_diff": source_rows - allocated_rows,
        "amount_diff": source_amount - allocated_amount,
    }

def reconcile_returns_dump(dump: dict, brand_allocations: dict[str, list[list[any]]]) -> dict:
    """Reconcile returns dump rows and financial amounts across brand allocations."""
    headers = dump["headers"]
    rows = dump["rows"]
    amount_col = headers.index("Sale Amount") if "Sale Amount" in headers else None

    source_rows = len(rows)
    source_amount = sum(
        (to_decimal(r[amount_col], "Sale Amount") for r in rows if amount_col is not None and amount_col < len(r)),
        Decimal("0.00")
    )

    allocated_rows = sum(len(b_rows) for b, b_rows in brand_allocations.items() if b != "__ALL__")
    allocated_amount = sum(
        (to_decimal(r[amount_col], "Sale Amount") for b, b_rows in brand_allocations.items() if b != "__ALL__"
         for r in b_rows if amount_col is not None and amount_col < len(r)),
        Decimal("0.00")
    )

    return {
        "source_rows": source_rows,
        "allocated_rows": allocated_rows,
        "source_amount": source_amount,
        "allocated_amount": allocated_amount,
        "row_diff": source_rows - allocated_rows,
        "amount_diff": source_amount - allocated_amount,
    }
