"""Data loader for Green Cloud POS exports."""
from __future__ import annotations

import logging
from pathlib import Path
import openpyxl
from .data_cleaner import clean_sales_dump, clean_returns_dump
from .date_utils import find_reference_day

log = logging.getLogger("reports")

def read_dump_file(path: Path) -> tuple[list[str], list[list[any]]]:
    """Read an Excel dump file. Row 1 is header."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        try:
            raw_headers = next(rows_iter)
            headers = [("" if h is None else str(h).strip()) for h in raw_headers]
        except StopIteration:
            return [], []

        rows = [list(r) for r in rows_iter if any(v is not None and str(v).strip() for v in r)]
        return headers, rows
    finally:
        wb.close()

def required_dump_keys(selected_targets: dict) -> set[str]:
    """Calculate the set of dump keys needed by the selected reporting jobs."""
    needed = set()
    for bc in selected_targets.values():
        for key, spec in bc.get("sheets", {}).items():
            needed.update(spec.get("stack") or [key])
    return needed

def load_dumps(dumps_dir: Path, dump_files_cfg: dict, needed_keys: set[str] | None = None) -> tuple[dict[str, dict], list[str]]:
    """
    Load required dumps from dumps_dir.
    Returns (loaded_dumps, missing_dump_files).
    """
    dumps = {}
    missing_files = []

    for key, fname in dump_files_cfg.items():
        if needed_keys is not None and key not in needed_keys:
            continue

        file_path = dumps_dir / fname
        if not file_path.is_file():
            missing_files.append(fname)
            log.warning("dump file missing, will skip sheets that need it: %s", fname)
            continue

        raw_headers, raw_rows = read_dump_file(file_path)

        # Schema cleaning
        if key.startswith("sales"):
            headers, rows = clean_sales_dump(raw_headers, raw_rows, fname)
        elif key.startswith("returns"):
            headers, rows = clean_returns_dump(raw_headers, raw_rows, fname)
        else:
            headers = raw_headers
            rows = raw_rows

        dumps[key] = {
            "headers": headers,
            "rows": rows,
            "file": fname,
            "key": key,
        }
        log.info("loaded %-25s %6d rows, %2d cols", fname, len(rows), len(headers))

    return dumps, missing_files
