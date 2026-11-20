"""
Report Execution Engine.
Orchestrates updates across all 13 Excel reporting workbooks.
"""
from __future__ import annotations

import logging
import os
import shutil
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from .brand_mapper import select_by_exact_locations, keyword_filter
from .data_cleaner import map_source_columns, to_decimal
from .date_utils import to_date, extract_dump_month
from .excel_preservation import repair_fashion_tracker_formulas
from .integrity_validator import validate_package_integrity, validate_dashboard_formulas
from .workbook_writer import (
    get_sheet_file_map,
    update_worksheet_xml,
    update_table_ref,
    verify_file_unlocked,
    calculate_sha256,
)

log = logging.getLogger("reports")

def truncate_rows(rows: list[list[any]], headers: list[str], kind: str, ref_day: int | None) -> list[list[any]]:
    """Keep only rows dated on or before ref_day of their month."""
    if not ref_day or not rows:
        return rows
    from .date_utils import date_column_index
    col = date_column_index(headers, kind)
    if col is None:
        return rows
    truncated = []
    for r in rows:
        if col < len(r):
            d = to_date(r[col])
            if d is None or d.day <= ref_day:
                truncated.append(r)
        else:
            truncated.append(r)
    return truncated

def process_report(
    brand: str,
    brand_cfg: dict,
    split_data: dict,
    dumps: dict,
    reports_dir: Path,
    dry_run: bool = False,
    backup_dir: Path | None = None,
    ref_day: int | None = None,
    match_fields: dict | None = None,
) -> dict:
    """
    Process a single reporting job using Architecture B.
    Returns status dict with counts, status, and messages.
    """
    file_rel_path = brand_cfg["file"]
    report_path = (reports_dir / file_rel_path).resolve()

    if not report_path.is_file():
        log.error("%-18s report file not found: %s", brand, report_path)
        return {"brand": brand, "status": "FAILED", "reason": f"File not found: {report_path}"}

    # Lock detection
    try:
        verify_file_unlocked(report_path)
    except PermissionError as exc:
        log.error("%-18s file is locked — close it in Excel: %s", brand, exc)
        return {"brand": brand, "status": "FAILED", "reason": str(exc)}

    t0 = time.time()
    counts = {}
    missing_dumps = []
    has_missing_stock = False

    # 1. Inspect existing workbook structure via zipfile
    with zipfile.ZipFile(report_path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}
        sheet_file_map = get_sheet_file_map(zin)

    # 2. Process each sheet configured for this report
    updated_sheet_parts = {}
    table_updates = {}
    data_sheet_names = []

    for dump_key, spec in brand_cfg["sheets"].items():
        sheet_name = spec["sheet"]
        data_sheet_names.append(sheet_name)

        if sheet_name not in sheet_file_map:
            log.error("%-18s sheet %r not in workbook", brand, sheet_name)
            continue

        sheet_file = sheet_file_map[sheet_name]
        sheet_part_name = f"xl/worksheets/{sheet_file}"

        # Check source dump availability
        sources = spec.get("stack") or [dump_key]
        absent_sources = [k for k in sources if k not in dumps]

        if absent_sources:
            if any("stock" in k for k in absent_sources):
                has_missing_stock = True
                log.warning("%-18s %-24s SKIPPED (stock dump missing: %s)", brand, sheet_name, ", ".join(absent_sources))
            else:
                log.warning("%-18s %-24s SKIPPED (dump not available: %s)", brand, sheet_name, ", ".join(absent_sources))
            missing_dumps.extend(absent_sources)
            continue

        kind = "returns" if "returns" in sources[0] else ("stock" if "stock" in sources[0] else "sales")
        sheet_rows = []

        for src_key in sources:
            pool = split_data.get(src_key, {})
            keywords = (
                spec.get("keywords")
                or spec.get("include_keywords")
                or brand_cfg.get("keywords")
                or brand_cfg.get("include_keywords")
            )
            exclude = spec.get("exclude_keywords") or brand_cfg.get("exclude_keywords")

            if "exact_locations" in spec:
                part_rows = select_by_exact_locations(
                    dumps[src_key], kind, spec["exact_locations"], match_fields
                )
            elif keywords:
                part_rows = keyword_filter(
                    pool.get("__ALL__", []),
                    dumps[src_key]["headers"],
                    kind,
                    match_fields,
                    keywords,
                    exclude,
                )
            elif brand_cfg.get("all_rows"):
                part_rows = pool.get("__ALL__", [])
            elif brand_cfg.get("brands"):
                part_rows = [r for b in brand_cfg["brands"] for r in pool.get(b, [])]
            else:
                part_rows = pool.get(brand, [])

            # Reference day truncation
            if (
                spec.get("truncate_to_day")
                and kind != "stock"
                and src_key not in ("sales_current_month", "returns_current_month")
            ):
                part_rows = truncate_rows(part_rows, dumps[src_key]["headers"], kind, ref_day)

            sheet_rows.extend(part_rows)

        # Schema mapping if configured
        target_month = None
        if spec.get("source_columns"):
            sheet_rows = map_source_columns(
                dumps[sources[0]]["headers"], sheet_rows, spec["source_columns"]
            )
            if spec.get("replace_month"):
                target_month = extract_dump_month(dumps[sources[0]], kind)

        # Update worksheet XML
        orig_sheet_xml = parts[sheet_part_name]
        new_sheet_xml, row_count = update_worksheet_xml(
            sheet_xml_bytes=orig_sheet_xml,
            data_start_row=spec["data_start_row"],
            paste_start_col=spec["paste_start_col"],
            new_rows=sheet_rows,
            skip_cols=spec.get("skip_cols"),
            max_cols=spec.get("max_cols"),
            date_target_header=spec.get("date_target_header"),
            source_columns=spec.get("source_columns"),
            replace_month=target_month,
        )

        updated_sheet_parts[sheet_part_name] = new_sheet_xml
        counts[sheet_name] = row_count

        # Check for Excel tables associated with this sheet
        rel_file_name = f"xl/worksheets/_rels/{sheet_file}.rels"
        if rel_file_name in parts:
            srels_root = ET.fromstring(parts[rel_file_name])
            for r in srels_root:
                if r.get("Type", "").endswith("/table"):
                    target_table_path = "xl/" + r.get("Target").lstrip("../").lstrip("/")
                    if target_table_path in parts:
                        new_last_row = spec["data_start_row"] + row_count - 1
                        updated_table_xml = update_table_ref(parts[target_table_path], new_last_row)
                        table_updates[target_table_path] = updated_table_xml

        tag = ""
        if brand_cfg.get("include_keywords"):
            tag = f" ({'+'.join(brand_cfg['include_keywords'])})"
        elif brand_cfg.get("brands"):
            tag = f" ({'+'.join(brand_cfg['brands'])})"
        elif brand_cfg.get("all_rows"):
            tag = " (all brands)"

        log.info("%-18s %-24s %6d rows%s", brand, sheet_name, row_count, tag)

    if dry_run:
        status_msg = "DRY RUN — PASSED (NOT SAVED)"
        if has_missing_stock:
            status_msg = "DRY RUN — INCOMPLETE (STOCK MISSING, SALES/RETURNS VALID)"
        log.info("%-18s %s", brand, status_msg)
        return {
            "brand": brand,
            "status": "PASS" if not has_missing_stock else "INCOMPLETE",
            "counts": counts,
            "has_missing_stock": has_missing_stock,
            "dry_run": True,
        }

    # Create backup with SHA-256 verification
    backup_path = None
    if backup_dir:
        backup_dir = Path(backup_dir)
        backup_path = backup_dir / file_rel_path
        backup_path.parent.mkdir(parents=True, exist_ok=True)
        if backup_path.exists():
            log.error("%s: backup already exists: %s", brand, backup_path)
            return {"brand": brand, "status": "FAILED", "reason": "Backup already exists"}

        shutil.copy2(report_path, backup_path)
        if calculate_sha256(report_path) != calculate_sha256(backup_path):
            raise RuntimeError(f"Backup hash mismatch for {brand}; refusing to modify report")
        log.info("%-18s verified backup created: %s", brand, backup_path)

    # Prepare staged copy
    tmp_path = report_path.with_suffix(report_path.suffix + ".tmp")
    if tmp_path.exists():
        tmp_path.unlink()

    # Apply updates to zip parts
    staged_parts = dict(parts)
    staged_parts.update(updated_sheet_parts)
    staged_parts.update(table_updates)

    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
            for n, data in staged_parts.items():
                zout.writestr(n, data)

        # If Fashion Tracker, apply verified formula repairs
        if brand == "FashionTracker":
            repair_fashion_tracker_formulas(tmp_path)

        # Integrity Validation on staged output
        integrity_errors = validate_package_integrity(tmp_path, backup_path)
        if integrity_errors:
            raise RuntimeError(f"Staged report integrity check failed: {integrity_errors}")

        formula_errors = validate_dashboard_formulas(tmp_path, backup_path, tuple(data_sheet_names))
        if formula_errors:
            raise RuntimeError(f"Staged dashboard formula check failed: {formula_errors}")

        # Atomic replacement
        os.replace(tmp_path, report_path)

        status_label = "PASS"
        if has_missing_stock:
            status_label = "INCOMPLETE"
            log.warning("%-18s saved with status INCOMPLETE (stock missing, sales/returns valid) in %.1fs", brand, time.time() - t0)
        else:
            log.info("%-18s saved successfully in %.1fs", brand, time.time() - t0)

        return {
            "brand": brand,
            "status": status_label,
            "counts": counts,
            "has_missing_stock": has_missing_stock,
            "elapsed_seconds": time.time() - t0,
        }

    except Exception as exc:
        log.exception("%-18s update failed: %s", brand, exc)
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass
        return {"brand": brand, "status": "FAILED", "reason": str(exc)}
