"""Read-only integrity validator for report workbooks."""
from __future__ import annotations

import logging
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal
from pathlib import Path
from .data_cleaner import to_decimal
from .date_utils import to_date
from .excel_preservation import inspect_workbook_components

log = logging.getLogger("reports")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": MAIN_NS}
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def validate_package_integrity(report_path: Path, backup_path: Path | None = None) -> list[str]:
    """Validate OpenXML package integrity and verify component stability."""
    errors = []
    report_path = Path(report_path)
    if not report_path.is_file():
        return [f"Report file missing: {report_path}"]

    # 1. Test zip CRC
    try:
        with zipfile.ZipFile(report_path, "r") as z:
            bad_crc = z.testzip()
            if bad_crc is not None:
                errors.append(f"ZIP CRC corrupted: {bad_crc}")
    except Exception as exc:
        return [f"ZIP package invalid: {exc}"]

    # 2. Check component preservation if backup provided
    if backup_path and Path(backup_path).is_file():
        before = inspect_workbook_components(Path(backup_path))
        after = inspect_workbook_components(report_path)

        for comp in ("pivots", "caches", "cache_records", "charts", "drawings", "external_links"):
            if after[comp] != before[comp]:
                errors.append(f"{comp} count changed: after={after[comp]}, before={before[comp]}")

        # Check drawing ownership (no shared drawings between multiple sheets)
        for sname, targets in after["sheet_drawings"].items():
            if len(targets) > 1:
                errors.append(f"Worksheet {sname} has multiple drawings: {targets}")

        all_drawings = [t for targets in after["sheet_drawings"].values() for t in targets]
        if len(all_drawings) != len(set(all_drawings)):
            errors.append(f"Drawing part shared across multiple worksheets: {all_drawings}")

    return errors

def validate_dashboard_formulas(report_path: Path, backup_path: Path, data_sheets: tuple[str, ...]) -> list[str]:
    """Verify that formulas on non-data sheets have not changed."""
    errors = []

    def _read_sheet_formulas(zip_handle: zipfile.ZipFile, sfile: str) -> dict[str, str]:
        root = ET.fromstring(zip_handle.read(f"xl/worksheets/{sfile}"))
        formulas = {}
        for c in root.findall(".//main:c", NS):
            f = c.find("main:f", NS)
            if f is not None and f.text:
                formulas[c.get("r")] = f.text
        return formulas

    with zipfile.ZipFile(report_path, "r") as z_new, zipfile.ZipFile(backup_path, "r") as z_old:
        wb_new = ET.fromstring(z_new.read("xl/workbook.xml"))
        wb_old = ET.fromstring(z_old.read("xl/workbook.xml"))
        rels_new = {r.get("Id"): r.get("Target") for r in ET.fromstring(z_new.read("xl/_rels/workbook.xml.rels"))}
        rels_old = {r.get("Id"): r.get("Target") for r in ET.fromstring(z_old.read("xl/_rels/workbook.xml.rels"))}

        sheets_new = {s.get("name"): rels_new[s.get(f"{{{DOC_REL_NS}}}id")].split("/")[-1] for s in wb_new.findall(".//main:sheet", NS)}
        sheets_old = {s.get("name"): rels_old[s.get(f"{{{DOC_REL_NS}}}id")].split("/")[-1] for s in wb_old.findall(".//main:sheet", NS)}

        for sname in sheets_old:
            if sname in data_sheets or sname not in sheets_new:
                continue
            old_formulas = _read_sheet_formulas(z_old, sheets_old[sname])
            new_formulas = _read_sheet_formulas(z_new, sheets_new[sname])

            for coord in set(old_formulas) | set(new_formulas):
                if old_formulas.get(coord) != new_formulas.get(coord):
                    # Check if this is an approved repair on Fashion Tracker
                    if (sname, coord) in (("Revenu summary", "D74"), ("Total Rev", "AH4")):
                        continue
                    errors.append(f"Formula regression in {sname}!{coord}: old={old_formulas.get(coord)}, new={new_formulas.get(coord)}")

    return errors
