"""Excel component preservation, formula repair, and integrity checks."""
from __future__ import annotations

import logging
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

log = logging.getLogger("reports")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": MAIN_NS}
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

FASHION_TRACKER_FORMULA_REPAIRS = {
    ("Revenu summary", "D74"): "SUM(D43:D73)",
    ("Total Rev", "AH4"): "'Revenu summary'!D74",
}

def repair_fashion_tracker_formulas(report_path: Path) -> int:
    """
    Repair verified formula regressions in Fashion Tracker without reserializing package:
    - 'Revenu summary'!D74: =SUM(D43:D73)
    - 'Total Rev'!AH4: ='Revenu summary'!D74
    Removes stale array formula attributes and cached values.
    """
    report_path = Path(report_path)
    if not report_path.is_file():
        return 0

    with zipfile.ZipFile(report_path, "r") as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}

    wb_xml = ET.fromstring(parts["xl/workbook.xml"])
    wb_rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rel_map = {r.get("Id"): r.get("Target") for r in wb_rels}

    sheet_files = {}
    for s in wb_xml.findall(".//main:sheet", NS):
        sname = s.get("name")
        rid = s.get(f"{{{DOC_REL_NS}}}id")
        sheet_files[sname] = rel_map[rid].split("/")[-1]

    repaired_count = 0
    grouped_repairs = {}
    for (sname, coord), formula in FASHION_TRACKER_FORMULA_REPAIRS.items():
        if sname in sheet_files:
            sfile = sheet_files[sname]
            part_name = f"xl/worksheets/{sfile}"
            grouped_repairs.setdefault(part_name, {})[coord] = formula

    for part_name, repairs in grouped_repairs.items():
        if part_name not in parts:
            continue
        root = ET.fromstring(parts[part_name])
        modified = False
        for coord, formula in repairs.items():
            cell = root.find(f".//main:c[@r='{coord}']", NS)
            if cell is not None:
                f_el = cell.find("main:f", NS)
                if f_el is None:
                    f_el = ET.SubElement(cell, f"{{{MAIN_NS}}}f")
                # Check if change needed
                if f_el.text != formula or f_el.attrib:
                    f_el.attrib.clear()
                    f_el.text = formula
                    # Remove old cached value
                    v_el = cell.find("main:v", NS)
                    if v_el is not None:
                        cell.remove(v_el)
                    cell.attrib.pop("t", None)
                    modified = True
                    repaired_count += 1
                    log.info("Repaired verified formula in %s!%s = %s", sname, coord, formula)
        if modified:
            parts[part_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    if repaired_count > 0:
        tmp_path = report_path.with_suffix(".tmp_repair")
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n, data in parts.items():
                zout.writestr(n, data)
        import os
        os.replace(tmp_path, report_path)
        log.info("Fashion Tracker formula repair committed: %d cell(s) repaired.", repaired_count)

    return repaired_count

def inspect_workbook_components(path: Path) -> dict:
    """Inspect all workbook OpenXML package components."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        pivots = [n for n in names if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml")]
        caches = [n for n in names if n.startswith("xl/pivotCache/pivotCacheDefinition") and n.endswith(".xml")]
        cache_records = [n for n in names if n.startswith("xl/pivotCache/pivotCacheRecords") and n.endswith(".xml")]
        charts = [n for n in names if n.startswith("xl/charts/chart") and n.endswith(".xml")]
        drawings = [n for n in names if n.startswith("xl/drawings/drawing") and n.endswith(".xml")]
        ext_links = [n for n in names if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml")]

        # Drawing ownership check
        sheet_drawings = {}
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in wb_rels}

        for s in wb_xml.findall(".//main:sheet", NS):
            sname = s.get("name")
            rid = s.get(f"{{{DOC_REL_NS}}}id")
            sfile = rel_map.get(rid, "").split("/")[-1]
            srel_file = f"xl/worksheets/_rels/{sfile}.rels"
            if srel_file in names:
                srels = ET.fromstring(z.read(srel_file))
                for r in srels:
                    if r.get("Type", "").endswith("/drawing"):
                        sheet_drawings.setdefault(sname, []).append(r.get("Target"))

        return {
            "total_parts": len(names),
            "pivots": len(pivots),
            "caches": len(caches),
            "cache_records": len(cache_records),
            "charts": len(charts),
            "drawings": len(drawings),
            "external_links": len(ext_links),
            "sheet_drawings": sheet_drawings,
        }
