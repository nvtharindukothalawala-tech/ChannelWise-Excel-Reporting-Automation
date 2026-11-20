"""
Selective XLSX Package Updater (Architecture B).
Preserves 100% of unedited worksheets, PivotTables, PivotCaches, charts, drawings,
external links, conditional formatting, styles, and calculation components.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from hashlib import sha256
from pathlib import Path
from openpyxl.formula.translate import Translator
from .date_utils import to_excel_serial, to_date

log = logging.getLogger("reports")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"main": MAIN_NS}
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def col_letter(col_idx: int) -> str:
    """Convert 1-based column index to letter (1 -> A, 27 -> AA)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def col_number(col_str: str) -> int:
    """Convert column letters to 1-based index (A -> 1, AA -> 27)."""
    num = 0
    for c in col_str:
        num = num * 26 + (ord(c.upper()) - ord("A") + 1)
    return num

def cell_coord(coord: str) -> tuple[int, int]:
    """Return (col_num, row_num) from coordinate string like 'AB12'."""
    m = re.match(r"([A-Z]+)(\d+)", coord)
    if not m:
        raise ValueError(f"Invalid coordinate: {coord}")
    return col_number(m.group(1)), int(m.group(2))

def get_sheet_file_map(zip_handle: zipfile.ZipFile) -> dict[str, str]:
    """Map logical worksheet name -> internal filename ('sheetX.xml')."""
    wb_xml = ET.fromstring(zip_handle.read("xl/workbook.xml"))
    wb_rels = ET.fromstring(zip_handle.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.get("Id"): r.get("Target") for r in wb_rels}
    mapping = {}
    for s in wb_xml.findall(".//main:sheet", NS):
        sname = s.get("name")
        rid = s.get(f"{{{DOC_REL_NS}}}id")
        target = rel_map.get(rid, "").split("/")[-1]
        mapping[sname] = target
    return mapping

def update_worksheet_xml(
    sheet_xml_bytes: bytes,
    data_start_row: int,
    paste_start_col: int,
    new_rows: list[list[any]],
    skip_cols: list[int] | None = None,
    max_cols: int | None = None,
    date_target_header: str | None = None,
    source_columns: list[str | None] | None = None,
    replace_month: tuple[int, int] | None = None,
) -> tuple[bytes, int]:
    """
    Surgically updates a worksheet XML.
    - Preserves rows 1 .. data_start_row - 1.
    - If replace_month is specified: extracts historical rows belonging to other months,
      replaces only target month rows, and appends new_rows.
    - Stretches helper formulas across all rows from col 1 .. paste_start_col - 1.
    - Writes data values starting at paste_start_col.
    - Preserves 100% of formatting, conditional formatting, margins, views, cols, etc.
    """
    root = ET.fromstring(sheet_xml_bytes)
    sheetData = root.find("main:sheetData", NS)
    if sheetData is None:
        raise ValueError("Missing sheetData in worksheet XML")

    # 1. First pass: find shared formula definitions across the sheet
    shared_defs = {}
    for row in sheetData.findall("main:row", NS):
        for c in row.findall("main:c", NS):
            f = c.find("main:f", NS)
            if f is not None and f.get("t") == "shared" and f.get("si") and f.text:
                shared_defs[f.get("si")] = {"formula": f.text, "origin_cell": c.get("r")}

    # 2. Capture template row info from row data_start_row
    template_row_el = sheetData.find(f'main:row[@r="{data_start_row}"]', NS)
    if template_row_el is None:
        raise ValueError(f"Template row {data_start_row} not found in sheet")

    helper_templates = {}
    raw_col_styles = {}
    for c in template_row_el.findall("main:c", NS):
        r_coord = c.get("r")
        c_num, r_num = cell_coord(r_coord)
        s_idx = c.get("s")
        if c_num < paste_start_col:
            f_el = c.find("main:f", NS)
            f_text = None
            f_attr = {}
            origin_ref = r_coord
            if f_el is not None:
                f_attr = dict(f_el.attrib)
                if f_el.get("t") == "shared":
                    si = f_el.get("si")
                    sdef = shared_defs.get(si)
                    if sdef:
                        f_text = sdef["formula"]
                        origin_ref = sdef["origin_cell"]
                    else:
                        f_text = f_el.text
                else:
                    f_text = f_el.text
            helper_templates[c_num] = {
                "formula": f_text,
                "formula_attr": f_attr,
                "style": s_idx,
                "origin_ref": origin_ref,
                "origin_col_letter": col_letter(c_num),
            }
        else:
            rel_col = c_num - paste_start_col
            raw_col_styles[rel_col] = s_idx

    # 3. If replace_month: extract existing historical rows to keep
    kept_historical_rows = []
    if replace_month and date_target_header and source_columns:
        date_pos = source_columns.index(date_target_header)
        width = len(source_columns)
        for r_el in sheetData.findall("main:row", NS):
            r_num = int(r_el.get("r"))
            if r_num < data_start_row:
                continue
            cells_by_col = {}
            for c in r_el.findall("main:c", NS):
                cn, _ = cell_coord(c.get("r"))
                is_el = c.find("main:is/main:t", NS)
                v_el = c.find("main:v", NS)
                if is_el is not None:
                    val = is_el.text
                elif v_el is not None:
                    val = v_el.text
                else:
                    val = None
                cells_by_col[cn] = val

            row_vals = [cells_by_col.get(paste_start_col + j) for j in range(width)]
            if all(v is None or str(v).strip() == "" for v in row_vals):
                continue

            date_val = row_vals[date_pos]
            d = to_date(date_val)
            if d is not None and (d.year, d.month) == replace_month:
                # Belongs to target month being replaced
                continue
            else:
                # Retain historical row
                kept_historical_rows.append(row_vals)

    all_rows = kept_historical_rows + new_rows

    # 4. Collect header rows (< data_start_row)
    header_rows = []
    for r in list(sheetData):
        r_num = int(r.get("r"))
        if r_num < data_start_row:
            header_rows.append(r)

    # 5. Build new row elements
    drop_set = {int(c) - 1 for c in (skip_cols or [])}
    new_row_elements = []

    for idx, row_vals in enumerate(all_rows):
        current_r = data_start_row + idx
        row_el = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(current_r)})

        # A. Helper columns (1 .. paste_start_col - 1)
        for col_idx in sorted(helper_templates.keys()):
            ht = helper_templates[col_idx]
            c_letter = col_letter(col_idx)
            cell_ref = f"{c_letter}{current_r}"
            c_attrib = {"r": cell_ref}
            if ht["style"]:
                c_attrib["s"] = ht["style"]

            if ht["formula"]:
                f_raw = ht["formula"]
                f_with_eq = f_raw if f_raw.startswith("=") else "=" + f_raw
                origin_ref = ht["origin_ref"]
                dest_ref = cell_ref
                try:
                    translated = Translator(f_with_eq, origin=origin_ref).translate_formula(dest_ref)
                    if translated.startswith("="):
                        translated = translated[1:]
                except Exception:
                    origin_r = cell_coord(origin_ref)[1]
                    translated = re.sub(rf"([A-Za-z]+){origin_r}\b", rf"\g<1>{current_r}", f_raw)

                c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
                f_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}f")
                f_el.text = translated
                ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
            else:
                c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)

        # B. Data columns
        target_col_offset = 0
        limit = min(len(row_vals), max_cols) if max_cols else len(row_vals)
        for src_col, val in enumerate(row_vals[:limit]):
            if src_col in drop_set:
                continue
            cur_col_num = paste_start_col + target_col_offset
            c_letter = col_letter(cur_col_num)
            cell_ref = f"{c_letter}{current_r}"
            c_attrib = {"r": cell_ref}

            s_style = raw_col_styles.get(target_col_offset)
            if s_style:
                c_attrib["s"] = s_style

            if val is None or val == "":
                c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                c_attrib["t"] = "n"
                c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
                v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                v_el.text = str(val)
            elif isinstance(val, (datetime, date)):
                serial = to_excel_serial(val)
                c_attrib["t"] = "n"
                c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
                v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                v_el.text = f"{serial:.6f}".rstrip("0").rstrip(".")
            else:
                text_val = str(val)
                if re.fullmatch(r"-?\d+(?:\.\d+)?", text_val.strip()):
                    c_attrib["t"] = "n"
                    c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
                    v_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}v")
                    v_el.text = text_val.strip()
                else:
                    c_attrib["t"] = "inlineStr"
                    c_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", c_attrib)
                    is_el = ET.SubElement(c_el, f"{{{MAIN_NS}}}is")
                    t_el = ET.SubElement(is_el, f"{{{MAIN_NS}}}t")
                    t_el.text = text_val
            target_col_offset += 1

        new_row_elements.append(row_el)

    # 6. Replace sheetData contents
    sheetData.clear()
    for r in header_rows:
        sheetData.append(r)
    for r in new_row_elements:
        sheetData.append(r)

    # 7. Update dimension ref
    dim_el = root.find("main:dimension", NS)
    if dim_el is not None:
        last_r = data_start_row + len(all_rows) - 1
        last_c = paste_start_col + (target_col_offset if all_rows else 1)
        dim_el.set("ref", f"A1:{col_letter(last_c)}{max(last_r, data_start_row)}")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), len(all_rows)

def update_table_ref(table_xml_bytes: bytes, new_last_row: int) -> bytes:
    """Update table ref attribute if table exists on sheet."""
    root = ET.fromstring(table_xml_bytes)
    old_ref = root.get("ref")
    if old_ref and ":" in old_ref:
        start_cell, end_cell = old_ref.split(":")
        start_col, start_row = cell_coord(start_cell)
        end_col, _ = cell_coord(end_cell)
        new_ref = f"{col_letter(start_col)}{start_row}:{col_letter(end_col)}{max(new_last_row, start_row + 1)}"
        root.set("ref", new_ref)
        autofilter = root.find("main:autoFilter", NS)
        if autofilter is not None:
            autofilter.set("ref", new_ref)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

def verify_file_unlocked(file_path: Path):
    """Raise PermissionError if file is locked by Excel or OneDrive."""
    if not file_path.exists():
        return
    try:
        with open(file_path, "r+b"):
            pass
    except (PermissionError, OSError) as exc:
        raise PermissionError(f"File {file_path.name} is locked by another process (e.g. Excel). Please close it.") from exc

def calculate_sha256(path: Path) -> str:
    """Calculate SHA-256 hash of a file."""
    h = sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()
