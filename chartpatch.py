"""Preserve original Excel charts and drawings across openpyxl saves.

Worksheets are matched by their actual names, not their XML part numbers.
Original drawing parts and relationships are restored after openpyxl saves.
"""

import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

CHART_PART = re.compile(r"^xl/(charts|drawings)/")

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", DOC_REL_NS)


def _part_path(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def _relationships(data):
    if not data:
        return ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    return ET.fromstring(data)


def _sheet_map(parts):
    """Return worksheet names mapped to their XML part paths."""
    workbook = ET.fromstring(parts["xl/workbook.xml"])
    rels = _relationships(parts["xl/_rels/workbook.xml.rels"])

    targets = {
        rel.get("Id"): _part_path("xl/workbook.xml", rel.get("Target"))
        for rel in rels
    }

    result = {}
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        name = sheet.get("name")
        rid = sheet.get(f"{{{DOC_REL_NS}}}id")
        if rid in targets:
            result[name] = targets[rid]
    return result


def _rels_path(sheet_path):
    directory, filename = posixpath.split(sheet_path)
    return f"{directory}/_rels/{filename}.rels"


def _next_rid(rels):
    used = {rel.get("Id") for rel in rels}
    n = 1
    while f"rId{n}" in used:
        n += 1
    return f"rId{n}"


def snapshot(xlsx_path):
    snap = {"parts": {}, "overrides": {}, "sheets": {}}

    with zipfile.ZipFile(xlsx_path) as z:
        parts = {name: z.read(name) for name in z.namelist() if not name.endswith("/")}

    for name, data in parts.items():
        if CHART_PART.match(name):
            snap["parts"][name] = data

    content_types = ET.fromstring(parts["[Content_Types].xml"])
    for element in content_types:
        part_name = element.get("PartName", "").lstrip("/")
        if CHART_PART.match(part_name):
            snap["overrides"][part_name] = element.get("ContentType")

    for sheet_name, sheet_path in _sheet_map(parts).items():
        sheet_xml = ET.fromstring(parts[sheet_path])
        drawing = sheet_xml.find(f"{{{MAIN_NS}}}drawing")

        if drawing is None:
            continue

        drawing_rid = drawing.get(f"{{{DOC_REL_NS}}}id")
        relpath = _rels_path(sheet_path)
        rels = _relationships(parts.get(relpath))

        for rel in rels:
            if rel.get("Id") == drawing_rid and rel.get("Type", "").endswith("/drawing"):
                snap["sheets"][sheet_name] = {
                    "target": _part_path(sheet_path, rel.get("Target")),
                }
                break

    return snap


def restore(xlsx_path, snap):
    if not snap["parts"]:
        return 0

    with zipfile.ZipFile(xlsx_path) as z:
        parts = {name: z.read(name) for name in z.namelist() if not name.endswith("/")}

    sheets = _sheet_map(parts)

    for sheet_name, sheet_path in sheets.items():
        sheet_root = ET.fromstring(parts[sheet_path])
        relpath = _rels_path(sheet_path)
        rels_root = _relationships(parts.get(relpath))

        original = snap["sheets"].get(sheet_name)

        # Remove existing drawing relationships from this worksheet.
        for rel in list(rels_root):
            if rel.get("Type", "").endswith("/drawing"):
                rels_root.remove(rel)

        # Remove existing drawing elements.
        for drawing in list(sheet_root.findall(f"{{{MAIN_NS}}}drawing")):
            sheet_root.remove(drawing)

        if original:
            rid = _next_rid(rels_root)

            ET.SubElement(
                rels_root,
                f"{{{PKG_REL_NS}}}Relationship",
                {
                    "Type": DOC_REL_NS + "/drawing",
                    "Target": "/" + original["target"],
                    "Id": rid,
                },
            )

            drawing = ET.Element(
                f"{{{MAIN_NS}}}drawing",
                {f"{{{DOC_REL_NS}}}id": rid},
            )

            # Insert before extension lists and table parts.
            children = list(sheet_root)
            index = len(children)

            for i, child in enumerate(children):
                local = child.tag.split("}")[-1]
                if local in ("tableParts", "extLst"):
                    index = i
                    break

            sheet_root.insert(index, drawing)

        parts[sheet_path] = ET.tostring(
            sheet_root, encoding="utf-8", xml_declaration=True
        )

        parts[relpath] = ET.tostring(
            rels_root, encoding="utf-8", xml_declaration=True
        )

    # Restore original chart and drawing components.
    for name in list(parts):
        if CHART_PART.match(name):
            del parts[name]

    parts.update(snap["parts"])

    # Restore content-type declarations.
    ct_root = ET.fromstring(parts["[Content_Types].xml"])
    for element in list(ct_root):
        if CHART_PART.match(element.get("PartName", "").lstrip("/")):
            ct_root.remove(element)

    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    for part_name, content_type in snap["overrides"].items():
        ET.SubElement(
            ct_root,
            f"{{{ct_ns}}}Override",
            {"PartName": "/" + part_name, "ContentType": content_type},
        )

    parts["[Content_Types].xml"] = ET.tostring(
        ct_root, encoding="utf-8", xml_declaration=True
    )

    temp = xlsx_path + ".chartpatch_test"

    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)

    os.replace(temp, xlsx_path)

    return len(snap["parts"])
