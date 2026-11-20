"""
Preserve charts that openpyxl deletes when a workbook is re-saved.

Excel keeps a chart in several pieces:
  xl/charts/chartN.xml        the chart itself (plus style/colors parts)
  xl/drawings/drawingN.xml    where it sits on the sheet
  the sheet's .rels           links the sheet to the drawing
  <drawing r:id="..."/>       one element inside the worksheet XML
  [Content_Types].xml         declares each part

openpyxl rewrites the chart and drawing parts from its own simplified model
and drops the style and colour parts altogether, so charts survive but lose
their formatting. This module snapshots the original pieces and puts them
back afterwards, replacing whatever openpyxl produced.

Charts point at their data with formulas such as SalesRepItem!$B$2:$B$40,
so they keep working as long as the sheet layout does not move. Only cell
values change during a run, so the snapshot stays valid.
"""

import os, re, shutil, zipfile

CHART_PART = re.compile(r"^xl/(charts|drawings)/")
DRAW_REL = re.compile(
    r'<Relationship[^>]*Type="[^"]*/drawing"[^>]*/>')
DRAW_EL = re.compile(r'<drawing\b[^>]*r:id="[^"]+"[^>]*/>')
# <drawing> must sit after headerFooter and before these, per the schema
TAIL_TAGS = ("<tableParts", "<extLst", "</worksheet>")


def snapshot(xlsx_path):
    """Collect every chart-related piece of the workbook."""
    snap = {"parts": {}, "overrides": [], "sheets": {}}
    with zipfile.ZipFile(xlsx_path) as z:
        ct = z.read("[Content_Types].xml").decode("utf-8")
        for m in re.finditer(r'<Override PartName="([^"]+)"[^>]*/>', ct):
            part = m.group(1).lstrip("/")
            if CHART_PART.match(part):
                snap["overrides"].append(m.group(0))

        for name in z.namelist():
            if name.endswith("/"):
                continue
            if CHART_PART.match(name):
                snap["parts"][name] = z.read(name)

        for name in z.namelist():
            m = re.match(r"xl/worksheets/(sheet\d+\.xml)$", name)
            if not m:
                continue
            xml = z.read(name).decode("utf-8")
            el = DRAW_EL.search(xml)
            if not el:
                continue
            relpath = f"xl/worksheets/_rels/{m.group(1)}.rels"
            rels = z.read(relpath).decode("utf-8") if relpath in z.namelist() else ""
            keep = DRAW_REL.findall(rels)
            if keep:
                snap["sheets"][name] = {"rel": keep[0], "relpath": relpath}
    return snap


def _free_rid(rels_xml):
    used = {int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)}
    n = 1
    while n in used:
        n += 1
    return f"rId{n}"


def _insert_drawing(xml, rid):
    if DRAW_EL.search(xml):
        return xml
    # the r: prefix must be declared on the element itself, because openpyxl
    # does not always put xmlns:r on the worksheet root
    el = ('<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument'
          f'/2006/relationships" r:id="{rid}"/>')
    for tag in TAIL_TAGS:
        i = xml.find(tag)
        if i != -1:
            return xml[:i] + el + xml[i:]
    return xml


def restore(xlsx_path, snap):
    """Put the charts back into a workbook openpyxl has just written."""
    if not snap["parts"]:
        return 0

    tmp = xlsx_path + ".chartpatch"
    existing = {}
    with zipfile.ZipFile(xlsx_path) as zin:
        names = set(zin.namelist())
        for n in names:
            existing[n] = zin.read(n)

    # 1. worksheets: re-insert the <drawing> element and its relationship
    restored = 0
    for sheet, info in snap["sheets"].items():
        if sheet not in existing:
            continue
        rels_xml = existing.get(info["relpath"], "").decode("utf-8") if info["relpath"] in existing else (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        if "/drawing" in rels_xml and DRAW_REL.search(rels_xml):
            rid = re.search(r'Id="([^"]+)"', DRAW_REL.search(rels_xml).group(0)).group(1)
        else:
            rid = _free_rid(rels_xml)
            rel = re.sub(r'Id="[^"]+"', f'Id="{rid}"', info["rel"], count=1)
            rels_xml = rels_xml.replace("</Relationships>", rel + "</Relationships>")
        existing[info["relpath"]] = rels_xml.encode("utf-8")

        xml = existing[sheet].decode("utf-8")
        new = _insert_drawing(xml, rid)
        if new != xml:
            existing[sheet] = new.encode("utf-8")

    # 2. the chart and drawing parts themselves, replacing openpyxl's versions
    for name, blob in snap["parts"].items():
        existing[name] = blob
        restored += 1

    # 3. content types
    ct = existing["[Content_Types].xml"].decode("utf-8")
    for ov in snap["overrides"]:
        part = re.search(r'PartName="([^"]+)"', ov).group(1)
        if f'PartName="{part}"' not in ct:
            ct = ct.replace("</Types>", ov + "</Types>")
    existing["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for name, blob in existing.items():
            zout.writestr(name, blob)
    os.replace(tmp, xlsx_path)
    return restored
