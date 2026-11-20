"""
Preserve extended conditional formatting (data bars with custom colours,
icon sets) that openpyxl silently drops on save.

Excel stores these in two places inside a worksheet's XML:
  1. an <extLst> at the end of the sheet holding <x14:conditionalFormattings>
  2. a small <extLst> inside individual <cfRule> elements, linking each rule
     to its extended definition by GUID

openpyxl understands neither, so both vanish when a workbook is re-saved.
This module snapshots them from the original file before the run and splices
them back into the saved file afterwards.

Rules are matched on their `priority` attribute, which is unique within a
worksheet and preserved by openpyxl. Only cell values change during a run,
never the formatting, so the snapshot stays valid.
"""

import os, re, zipfile

WS_RE     = re.compile(r'xl/worksheets/sheet\d+\.xml$')
TAIL_RE   = re.compile(r'<extLst>(?:(?!</extLst>).)*?x14:conditionalFormattings.*?</extLst>', re.S)
RULE_RE   = re.compile(r'<cfRule\b[^>]*/>|<cfRule\b[^>]*>.*?</cfRule>', re.S)
INLINE_RE = re.compile(r'<extLst>.*?</extLst>', re.S)
PRIO_RE   = re.compile(r'priority="(\d+)"')


def snapshot(xlsx_path):
    """Return {sheet_xml_name: {'tail': str, 'by_priority': {prio: extLst}}}."""
    snap = {}
    with zipfile.ZipFile(xlsx_path) as z:
        for name in z.namelist():
            if not WS_RE.search(name):
                continue
            xml = z.read(name).decode("utf-8")
            if "x14:conditionalFormattings" not in xml:
                continue

            m = TAIL_RE.search(xml)
            tail = m.group(0) if m else ""

            by_priority = {}
            for rm in RULE_RE.finditer(xml):
                rule = rm.group(0)
                ext = INLINE_RE.search(rule)
                prio = PRIO_RE.search(rule)
                if ext and prio:
                    by_priority[prio.group(1)] = ext.group(0)

            if tail or by_priority:
                snap[name] = {"tail": tail, "by_priority": by_priority}
    return snap


def _restore_sheet(xml, data):
    links = data["by_priority"]

    def fix_rule(rm):
        rule = rm.group(0)
        if "<extLst>" in rule:
            return rule
        prio = PRIO_RE.search(rule)
        if not prio or prio.group(1) not in links:
            return rule
        ext = links[prio.group(1)]
        if rule.endswith("/>"):
            return rule[:-2] + ">" + ext + "</cfRule>"
        return rule[: -len("</cfRule>")] + ext + "</cfRule>"

    xml = RULE_RE.sub(fix_rule, xml)

    if data["tail"] and "x14:conditionalFormattings" not in xml:
        xml = xml.replace("</worksheet>", data["tail"] + "</worksheet>")
    return xml


def restore(xlsx_path, snap):
    """Rewrite xlsx_path with the snapshotted formatting spliced back in."""
    if not snap:
        return 0
    tmp = xlsx_path + ".cfpatch"
    patched = 0
    with zipfile.ZipFile(xlsx_path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename in snap:
                xml = blob.decode("utf-8")
                if "x14:conditionalFormattings" not in xml:
                    blob = _restore_sheet(xml, snap[item.filename]).encode("utf-8")
                    patched += 1
            zout.writestr(item, blob)
    os.replace(tmp, xlsx_path)
    return patched
