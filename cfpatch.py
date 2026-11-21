"""Preserve extended Excel conditional formatting by worksheet name."""

import os
import re
import zipfile
import xml.etree.ElementTree as ET

from chartpatch_fixed import _sheet_map

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

TAIL_RE = re.compile(
    r'<extLst\b[^>]*>(?:(?!</extLst>).)*?'
    r'x14:conditionalFormattings.*?</extLst>',
    re.S,
)
RULE_RE = re.compile(
    r'<cfRule\b[^>]*(?:/>|>.*?</cfRule>)',
    re.S,
)
INLINE_RE = re.compile(r'<extLst\b[^>]*>.*?</extLst>', re.S)
PRIO_RE = re.compile(r'priority="(\d+)"')


def snapshot(xlsx_path):
    result = {}

    with zipfile.ZipFile(xlsx_path) as z:
        parts = {
            name: z.read(name)
            for name in z.namelist()
            if not name.endswith("/")
        }

    for sheet_name, sheet_path in _sheet_map(parts).items():
        xml = parts[sheet_path].decode("utf-8")

        if "x14:conditionalFormattings" not in xml:
            continue

        match = TAIL_RE.search(xml)
        tail = match.group(0) if match else ""

        by_priority = {}

        for rule_match in RULE_RE.finditer(xml):
            rule = rule_match.group(0)
            ext = INLINE_RE.search(rule)
            priority = PRIO_RE.search(rule)

            if ext and priority:
                by_priority[priority.group(1)] = ext.group(0)

        if tail or by_priority:
            result[sheet_name] = {
                "tail": tail,
                "by_priority": by_priority,
            }

    return result


def _restore_sheet(xml, data):
    links = data["by_priority"]

    def fix_rule(match):
        rule = match.group(0)

        if "<extLst" in rule:
            return rule

        priority = PRIO_RE.search(rule)

        if not priority:
            return rule

        ext = links.get(priority.group(1))

        if not ext:
            return rule

        if rule.endswith("/>"):
            return rule[:-2] + ">" + ext + "</cfRule>"

        return rule[:-len("</cfRule>")] + ext + "</cfRule>"

    xml = RULE_RE.sub(fix_rule, xml)

    if data["tail"] and "x14:conditionalFormattings" not in xml:
        # Insert the extension list at the end of the worksheet.
        xml = xml.replace(
            "</worksheet>",
            data["tail"] + "</worksheet>",
        )

    return xml


def restore(xlsx_path, snap):
    if not snap:
        return 0

    with zipfile.ZipFile(xlsx_path) as z:
        parts = {
            name: z.read(name)
            for name in z.namelist()
            if not name.endswith("/")
        }

    sheet_map = _sheet_map(parts)
    patched = 0

    for sheet_name, data in snap.items():
        sheet_path = sheet_map.get(sheet_name)

        if not sheet_path:
            raise ValueError(
                f"Worksheet missing during CF restore: {sheet_name}"
            )

        xml = parts[sheet_path].decode("utf-8")

        if "x14:conditionalFormattings" in xml:
            continue

        updated = _restore_sheet(xml, data)

        if updated != xml:
            # Validate XML before writing.
            ET.fromstring(updated.encode("utf-8"))
            parts[sheet_path] = updated.encode("utf-8")
            patched += 1

    temp = xlsx_path + ".cfpatch_fixed"

    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in parts.items():
            z.writestr(name, blob)

    os.replace(temp, xlsx_path)

    return patched
