"""Preserve original Excel external-link XML in a staged XLSX workbook.

Only use on a staged copy created from the SAME original workbook.
Never edits the original XLSX.  Restore is atomic and fails closed.
"""
from __future__ import annotations

import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree as ET
from zipfile import ZipFile

XL_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _names(z: ZipFile) -> set[str]:
    return {n for n in z.namelist() if n.startswith('xl/externalLinks/') and not n.endswith('/')}


def _validate_refs(z: ZipFile) -> None:
    """Ensure all externalBook relationship IDs exist in the matching rels part."""
    for xml_name in sorted(n for n in _names(z) if n.startswith('xl/externalLinks/externalLink') and n.endswith('.xml')):
        root = ET.fromstring(z.read(xml_name))
        book = root.find(f'{{{MAIN_NS}}}externalBook')
        if book is None:
            continue  # DDE/OLE external links need a different relationship model
        rid = book.get(f'{{{XL_REL_NS}}}id')
        if not rid:
            raise ValueError(f'{xml_name}: externalBook has no relationship ID')
        fname = xml_name.rsplit('/', 1)[-1]
        rel_name = f'xl/externalLinks/_rels/{fname}.rels'
        if rel_name not in z.namelist():
            raise ValueError(f'{xml_name}: relationship file missing')
        relationships = ET.fromstring(z.read(rel_name))
        matches = [r for r in relationships.findall(f'{{{PKG_REL_NS}}}Relationship') if r.get('Id') == rid]
        if len(matches) != 1 or not matches[0].get('Target'):
            raise ValueError(f'{xml_name}: relationship {rid} missing or ambiguous')


def snapshot(original: str | Path) -> dict[str, bytes]:
    """Read original external-link XML and relationship bytes, checking references."""
    with ZipFile(original, 'r') as z:
        if z.testzip() is not None:
            raise ValueError('Original workbook ZIP integrity check failed')
        _validate_refs(z)
        return {name: z.read(name) for name in sorted(_names(z))}


def restore(staged: str | Path, components: dict[str, bytes]) -> int:
    """Atomically restore exact original external-link bytes into staged XLSX.

    Fails if component names differ. This prevents mixing unrelated link layouts.
    """
    staged = Path(staged)
    if not components:
        with ZipFile(staged) as z:
            if _names(z):
                raise ValueError('Staged workbook unexpectedly acquired external links')
        return 0

    temp = None
    try:
        with ZipFile(staged, 'r') as source:
            actual = _names(source)
            if actual != set(components):
                raise ValueError(f'External-link part names changed: missing={sorted(set(components)-actual)}, extra={sorted(actual-set(components))}')
            with NamedTemporaryFile(prefix='external_links_', suffix='.tmp', dir=staged.parent, delete=False) as handle:
                temp = Path(handle.name)
            with ZipFile(temp, 'w') as target:
                for info in source.infolist():
                    target.writestr(info, components.get(info.filename, source.read(info.filename)))

        with ZipFile(temp, 'r') as check:
            if check.testzip() is not None:
                raise ValueError('Patched workbook ZIP CRC test failed')
            for name, data in components.items():
                if check.read(name) != data:
                    raise ValueError(f'External-link component mismatch: {name}')
            _validate_refs(check)

        os.replace(temp, staged)
        return len(components)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
