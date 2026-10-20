"""Independent OPC assertions for generated conditional witness sheets."""
from __future__ import annotations

import posixpath
from pathlib import Path
from zipfile import ZipFile

from pyopenvba.apps.excel._charts_file import xml_tree

SHEET = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'


def assert_witness_part(path: Path) -> None:
    with ZipFile(path) as package:
        workbook = xml_tree(package.read('xl/workbook.xml').decode('utf-8'))
        witnesses = [sheet for sheet in workbook.findall(SHEET + 'sheets/' + SHEET + 'sheet')
                     if sheet.get('name') == 'Witness']
        assert len(witnesses) == 1
        identifier = witnesses[0].get(RID)
        assert identifier
        relationships = xml_tree(package.read('xl/_rels/workbook.xml.rels').decode('utf-8'))
        matches = [relation for relation in relationships.findall(REL + 'Relationship')
                   if relation.get('Id') == identifier]
        assert len(matches) == 1
        relation = matches[0]
        assert relation.get('Type', '').endswith('/worksheet')
        assert relation.get('TargetMode', 'Internal') == 'Internal'
        target = relation.get('Target', '')
        assert target
        part = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        assert part in package.namelist()
        assert xml_tree(package.read(part).decode('utf-8')).tag == SHEET + 'worksheet'
        content_types = xml_tree(package.read('[Content_Types].xml').decode('utf-8'))
        declarations = [entry for entry in content_types.findall(CT + 'Override')
                        if entry.get('PartName') == '/' + part]
        assert len(declarations) == 1
        assert declarations[0].get('ContentType') == 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
