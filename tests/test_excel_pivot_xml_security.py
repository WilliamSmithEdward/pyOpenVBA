"""Pivot package readers refuse DTDs before processing declared entities."""
from pathlib import Path
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication

SEED = Path(__file__).parent / 'fixtures/pivots/native-simple.xlsx'


@pytest.mark.parametrize('part', ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
                                  'xl/worksheets/_rels/sheet1.xml.rels',
                                  'xl/pivotTables/pivotTable1.xml',
                                  'xl/pivotTables/_rels/pivotTable1.xml.rels',
                                  'xl/pivotCache/pivotCacheDefinition1.xml',
                                  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
                                  'xl/pivotCache/pivotCacheRecords1.xml'])
@pytest.mark.parametrize('encoding', ['utf-8', 'utf-16'])
@pytest.mark.parametrize('doctype', ['<!DOCTYPE r [<!ENTITY x "expanded">]>',
                                   '<!DOCTYPE r SYSTEM "file:///not-read.dtd">',
                                   '<!DOCTYPE r [<!ENTITY % x SYSTEM "https://example.invalid/not-fetched">%x;]>'])
def test_pivot_xml_rejects_declared_entities(tmp_path: Path, part: str, encoding: str, doctype: str) -> None:
    path = tmp_path / 'hostile.xlsx'
    with ZipFile(SEED) as original, ZipFile(path, 'w') as changed:
        for name in original.namelist():
            payload = original.read(name)
            if name == part:
                xml = payload.decode('utf-8')
                xml = xml[xml.index('?>') + 2:]
                payload = (f'<?xml version="1.0" encoding="{encoding}"?>' + doctype + xml).encode(encoding)
            changed.writestr(name, payload)
    with pytest.raises(ValueError, match='DTD'):
        app = ExcelApplication.open(path, with_vba=False)
        # Cache records are parsed on demand, rather than during workbook import.
        app.evaluate('Worksheets(1).PivotTables(1).PivotFields(2).PivotItems.Count')
