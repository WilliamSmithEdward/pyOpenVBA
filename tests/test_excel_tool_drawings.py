"""A drawing not spelled with xdr:, as openpyxl writes it and under a prefix of another writer's: Excel reads its shapes and charts alike
(scripts/measure_tool_drawings.py, tests/fixtures/tool_drawings.json), and a save with no edit keeps the part as it came."""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._prefixed_parts import drawing_spelled

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'tool_drawings.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('name', RECORD['answers'])
def test_each_read_answers_as_excel_did(name: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    theirs = RECORD['answers'][name].split('|')
    ours = str(app.run('P.Probe', str(FIXTURES / name))).split('|')
    assert dict(zip(RECORD['reads'], ours)) == dict(zip(RECORD['reads'], theirs))


@pytest.mark.parametrize('name', RECORD['answers'])
def test_a_save_with_no_edit_keeps_the_drawing_as_it_came(name: str, tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Save(ByVal path As String, ByVal out As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
End Sub
''', name='P')
    book, out = tmp_path / 'in.xlsx', tmp_path / 'out.xlsx'
    shutil.copy(FIXTURES / name, book)
    app.run('P.Save', str(book), str(out))
    with zipfile.ZipFile(book) as before, zipfile.ZipFile(out) as after:
        assert after.read('xl/drawings/drawing1.xml') == before.read('xl/drawings/drawing1.xml')


@pytest.mark.parametrize('xml', [
    '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><graphic xmlns="urn:other"/></wsDr>',
    '<other xmlns="urn:other"><wsDr/></other>',
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:pic/></xdr:wsDr>',
])
def test_what_cannot_be_spelled_safely_is_left_alone(xml: str) -> None:
    assert drawing_spelled(xml) == xml
