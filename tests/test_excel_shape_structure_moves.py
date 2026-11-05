"""Shapes following their cells as whole rows and columns, or cells, are inserted and deleted, replayed from live
Excel (scripts/measure_shape_structure_moves.py, tests/fixtures/shape_structure_moves.json)."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/shape_structure_moves.json').read_text(encoding='utf-8'))
#: The probe as recorded.
SOURCE = RECORD['source']


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='StructureMovesProbe')
    return str(app.run('StructureMovesProbe.Probe')).split('@@@')


#: A rectangle moved by an inserted row, saved, opened again, moved by another and saved again: each time it reads
#: where the rows put it, a loaded shape's anchor in the file included.
ROUND_TRIP = '''
Public Function Trip(ByVal folder As String) As String
    Dim wb As Workbook, sh As Shape, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set sh = wb.Worksheets(1).Shapes.AddShape(msoShapeRectangle, 110, 20, 120, 40)
    wb.Worksheets(1).Rows(1).Insert
    wb.SaveAs folder & "{sep}a.xlsx", xlOpenXMLWorkbook: wb.Close False
    Set wb = Workbooks.Open(folder & "{sep}a.xlsx")
    Set sh = wb.Worksheets(1).Shapes(1)
    s = sh.Top & "," & sh.Height
    wb.Worksheets(1).Rows(4).Insert
    s = s & "|" & sh.Top & "," & sh.Height
    wb.SaveAs folder & "{sep}b.xlsx", xlOpenXMLWorkbook: wb.Close False
    Set wb = Workbooks.Open(folder & "{sep}b.xlsx")
    Trip = s & "|" & wb.Worksheets(1).Shapes(1).Top & "," & wb.Worksheets(1).Shapes(1).Height
    wb.Close False
End Function
'''


def test_moved_shapes_save_where_the_rows_put_them(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(ROUND_TRIP.replace('{sep}', os.sep), name='RoundTrip')
    # Inserted above, the rectangle moves down a row; inserted through it, it grows a row (as the fixture's cases).
    assert str(app.run('RoundTrip.Trip', str(tmp_path))) == '35,40|35,55|35,55'


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    _, answer = RECORD['cases'][index]
    assert answers[index] == answer
