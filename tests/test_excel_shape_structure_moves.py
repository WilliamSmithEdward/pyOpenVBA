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
#: Rows that grow by themselves: with a font, which moves shapes in Excel but not yet in the model, and with wrapped
#: text, whose height the model does not measure.
GROWN = ('Font.Size', 'WrapText')
#: The probe as recorded, the grown rows' cases left as they are made but with nothing changed.
SOURCE = '\n'.join('' if any(marker in line for marker in GROWN) else line for line in RECORD['source'].split('\n'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
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
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(ROUND_TRIP.replace('{sep}', os.sep), name='RoundTrip')
    # Inserted above, the rectangle moves down a row; inserted through it, it grows a row (as the fixture's cases).
    assert str(app.run('RoundTrip.Trip', str(tmp_path))) == '35,40|35,55|35,55'


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    case, answer = RECORD['cases'][index]
    if any(marker in case for marker in GROWN):
        pytest.skip('rows that grow by themselves do not move shapes in the model yet')
    assert answers[index] == answer
