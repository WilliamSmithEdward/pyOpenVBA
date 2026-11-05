"""Writing and reading a column costs each cell its own work, not a walk of the book. These count the walks, where a time would vary by machine."""
from __future__ import annotations

from typing import Any
from unittest.mock import patch

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._dimensions import SheetDimensions


def _app(source: str) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='H')
    return app


COLUMN = '''Function Probe() As String
    Dim ws As Worksheet, i As Long
    Set ws = ActiveSheet
    ws.Range("A1").Resize(200, 1).Value = 3
    For i = 1 To 200
        ws.Cells(i, 2).Formula = "=COUNTIF(A:A,A" & i & ")"
    Next
    Probe = ws.Cells(200, 2).Value
End Function
'''


def test_a_plain_cell_does_not_send_the_sheet_back_through_its_row_fonts() -> None:
    app = _app(COLUMN)
    real = SheetDimensions._work_out_growth
    calls: list[int] = []

    def counted(self: Any) -> None:
        calls.append(1)
        real(self)

    with patch.object(SheetDimensions, '_work_out_growth', counted):
        assert str(app.run('H.Probe')) == '200'
    assert len(calls) < 10


@pytest.mark.parametrize('write, grows', [
    ('ws.Range("A1").Font.Size = 24: ws.Range("A1").Value = "x"', True),
    ('ws.Range("A1").Value = "x": ws.Range("A1").Font.Size = 24', True),
    ('ws.Range("A1").Value = "x"', False),
])
def test_a_row_still_grows_by_the_font_of_what_is_written(write: str, grows: bool) -> None:
    lines = write.replace(': ', '\n    ')
    app = _app(f'Function Probe()\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    {lines}\n    Dim n As Long: n = ws.UsedRange.Rows.Count\n    Probe = ws.Rows(1).RowHeight\nEnd Function\n')
    assert (float(str(app.run('H.Probe'))) > 15.0) is grows


@pytest.mark.parametrize('formula', ['=SUM(A1:A3)', '=SUM(A:A)', '=SUM(1:1)', '=SUM(A1:Z1)', '=SUM(A1:A5000)', '=SUM(A1000:A1030)', '=SUM(A1023:A1025)'])
@pytest.mark.parametrize('place', ['A1', 'A3', 'A4', 'B1', 'Z1', 'A1024', 'A1030', 'A5000', 'XFD1', 'C5000'])
def test_a_formula_follows_a_cell_exactly_when_its_block_holds_it(formula: str, place: str) -> None:
    """Filed by column and stretch of rows, a block must still answer for each of its cells and for none outside it."""
    app = _app(f'''Function Probe()
    Dim ws As Worksheet, before As Double
    Set ws = ActiveSheet
    If ws.Range("{place}").Address(False, False) = "G7" Then Exit Function
    ws.Range("G7").Formula = "{formula}"
    before = ws.Range("G7").Value
    ws.Range("{place}").Value = 5
    Probe = ws.Range("G7").Value - before
End Function
''')
    held = {
        '=SUM(A1:A3)': {'A1', 'A3'}, '=SUM(A:A)': {'A1', 'A3', 'A4', 'A1024', 'A1030', 'A5000'}, '=SUM(1:1)': {'A1', 'B1', 'Z1', 'XFD1'},
        '=SUM(A1:Z1)': {'A1', 'B1', 'Z1'}, '=SUM(A1:A5000)': {'A1', 'A3', 'A4', 'A1024', 'A1030', 'A5000'}, '=SUM(A1000:A1030)': {'A1024', 'A1030'},
        '=SUM(A1023:A1025)': {'A1024'},
    }[formula]
    assert app.run('H.Probe') == (5 if place in held else 0)


def test_a_formula_given_a_new_block_leaves_the_old_one() -> None:
    app = _app('''Function Probe() As String
    Dim ws As Worksheet, s As String
    Set ws = ActiveSheet
    ws.Range("G7").Formula = "=SUM(A1:A3)"
    ws.Range("G7").Formula = "=SUM(B1:B3)"
    ws.Range("A1").Value = 5
    s = ws.Range("G7").Value
    ws.Range("B1").Value = 7
    Probe = s & "|" & ws.Range("G7").Value
End Function
''')
    assert app.run('H.Probe') == '0|7'
