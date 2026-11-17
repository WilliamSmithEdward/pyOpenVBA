"""PivotTable round 57102's case (tests/fixtures/pivot_table_rounds/, on branch verify02-pivot-rounds): Excel moves a
table down for a row inserted above it, and an edit before every table, beside it, or on another sheet leaves it where
it was. tests/test_excel_pivot_shifts.py replays the measured rules."""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication

#: A table at C3 on Second: rows 3 to 6, columns C to D, from Main's keys and values.
SETUP = '''    Dim wb As Object, ws As Object, ws2 As Object, pt As Object
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Main"
    Set ws2 = wb.Worksheets.Add(After:=ws): ws2.Name = "Second"
    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws2.Range("C3"), "PT")
    pt.PivotFields("k").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum
'''


def _run(edit: str) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function One() As String\n' + SETUP + f'    {edit}\n'
                   '    One = ws2.PivotTables(1).TableRange1.Address & " " & ws2.Range("C3").Value\nEnd Function\n', name='P')
    return app.run('P.One')


def test_a_row_inserted_above_a_table_on_its_sheet_moves_it() -> None:
    """Round 57102's case: Excel moves the table down a row (tests/test_excel_pivot_shifts.py has the rest)."""
    assert _run('ws2.Rows(1).Insert') == '$C$4:$D$7 '


@pytest.mark.parametrize('edit', [
    'ws2.Rows(7).Insert', 'ws2.Rows(20).Delete', 'ws2.Columns("E").Insert', 'ws2.Columns("Z").Delete',
    'ws2.Range("A1").Insert xlShiftDown', 'ws2.Range("F1").Delete xlShiftUp', 'ws2.Range("C9").Insert xlShiftDown',
    'ws2.Range("A20").Insert xlShiftToRight', 'ws2.Range("F3").Insert xlShiftToRight', 'ws.Rows(1).Insert', 'ws.Columns(1).Delete',
])
def test_an_edit_before_beside_or_away_from_every_table_answers(edit: str) -> None:
    assert _run(edit) == '$C$3:$D$6 Row Labels'
