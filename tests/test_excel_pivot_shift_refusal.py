"""An insert or delete that would move, cut or split a pivot table reports itself, until the model moves tables as Excel does.
PivotTable round 57102 (tests/fixtures/pivot_table_rounds/, on branch verify02-pivot-rounds) found Excel moving a table down
for a row inserted above it, where the model left it in place. An edit before every table, beside it, or on another sheet
still answers, and a refused edit changes nothing."""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

#: A table at C3 on Second: rows 3 to 6, columns C to D, from Main's keys and values.
SETUP = '''    Dim wb As Object, ws As Object, ws2 As Object, pt As Object
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Main"
    Set ws2 = wb.Worksheets.Add(After:=ws): ws2.Name = "Second"
    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws2.Range("C3"), "PT")
    pt.PivotFields("k").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum
'''
REFUSED = 'Inserting or deleting cells that would move a pivot table is not implemented'


def _run(edit: str) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function One() As String\n' + SETUP + f'    {edit}\n'
                   '    One = ws2.PivotTables(1).TableRange1.Address & " " & ws2.Range("C3").Value\nEnd Function\n', name='P')
    return app.run('P.One')


def test_a_row_inserted_above_a_table_on_its_sheet_reports_itself() -> None:
    """Round 57102's case: Excel moves the table down a row."""
    with pytest.raises(VBAUnsupportedError, match=REFUSED):
        _run('ws2.Rows(1).Insert')


@pytest.mark.parametrize('edit', [
    'ws2.Rows(1).Delete', 'ws2.Rows(4).Insert', 'ws2.Rows(6).Delete', 'ws2.Columns(1).Insert', 'ws2.Columns("D").Delete',
    'ws2.Range("C1").Insert xlShiftDown', 'ws2.Range("D2").Delete xlShiftUp', 'ws2.Range("A3").Insert xlShiftToRight',
    'ws2.Range("B5").Delete xlShiftToLeft', 'ws2.Range("A1:H1").Insert xlShiftDown',
])
def test_an_edit_that_moves_cuts_or_splits_a_table_reports_itself(edit: str) -> None:
    with pytest.raises(VBAUnsupportedError, match=REFUSED):
        _run(edit)


@pytest.mark.parametrize('edit', [
    'ws2.Rows(7).Insert', 'ws2.Rows(20).Delete', 'ws2.Columns("E").Insert', 'ws2.Columns("Z").Delete',
    'ws2.Range("A1").Insert xlShiftDown', 'ws2.Range("F1").Delete xlShiftUp', 'ws2.Range("C9").Insert xlShiftDown',
    'ws2.Range("A20").Insert xlShiftToRight', 'ws2.Range("F3").Insert xlShiftToRight', 'ws.Rows(1).Insert', 'ws.Columns(1).Delete',
])
def test_an_edit_before_beside_or_away_from_every_table_answers(edit: str) -> None:
    assert _run(edit) == '$C$3:$D$6 Row Labels'


def test_a_refused_edit_changes_nothing() -> None:
    """The report is raised before anything moves: the cell above the table and the table stay where they were."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Sub Edit()\n' + SETUP + '    ws2.Range("A1").Value = 7\n    ws2.Rows(1).Insert\nEnd Sub\n'
                   'Function Read() As String\n    Dim ws2 As Object: Set ws2 = ActiveWorkbook.Worksheets("Second")\n'
                   '    Read = ws2.Range("A1").Value & " " & ws2.PivotTables(1).TableRange1.Address\nEnd Function\n', name='P')
    with pytest.raises(VBAUnsupportedError, match=REFUSED):
        app.run('P.Edit')
    assert app.run('P.Read') == '7 $C$3:$D$6'
