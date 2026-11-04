"""What Range.Sort does to the references of a formula it moves, replayed from live Excel
(scripts/measure_sort_formulas.py, tests/fixtures/sort_formulas.json). A sheet's formulas move with its rows, and a formula on
another sheet that follows a Cut reaches the saved file."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/sort_formulas.json').read_text(encoding='utf-8'))
SORTS = {'rows': ('ws.Cells(i, 1).Value = i', 'ws.Range("A1:C5").Sort Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo'),
         'columns': ('ws.Cells(1, i).Value = i', 'ws.Range("A1:E3").Sort Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo, Orientation:=xlSortRows')}


@pytest.mark.parametrize('case', RECORD['cases'], ids=[f'{case[0]} {case[1]} {case[2]}' for case in RECORD['cases']])
def test_each_formula_as_excel_moved_it(case: list[str]) -> None:
    kind, start, formula, landing, answer = case
    fill, sort = SORTS[kind]
    app = ExcelApplication()
    app.add_workbook()
    doubled = formula.replace('"', '""')
    app.add_module(f'''Public Function Go() As String
    Dim wb As Workbook, ws As Worksheet, i As Long
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    wb.Worksheets.Add(After:=ws).Name = "Two"
    wb.Names.Add "Nm", "=One!$A$1"
    For i = 1 To 5
        wb.Worksheets("Two").Cells(i, 1).Value = i * 100
        {fill}
    Next
    ws.Range("{start}").Formula = "{doubled}"
    {sort}
    Go = ws.Range("{landing}").Formula & "|" & ws.Range("{landing}").Text
End Function
''', name='Sorted')
    assert app.run('Sorted.Go') == answer


def test_a_formula_on_another_sheet_that_follows_a_cut_is_saved(tmp_path: Path) -> None:
    """In a file opened, where the other sheet's part would otherwise go back as it came."""
    first = ExcelApplication()
    first.add_workbook()
    first.add_module('''Public Sub Make()
    ActiveWorkbook.Worksheets(1).Name = "One"
    ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(1)).Name = "Two"
    ActiveWorkbook.Worksheets("One").Range("C3").Value = 5
    ActiveWorkbook.Worksheets("Two").Range("A1").Formula = "=One!C3"
End Sub
''', name='Make')
    first.run('Make.Make')
    source = tmp_path / 'source.xlsx'
    first.save(source)
    app = ExcelApplication.open(source, with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets("One").Range("C3").Cut ActiveWorkbook.Worksheets("One").Range("H8")\nEnd Sub\n', name='Moved')
    app.run('Moved.Go')
    out = tmp_path / 'out.xlsx'
    app.save(out)
    again = ExcelApplication.open(out, with_vba=False)
    again.add_module('Public Function Say() As String\n    Say = ActiveWorkbook.Worksheets("Two").Range("A1").Formula\nEnd Function\n', name='Read')
    assert again.run('Read.Say') == '=One!H8'
