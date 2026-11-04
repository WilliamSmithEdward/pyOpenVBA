"""What an object answers when a macro still holds it after it has gone, replayed from live Excel (scripts/measure_held_objects.py,
tests/fixtures/held_objects.json)."""
from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/held_objects.json').read_text(encoding='utf-8'))
#: The probe's run ended instead of answering: Excel hung, or died, on one of these members of something that had gone.
ENDED = {'TIMEOUT', 'RUNNER-ERROR', 'VBA-ERROR', 'MODAL-BLOCKED'}
#: Unbuilt: ListColumns.Add and Delete; a table's AutoFilter and Sort, and a sheet's Protection, gone; Cut and Copy of a shape, which do what the fixture records.
UNBUILT = ('t.ListColumns(1).Delete', 't.ListColumns.Add', 'ws.Columns(1).Delete', 'Cut', 'Copy')
#: A chart object made and held directly answers 424 with its sheet gone, one with no data reached through ChartObjects(1) the chart's errors: no rule splits them yet.
UNBUILT_HELD = ('t.ChartObjects(1)', 't.Shapes(1).Fill', 't.ChartObjects', 't.AutoFilter', 't.Sort', 't.Protection', 't.Tab', 't.Chart.ChartArea', 't.Chart.PlotArea', 't.Chart.Axes(1)', 't.Chart.ChartArea.Format.Fill', 't.Chart.ChartTitle', 't.Chart.Legend', 't.Chart.SeriesCollection', 't.Chart.SeriesCollection(1)', 't.Chart.SeriesCollection(1).Points(1)', 't.Chart.SeriesCollection(1).Format', 't.ShapeRange')
#: The model has no notion of a variable's type, and follows what a typed variable answers (see test_excel_deleted_sheet.py): these are of a Worksheet in an Object.
TYPED = ('t',)
#: Left out of a case, each for its reason: what a chart held after its shape has gone goes on answering; resizing a table the model refuses; naming a pivot table, not built;
#: Range.Range with no argument, 450 late bound and 449 typed, where the model follows the typed (see test_excel_deleted_sheet.py).
LEFT_OUT = {'ChartObject': ('h.ChartType', 'h.HasTitle', 'h.ChartType = xlLine'), 'ListObject': ('h.Resize ws.Range("A1:B4")', 'h.Range.Address'), 'PivotTable': ('h.Name = "Z"',), 'Hyperlink': ('h.Range.Address',)}
CASES = [case for case in RECORD['cases'] if case[5] not in ENDED and case[4][0] not in LEFT_OUT.get(case[0], ('',)) and not any(one in case[3] for one in UNBUILT) and case[2] not in UNBUILT_HELD and not (case[2] == 't.ListRows(1)' and case[3] in ('t.ListRows(1).Delete', 'ws.Rows(2).Delete')) and not (case[0] == 'Worksheet' and case[2] in TYPED)
         and not (case[0] == 'FormatCondition' and case[3] == 'ws.Delete' and case[4] == ['h.NoSuchMember'])]
IDS = [f'{index}: {case[0]} | {case[2]} | {case[3]} | {case[4][0] if len(case[4]) == 1 else "*"}' for index, case in enumerate(CASES)]
_STATEMENTS = ('h.Delete', 'h.Unlist', 'h.Resize', 'h.Text "', 'h.Activate', 'h.Select', 'h.Modify', 'h.SetFirst', 'h.Refresh', 'h.BringTo')


def kept(case: list[Any]) -> list[str]:
    return [one for one in case[4] if one not in LEFT_OUT.get(case[0], ())]


def source(case: list[Any]) -> str:
    _, made, held, gone = case[:4]
    tries = kept(case)
    lines = ['Public Function Go() As String', '    Dim wb As Object, ws As Object, t As Object, h As Object, s2 As Variant, t2 As Variant, v As Variant, s As String', '    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1): ws.Name = "Sheet1"',
             '    wb.Worksheets.Add(After:=ws).Name = "Sheet2"', '    ws.Activate', '    On Error Resume Next', f'    {made}'.rstrip(': '), f'    Set h = {held}', '    s = "made " & Err.Number: Err.Clear', f'    {gone}', '    s = s & "~gone " & Err.Number: Err.Clear',
             '    v = "": v = TypeName(h): s = s & "~" & v & "/" & Err.Number: Err.Clear', '    v = "": v = (h Is Nothing): s = s & "~" & v & "/" & Err.Number: Err.Clear']
    for one in tries:
        lines += [f'    {one}: s = s & "~" & Err.Number: Err.Clear' if ' = ' in one or one == 'h.Add' or one.startswith(_STATEMENTS) else f'    v = Empty: v = {one}: s = s & "~" & Err.Number: Err.Clear']
    return '\n'.join([*lines, '    Go = s', 'End Function']) + '\n'


def answered(case: list[Any]) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case), name='Held')
    return str(app.run('Held.Go'))


def labelled(case: list[Any], answer: str) -> dict[str, str]:
    return dict(zip(['made', 'gone', 'type', 'nothing', *kept(case)], answer.split('~'), strict=True))


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any]) -> None:
    parts = case[5].split('~')
    theirs = '~'.join([*parts[:4], *(parts[4 + at] for at, one in enumerate(case[4]) if one in kept(case))])
    assert labelled(case, answered(case)) == labelled(case, theirs)


def _saved_after(statement: str, tmp_path: Path) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    Dim ws As Object, t As Object\n    Set ws = ActiveSheet\n    {PIVOT}\n    {statement}\nEnd Sub\n', name='Saved')
    app.run('Saved.Go')
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        return [name for name in package.namelist() if 'pivot' in name.lower()]


PIVOT = next(case[1] for case in RECORD['cases'] if case[0] == 'PivotTable').rstrip(': ')


def test_a_pivot_table_cleared_away_is_not_saved(tmp_path: Path) -> None:
    assert _saved_after('t.ClearTable', tmp_path)
    assert not [one for one in _saved_after('t.TableRange2.Clear', tmp_path) if 'pivotTables/' in one]
