"""What Worksheets.Add calls a new sheet, replayed from live Excel (scripts/measure_new_sheet_names.py,
tests/fixtures/new_sheet_names.json)."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/new_sheet_names.json').read_text(encoding='utf-8'))
#: The probe's own wording of each step, kept out of a script that imports the harness.
STEPS: dict[str, str] = {
    'add': 'wb.Worksheets.Add: r = r & " " & ActiveSheet.Name',
    'OPEN': 'wb.SaveAs folder & "{sep}o.xlsx", xlOpenXMLWorkbook: wb.Close False: Set wb = Workbooks.Open(folder & "{sep}o.xlsx")',
    'Sheets.Add': 'wb.Sheets.Add: r = r & " " & ActiveSheet.Name',
    'Add Count 2': 'wb.Worksheets.Add Count:=2: r = r & " " & wb.Worksheets(1).Name & " " & wb.Worksheets(2).Name',
    'Add After': 'wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): r = r & " " & ActiveSheet.Name',
    'copy': 'wb.Worksheets(1).Copy After:=wb.Worksheets(1): r = r & " " & ActiveSheet.Name',
    'chart': 'wb.Charts.Add: r = r & " " & ActiveSheet.Name',
    'Add xlWorksheet': 'wb.Sheets.Add Type:=xlWorksheet: r = r & " " & ActiveSheet.Name',
    'other': 'Set other = Workbooks.Add(xlWBATWorksheet): other.Worksheets.Add: r = r & " " & ActiveSheet.Name: wb.Activate',
    'other add': 'Set other = Workbooks.Add(xlWBATWorksheet): other.Worksheets.Add: wb.Activate: r = r & " " & ActiveSheet.Name',
}


def _statement(step: str) -> str:
    words = step.split()
    if step in STEPS:
        return STEPS[step].replace('{sep}', os.sep)
    if words[0] == 'name':
        return f'wb.Worksheets(1).Name = "{words[1]}"'
    if words[0] == 'rename':
        return f'wb.Worksheets("{words[1]}").Name = "{words[2]}"'
    assert words[0] == 'drop', step
    return f'Application.DisplayAlerts = False: wb.Worksheets("{words[1]}").Delete'


#: A new chart sheet is not built, and reports itself.
REFUSED = [case for case in RECORD['cases'] if 'chart' in case[0]]
KEPT = [case for case in RECORD['cases'] if 'chart' not in case[0]]


@pytest.mark.parametrize(('steps', 'answer'), KEPT, ids=[' / '.join(case[0]) for case in KEPT])
def test_each_sheet_as_excel_named_it(tmp_path: Path, steps: list[str], answer: str) -> None:
    assert _named(tmp_path, steps) == answer


@pytest.mark.parametrize(('steps', 'answer'), REFUSED, ids=[' / '.join(case[0]) for case in REFUSED])
def test_a_chart_sheet_says_so(tmp_path: Path, steps: list[str], answer: str) -> None:
    assert answer.startswith('ok Chart1')
    with pytest.raises(VBAUnsupportedError):
        _named(tmp_path, steps)


def _named(folder: Path, steps: list[str]) -> object:
    body = '\n'.join(f'    {_statement(step)}' for step in steps)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Function Tried(ByVal folder As String) As String
    Dim wb As Workbook, other As Workbook, r As Variant, s As Object
    Application.DisplayAlerts = False
    r = ""
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo Failed
{body}
    For Each s In wb.Sheets
        r = r & " [" & s.Name & "]"
    Next
    Tried = "ok" & r
    GoTo Finish
Failed:
    Tried = "Err " & Err.Number & r
    Resume Finish
Finish:
End Function
''', name='Names')
    return app.run('Names.Tried', str(folder))
