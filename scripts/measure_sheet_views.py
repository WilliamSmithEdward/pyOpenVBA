"""Window.SheetViews and ActiveSheetView, and the WorksheetView each gives: what they read, which sheets they hold
with one sheet or several selected, and how a view's settings reach the window.

Each case makes a workbook of three sheets, runs its statements with win as the workbook's window, and reads READS,
each as its value (TypeName for an object) or "E" and the error number.

    python scripts/measure_sheet_views.py

writes tests/fixtures/sheet_views.json, which tests/test_excel_sheet_views.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sheet_views.json'
CASES = [
    '',
    'wb.Worksheets(2).Activate',
    'wb.Worksheets(Array(1, 3)).Select',
    'wb.Worksheets(Array(3, 1, 2)).Select',
    'win.ActiveSheetView.DisplayGridlines = False',
    'win.ActiveSheetView.DisplayFormulas = True: win.ActiveSheetView.DisplayZeros = False',
    'win.ActiveSheetView.DisplayHeadings = False: win.ActiveSheetView.DisplayOutline = False',
    'wb.Worksheets(Array(1, 3)).Select: win.SheetViews(2).DisplayGridlines = False: wb.Worksheets(3).Activate',
    'win.DisplayGridlines = False',
    'Set v = win.ActiveSheetView: wb.Worksheets(2).Activate: v.DisplayGridlines = False: wb.Worksheets(1).Activate',
]
READS = [
    'TypeName(win.ActiveSheetView)', 'win.ActiveSheetView.Sheet.Name', 'win.ActiveSheetView.DisplayGridlines',
    'win.ActiveSheetView.DisplayFormulas', 'win.ActiveSheetView.DisplayHeadings',
    'win.ActiveSheetView.DisplayOutline', 'win.ActiveSheetView.DisplayZeros',
    'TypeName(win.ActiveSheetView.Parent)', 'TypeName(win.ActiveSheetView.Application)',
    'TypeName(win.SheetViews)', 'win.SheetViews.Count', 'win.SheetViews(1).Sheet.Name',
    'win.SheetViews(2).Sheet.Name', 'win.SheetViews(3).Sheet.Name', 'win.SheetViews(0).Sheet.Name',
    'win.SheetViews("Sheet1").Sheet.Name', 'TypeName(win.SheetViews.Parent)', 'TypeName(win.SheetViews.Application)',
    'win.SheetViews.Item(1).Sheet.Name',
    'win.DisplayGridlines', 'wb.ActiveSheet.Name',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, win As Window, v As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    wb.Worksheets.Add After:=wb.Worksheets(1)',
             '    wb.Worksheets.Add After:=wb.Worksheets(2)', '    wb.Worksheets(1).Activate',
             '    Set win = wb.Windows(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo 0', '    Tried = "ok " & Mid(s, 2)', 'Closed:', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Closed', 'End Function',
              '', 'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[statements, answer] for statements, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, answer in record['cases']:
        print(statements, '->', answer)


if __name__ == '__main__':
    main()
