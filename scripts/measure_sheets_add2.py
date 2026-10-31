"""Sheets.Add2 and Worksheets.Add2 on a workbook's sheets, with each argument and none, and how many sheets are left.

    python scripts/measure_sheets_add2.py

writes tests/fixtures/sheets_add2.json, which tests/test_excel_sheets_add2.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sheets_add2.json'
CALLS = ['wb.Sheets.Add2', 'wb.Sheets.Add2(After:=wb.Sheets(1))', 'wb.Sheets.Add2(Count:=2)',
         'wb.Sheets.Add2(Before:=wb.Sheets(1), Count:=2)', 'wb.Sheets.Add2(NewLayout:=True)',
         'wb.Sheets.Add2(Count:=0)', 'wb.Worksheets.Add2', 'wb.Worksheets.Add2(After:=wb.Sheets(1))']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String', '    Dim wb As Object, o As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next', '    Select Case which']
    for index, call in enumerate(CALLS):
        lines += [f'    Case {index}', f'        Set o = {call}']
    lines += ['    End Select', '    Tried = Err.Number & " " & wb.Sheets.Count', '    On Error GoTo 0',
              '    wb.Close False', 'End Function', '', 'Public Function Probe() As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CALLS) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'calls': [[call, answer] for call, answer in zip(CALLS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for call, answer in record['calls']:
        print(call, '->', answer)


if __name__ == '__main__':
    main()
