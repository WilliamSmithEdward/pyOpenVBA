"""How merged areas follow whole rows and columns, and cells, inserted and deleted: every merged area afterwards, the
errors, and where the merged value went.

Each case makes a workbook, merges B2:C3 with "m" in it and E5:E6 with "n", runs the case's statement, then reads
every merged area in A1:H10, each with its top left cell's value.

    python scripts/measure_merge_structure_edits.py

writes tests/fixtures/merge_structure_edits.json, which tests/test_excel_merge_structure_edits.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/merge_structure_edits.json'
CASES = [
    '',
    'ws.Rows(1).Insert',
    'ws.Rows(3).Insert',
    'ws.Rows(2).Insert',
    'ws.Rows(4).Insert',
    'ws.Rows("3:4").Insert',
    'ws.Rows(1).Delete',
    'ws.Rows(2).Delete',
    'ws.Rows(3).Delete',
    'ws.Rows("2:3").Delete',
    'ws.Rows("1:2").Delete',
    'ws.Rows("3:5").Delete',
    'ws.Columns(1).Insert',
    'ws.Columns(3).Insert',
    'ws.Columns(2).Delete',
    'ws.Columns("B:C").Delete',
    'ws.Columns("C:E").Delete',
    'ws.Range("A1").Insert xlShiftDown',
    'ws.Range("B1").Insert xlShiftDown',
    'ws.Range("B1:C1").Insert xlShiftDown',
    'ws.Range("B1:D1").Insert xlShiftDown',
    'ws.Range("A2").Insert xlShiftToRight',
    'ws.Range("A2:A3").Insert xlShiftToRight',
    'ws.Range("B4").Delete xlShiftUp',
    'ws.Range("B2:C3").Delete xlShiftUp',
    'ws.Range("B2:C2").Delete xlShiftUp',
    'ws.Range("B2").EntireRow.Insert',
    'ws.Range("B2:C3").EntireColumn.Delete',
]


def _probe() -> str:
    lines = ['Private Function Merges(ByVal ws As Worksheet) As String',
             '    Dim c As Range, s As String', '    For Each c In ws.Range("A1:H10")',
             '        If c.MergeCells Then',
             '            If c.Address = c.MergeArea.Cells(1, 1).Address Then s = s & "|" & '
             'c.MergeArea.Address(0, 0) & "=" & c.Value',
             '        End If', '    Next', '    s = s & " values"',
             '    For Each c In ws.Range("A1:F8")',
             '        If Not IsEmpty(c.Value) Then s = s & " " & c.Address(0, 0) & "=" & c.Value', '    Next',
             '    Merges = s', 'End Function', '',
             'Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("B2").Value = "m": ws.Range("B2:C3").Merge',
             '    ws.Range("E5").Value = "n": ws.Range("E5:E6").Merge',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {case}'] if case else [])
    lines += ['    End Select', '    Tried = "ok " & Merges(ws)', 'Done:', '    wb.Close False',
              '    Application.DisplayAlerts = True', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Merges(ws)', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
