"""A Worksheet variable whose sheet was deleted (issue #246): what each member answers or raises, whether a new sheet
of the same name takes its place, and a sheet's own names after the sheet is renamed.

Predictions written before measuring: every member of the deleted sheet's variable raises the same automation error,
-2147221080, "Method 'X' of object '_Worksheet' failed"; TypeName still says Worksheet; Is Nothing is False; a new
sheet of the old name is a different object; a sheet's own name stays with the sheet through a rename.

    python scripts/measure_deleted_sheet.py

writes tests/fixtures/deleted_sheet.json, which tests/test_excel_deleted_sheet.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/deleted_sheet.json'
#: Each case: the expression read through ``old``, a variable set to the sheet "Sum" before it was deleted and a new
#: "Sum" added in its place.
CASES = [
    'old.Rows.Count', 'old.Name', 'old.Index', 'TypeName(old)', 'CStr(old Is Nothing)', 'old.Range("A1").Value',
    'old.Cells(1, 1).Value', 'old.Parent.Name', 'old.Visible', 'CStr(old Is ThisWorkbook.Worksheets("Sum"))',
    'old.Name = "Other"', 'old.Activate', 'old.Delete', 'old.CodeName', 'old.UsedRange.Address',
]
#: The cell read from a Range taken from the sheet before it was deleted.
HELD = ['held.Value', 'held.Address', 'TypeName(held)', 'held.Worksheet.Name', 'page.PrintArea', 'page.PrintArea = "$A$1:$C$9"']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim old As Worksheet, sh As Worksheet, held As Range, wb As Workbook, page As PageSetup',
             '    Set wb = ThisWorkbook',
             '    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))', '    sh.Name = "Sum"',
             '    sh.Range("A1").Value = 1', '    Set old = wb.Worksheets("Sum")', '    Set held = old.Range("A1")', '    Set page = old.PageSetup',
             '    Application.DisplayAlerts = False', '    wb.Worksheets("Sum").Delete',
             '    Application.DisplayAlerts = True',
             '    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))', '    sh.Name = "Sum"',
             '    sh.Range("A1").Value = 5', '    On Error GoTo Failed', '    Select Case which']
    for index, read in enumerate(CASES + HELD):
        if '=' in read and not read.startswith('CStr') and 'Is' not in read:
            lines += [f'    Case {index}', f'        {read}', '        Tried = "done"']
        elif read in ('old.Activate', 'old.Delete'):
            lines += [f'    Case {index}', f'        {read}', '        Tried = "done"']
        else:
            lines += [f'    Case {index}', f'        Tried = {read}']
    lines += ['    End Select', 'Done:', '    Application.DisplayAlerts = False', '    sh.Delete',
              '    Application.DisplayAlerts = True', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Private Function Renamed() As String',
              '    Dim a As Worksheet, b As Worksheet, wb As Workbook, s As String', '    Set wb = ThisWorkbook',
              '    Set a = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))', '    a.Name = "X"',
              '    a.Names.Add Name:="mark", RefersTo:="=""owner""", Visible:=False', '    a.Name = "Y"',
              '    Set b = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))', '    b.Name = "X"',
              '    s = "X:" & b.Names.Count & " Y:" & a.Names.Count & " " & a.Names(1).Name',
              '    Application.DisplayAlerts = False', '    a.Delete', '    b.Delete', '    Application.DisplayAlerts = True',
              '    Renamed = s', 'End Function', '',
              'Public Function Probe() As String', '    Dim s As String, i As Long',
              f'    For i = 0 To {len(CASES) + len(HELD) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4) & "@@@" & Renamed()', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': dict(zip([*CASES, *HELD, 'renamed'], answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(case, '->', answer)


if __name__ == '__main__':
    main()
