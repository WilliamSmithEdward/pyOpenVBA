"""What Workbooks.Open and Workbooks.Add make of a template: a new book named after it, with no path, or the file itself.

    python scripts/measure_template_opens.py

writes tests/fixtures/template_opens.json, which tests/test_excel_template_saves.py replays. The templates hold A1 = 7 and a
named cell, in a temporary folder spelled <folder> in what is recorded.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/template_opens.json'
SAY = 'r = wb.Name & "|" & wb.Path & "|" & wb.FullName & "|" & wb.FileFormat & "|" & wb.Saved & "|" & wb.Worksheets(1).Range("A1").Value & "|" & wb.Names.Count & "|" & wb.ReadOnly & "|" & wb.HasVBProject'
OPEN_X = 'Set wb = Workbooks.Open(folder & "\\t.xltx"): '
OPEN_M = 'Set wb = Workbooks.Open(folder & "\\t.xltm"): '
ADD_X = 'Set wb = Workbooks.Add(folder & "\\t.xltx"): '
CASES = (f'{OPEN_X}{SAY}', f'{OPEN_M}{SAY}', f'{ADD_X}{SAY}', f'Set wb = Workbooks.Add(folder & "\\t.xltm"): {SAY}', f'Set wb = Workbooks.Open(folder & "\\u.xlsx"): {SAY}', f'Set wb = Workbooks.Add(folder & "\\u.xlsx"): {SAY}',
         # How the number goes.
         f'{OPEN_X}Set other = Workbooks.Open(folder & "\\t.xltx"): r = wb.Name & " " & other.Name: other.Close False',
         f'{OPEN_X}Set other = Workbooks.Open(folder & "\\u2.xltx"): r = wb.Name & " " & other.Name: other.Close False',
         f'{OPEN_X}n = wb.Name: wb.Close False: Set wb = Workbooks.Open(folder & "\\t.xltx"): r = n & " " & wb.Name',
         f'{ADD_X}Set other = Workbooks.Open(folder & "\\t.xltx"): r = wb.Name & " " & other.Name: other.Close False',
         f'Set wb = Workbooks.Add: Set other = Workbooks.Open(folder & "\\t.xltx"): r = (Left(wb.Name, 4) = "Book") & " " & other.Name: other.Close False',
         # Saving what came of it: to Save it needs a name; the template is untouched.
         f'{OPEN_X}wb.Worksheets(1).Range("A2").Value = 1: wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook: r = wb.Name & " " & wb.FileFormat & " " & wb.Path',
         f'{OPEN_X}wb.Worksheets(1).Range("A2").Value = 1: wb.SaveAs folder & "\\saved2.xltx", xlOpenXMLTemplate: r = wb.Name & " " & wb.FileFormat',
         f'{OPEN_X}r = Dir(folder & "\\t.xltx") & " " & wb.Saved: wb.Worksheets(1).Range("A2").Value = 1: r = r & " " & wb.Saved',
         f'{OPEN_X}wb.SaveCopyAs folder & "\\copy.xlsx": r = wb.Name & " " & Dir(folder & "\\copy.xlsx")',
         # By name, and arguments.
         f'{OPEN_X}r = Workbooks(wb.Name).Name', f'{OPEN_X}r = TypeName(Workbooks("t.xltx"))', 'Set wb = Workbooks.Open(folder & "\\t.xltx", ReadOnly:=True): ' + SAY, 'Set wb = Workbooks.Open(folder & "\\t.xltx", Editable:=True): ' + SAY,
         'Set wb = Workbooks.Open(folder & "\\t.xltm", Editable:=True): ' + SAY,
         # A template that had been saved by the workbook still open: SaveAs a template keeps the file itself open as a template.
         'Set wb = Workbooks.Add: wb.SaveAs folder & "\\live.xltx", xlOpenXMLTemplate: r = wb.Name & " " & wb.FileFormat & " " & wb.Path: wb.Save: r = r & " " & wb.Name & " " & wb.FileFormat')

PROBE = '''Public Function Tried(ByVal folder As String, ByVal which As Long) As String
    Dim wb As Workbook, other As Workbook, r As Variant, n As Variant
    Application.DisplayAlerts = False
    On Error GoTo Failed
    Select Case which
{cases}
    End Select
    Tried = "ok " & r
    GoTo Finish
Failed:
    Tried = "Err " & Err.Number
    Resume Finish
Finish:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.DisplayAlerts = True
End Function

Public Sub Prepare(ByVal folder As String)
    Dim wb As Workbook
    Application.DisplayAlerts = False
    For Each n In Array(Array("t.xltx", xlOpenXMLTemplate), Array("u2.xltx", xlOpenXMLTemplate), Array("t.xltm", xlOpenXMLTemplateMacroEnabled), Array("u.xlsx", xlOpenXMLWorkbook))
        Set wb = Workbooks.Add
        wb.Worksheets(1).Range("A1").Value = 7
        wb.Names.Add "Seven", "=" & wb.Worksheets(1).Name & "!$A$1"
        wb.SaveAs folder & "\\" & n(0), n(1)
        wb.Close False
    Next
    Application.DisplayAlerts = True
End Sub
'''


def main() -> None:
    cases = '\n'.join(f'    Case {index}\n        {statement}' for index, statement in enumerate(CASES))
    source = PROBE.format(cases=cases).replace('Dim wb As Workbook\n    Application.DisplayAlerts = False\n    For Each', 'Dim wb As Workbook, n As Variant\n    Application.DisplayAlerts = False\n    For Each')
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        # Each case in a session of its own: Excel's numbering goes on through a session.
        for index in range(len(CASES)):
            for name in Path(folder).iterdir():
                name.unlink()
            with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
                excel.new_document()
                excel.run_vba(source, 'Prepare', (folder,), timeout=120.0, line_numbers=False)
                done = excel.run_vba(source, 'Tried', (folder, index), timeout=120.0, line_numbers=False)
                answers.append(re.sub(re.escape(folder), '<folder>', str(done.value) if done.ok else f'STOPPED {done.outcome}', flags=re.IGNORECASE))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': [[statement, answer] for statement, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer in record['cases']:
        print(statement[:90], '=>', answer)


if __name__ == '__main__':
    main()
