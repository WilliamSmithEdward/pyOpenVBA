"""Workbooks.Close: the order it closes in, what it does with changes, and what stops it.

The probe's workbook is made an add-in, which the Workbooks collection does not hold, so the macro outlives the close and can
report. Rival rules for the order: first to last in the collection, or front to back by window, which `order` splits by bringing
the middle book to the front. The last case leaves the macro's workbook in the collection, and a file says how far it ran.
A save prompt is cancelled by the worker of scripts/measure_dialogs.py, which only touches its own Excel's windows.

    python scripts/measure_workbooks_close.py

writes tests/fixtures/workbooks_close.json, which tests/test_excel_workbooks_close.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

HERE = Path(__file__).resolve()
OUT = HERE.parents[1] / 'tests/fixtures/workbooks_close.json'

SINK = '''Public WithEvents App As Application
Public Log As String
Public Refuse As String

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Log = Log & "|" & Wb.Name & " " & Wb.Saved
    If Wb.Name = Refuse Then Cancel = True
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    Log = Log & "|left " & Wb.Name
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Log = Log & "|came " & Wb.Name
End Sub
'''

CASES = {
    'plain': '',
    'order': 'b.Activate',
    'changed quiet': 'b.Worksheets(1).Range("A1").Value = 2: Application.DisplayAlerts = False',
    'changed prompt': 'b.Worksheets(1).Range("A1").Value = 2: Application.DisplayAlerts = True',
    'new quiet': 'Workbooks.Add xlWBATWorksheet: Application.DisplayAlerts = False',
    'cancel': 'sink.Refuse = b.Name',
    'hidden': 'b.Windows(1).Visible = False',
    'events off': 'Application.EnableEvents = False',
    'argument': 'mode = "argument"',
    'value': 'mode = "value"',
}

MODULE = r'''Option Explicit

Public Function Tried(ByVal setup As Long, ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, c As Workbook, one As Workbook, sink As CloseSink, s As String, mode As String, r As Variant, held As Object
    ThisWorkbook.IsAddin = True
    Application.DisplayAlerts = False
    Application.EnableEvents = True
    Set a = Made(folder, "A")
    Set b = Made(folder, "B")
    Set c = Made(folder, "C")
    Set sink = New CloseSink
    Set sink.App = Application
    Set held = Workbooks
    Prepare setup, a, b, c, sink, mode
    On Error GoTo Failed
    Select Case mode
    Case "argument": held.Close 1
    Case "value": r = TypeName(held.Close)
    Case Else: Workbooks.Close
    End Select
    s = "ok " & r
Reported:
    On Error Resume Next
    s = s & sink.Log & "|count " & Workbooks.Count
    For Each one In Workbooks
        s = s & "|open " & one.Name
    Next
    Err.Clear
    s = s & "|active " & (ActiveWorkbook Is Nothing) & " e" & Err.Number
    Err.Clear
    r = ActiveSheet.Name
    s = s & "|sheet e" & Err.Number & " " & Err.Description
    Err.Clear
    r = Range("A1").Value
    s = s & "|range e" & Err.Number & " " & Err.Description
    Err.Clear
    Workbooks.Close
    s = s & "|again e" & Err.Number
    Set sink.App = Nothing
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Workbooks.Close
    Set b = Workbooks.Open(folder & "\B.xlsx", 0)
    s = s & "|B holds " & b.Worksheets(1).Range("A1").Value
    b.Close False
    ThisWorkbook.IsAddin = False
    Application.EnableEvents = True
    Tried = s
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Reported
End Function

Private Function Made(ByVal folder As String, ByVal name As String) As Workbook
    Set Made = Workbooks.Add(xlWBATWorksheet)
    Made.Worksheets(1).Range("A1").Value = 1
    Made.SaveAs folder & "\" & name & ".xlsx", xlOpenXMLWorkbook
End Function

Private Sub Prepare(ByVal setup As Long, a As Workbook, b As Workbook, c As Workbook, sink As CloseSink, mode As String)
    Select Case setup
@@PREPARE@@
    End Select
End Sub
'''

OWN = r'''Option Explicit

Public Sub Tried(ByVal folder As String)
    Dim other As Workbook, file As Long
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    Set other = Workbooks.Add(xlWBATWorksheet)
    other.SaveAs folder & "\Other.xlsx", xlOpenXMLWorkbook
    Note folder, "before " & Workbooks.Count
    Workbooks.Close
    Note folder, "after " & Workbooks.Count
End Sub

Private Sub Note(ByVal folder As String, ByVal text As String)
    Dim file As Long
    file = FreeFile
    Open folder & "\log.txt" For Append As #file
    Print #file, text
    Close #file
End Sub
'''


def main() -> None:
    config = HarnessConfig(exclusive=False, worker_argv=[sys.executable, str(HERE.with_name('measure_dialogs.py')), '--worker'])
    lines = ''.join(f'    Case {index}\n        {statement}\n' for index, statement in enumerate(CASES.values()))
    source = MODULE.replace('@@PREPARE@@', lines.rstrip('\n'))
    answers: dict[str, str] = {}
    for index, name in enumerate(CASES):
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(_with_log(config, folder)) as excel:
            excel.new_document()
            excel.run_vba('Public Function Nothing1() As Long\nEnd Function\n', 'Nothing1', (), timeout=20.0)
            excel.add_module('CloseSink', SINK, kind='class')
            result = excel.run_vba(source, 'Tried', (index, folder), timeout=30.0)
            answers[name] = (str(result.value) if result.ok else f'breaks: {result.outcome} {result.error}').replace(folder, '<DIR>')
            print(name, '=>', answers[name], flush=True)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(_with_log(config, folder)) as excel:
        excel.new_document()
        result = excel.run_vba(OWN, 'Tried', (folder,), timeout=30.0)
        answers['own workbook'] = f'{result.outcome}: ' + (Path(folder) / 'log.txt').read_text(encoding='ascii').replace('\r\n', ';').replace('\n', ';')
        print('own workbook =>', answers['own workbook'], flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': CASES, 'sink': SINK, 'source': source, 'answers': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


def _with_log(config: HarnessConfig, folder: str) -> HarnessConfig:
    from dataclasses import replace

    return replace(config, extra_env={'DIALOG_LOG': str(Path(folder) / 'driver.log')})


if __name__ == '__main__':
    main()
