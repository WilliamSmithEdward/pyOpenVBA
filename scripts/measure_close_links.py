"""When closing a workbook, at what point the links other workbooks have to it take its cells for the last time.

Dest.xlsx reads A1 of Alpha.xlsx, which holds 1. Both are open, Alpha active. An application sink logs each event of the
close, and in the one the case names edits A1 of Alpha to 2. What Dest reads afterward says whether the link read Alpha
after that event. The rival rules this splits: the link reads at BeforeClose, so only an edit before it counts; or after
the window events, so an edit in BeforeClose, WindowDeactivate and WorkbookDeactivate all count, and one in the
activation of the window behind, if Alpha is still there to edit, does not.

    python scripts/measure_close_links.py

writes tests/fixtures/close_links.json, which tests/test_excel_close_links.py replays.
"""
from __future__ import annotations

import base64
import io
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/close_links.json'
MODES = ['None', 'BeforeClose', 'WindowDeactivate', 'WorkbookDeactivate', 'WindowActivate', 'WorkbookActivate']

SINK = '''Public WithEvents App As Application
Public Mode As String
Public Log As String

Private Sub Happened(ByVal what As String, ByVal book As Object)
    Log = Log & "|" & what & " " & book.Name
    If what = Mode Then
        On Error Resume Next
        Workbooks("Alpha.xlsx").Worksheets(1).Range("A1").Value = 2
        If Err.Number <> 0 Then Log = Log & " edit failed " & Err.Number
        On Error GoTo 0
    End If
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Happened "BeforeClose", Wb
End Sub

Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Happened "WindowDeactivate", Wb
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook)
    Happened "WorkbookDeactivate", Wb
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Happened "WindowActivate", Wb
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Happened "WorkbookActivate", Wb
End Sub
'''

MODULE = r'''Option Explicit

Public Function Made(ByVal folder As String) As String
    Dim alpha As Workbook, dest As Workbook
    Application.DisplayAlerts = False
    Set alpha = Workbooks.Add(xlWBATWorksheet)
    alpha.Worksheets(1).Range("A1").Value = 1
    alpha.SaveAs folder & "\Alpha.xlsx", xlOpenXMLWorkbook
    Set dest = Workbooks.Add(xlWBATWorksheet)
    dest.Worksheets(1).Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"
    dest.SaveAs folder & "\Dest.xlsx", xlOpenXMLWorkbook
    dest.Close False
    alpha.Close False
    Made = "done"
End Function

Public Function Tried(ByVal mode As String, ByVal folder As String) As String
    Dim alpha As Workbook, dest As Workbook, sink As CloseSink, wasOn As Boolean, s As String, v As Variant
    Application.DisplayAlerts = False
    Set dest = Workbooks.Open(folder & "\Dest.xlsx", 0)
    Set alpha = Workbooks.Open(folder & "\Alpha.xlsx", 0)
    wasOn = Application.EnableEvents
    Application.EnableEvents = True
    Set sink = New CloseSink
    Set sink.App = Application
    sink.Mode = mode
    alpha.Close False
    Application.EnableEvents = wasOn
    Set sink.App = Nothing
    v = dest.Worksheets(1).Range("A1").Value
    s = TypeName(v) & ":" & CStr(v) & sink.Log
    dest.Close False
    Tried = s
End Function
'''


def _scrubbed(data: bytes) -> str:
    """A saved file with no machine's folder in it."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as saved, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as kept:
        for item in saved.infolist():
            body = saved.read(item.filename)
            if item.filename.endswith(('.xml', '.rels')):
                body = re.sub(rb'C:\\Users\\[^"<>]*?\\tmp[a-z0-9_]+', rb'C:\\folder', body)
            kept.writestr(item.filename, body)
    return base64.b64encode(out.getvalue()).decode('ascii')


def main() -> None:
    source = MODULE
    answers: dict[str, str] = {}
    files: dict[str, str] = {}
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        excel.run_vba('Public Function Nothing1() As Long\nEnd Function\n', 'Nothing1', (), timeout=20.0)
        excel.add_module('CloseSink', SINK, kind='class')
        for mode in MODES:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                done = excel.run_vba(source, 'Made', (folder,), timeout=30.0)
                assert done.ok, done
                files = {name: _scrubbed((Path(folder) / name).read_bytes()) for name in ('Alpha.xlsx', 'Dest.xlsx')}
                result = excel.run_vba(source, 'Tried', (mode, folder), timeout=30.0)
            answers[mode] = str(result.value) if result.ok and not result.dialogs else f'breaks: {result.outcome} {result.error}'
            print(mode, '=>', answers[mode], flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'sink': SINK, 'source': source, 'files': files, 'answers': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
