"""Measure Workbook_Open and activation across two native VBA projects."""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/workbook_open.json"
LOG = '''Public EventLog As String
Public Sub Record(ByVal text As String)
EventLog = EventLog & text & "|"
End Sub
'''
SINK = '''Public WithEvents Source As Application
Private Sub Source_WorkbookOpen(ByVal Wb As Workbook)
If Wb.Name = "target.xlsm" Then Record "app-open:" & Wb.Name & ":" & CStr(Wb.Worksheets(1).Range("A1").Value)
End Sub
Private Sub Source_WorkbookActivate(ByVal Wb As Workbook)
If Wb.Name = "target.xlsm" Then Record "app-book-activate:" & Wb.Name
End Sub
Private Sub Source_WorkbookDeactivate(ByVal Wb As Workbook)
If Wb.Name = "target.xlsm" Then Record "app-book-deactivate:" & Wb.Name
End Sub
Private Sub Source_SheetActivate(ByVal Sh As Object)
If Sh.Parent.Name = "target.xlsm" Then Record "app-sheet-activate:" & Sh.Name
End Sub
Private Sub Source_SheetChange(ByVal Sh As Object, ByVal Target As Range)
If Sh.Parent.Name = "target.xlsm" Then Record "app-change:" & Sh.Name
End Sub
'''
BOOK = '''Private Sub Workbook_Open()
Application.Run "'caller.xlsm'!Record", "book-open:" & ThisWorkbook.Name & ":" & ActiveWorkbook.Name
Worksheets(1).Range("A1").Value = 41
End Sub
Private Sub Workbook_Activate()
Application.Run "'caller.xlsm'!Record", "book-activate:" & ThisWorkbook.Name
End Sub
Private Sub Workbook_Deactivate()
Application.Run "'caller.xlsm'!Record", "book-deactivate:" & ThisWorkbook.Name
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Application.Run "'caller.xlsm'!Record", "book-change:" & Sh.Name
End Sub
'''
SHEET = '''Private Sub Worksheet_Activate()
Application.Run "'caller.xlsm'!Record", "sheet-activate:" & Me.Name
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
Application.Run "'caller.xlsm'!Record", "sheet-change:" & Me.Name
End Sub
'''
ACTIONS = {
    "open_enabled": 'Application.EnableEvents = True\nSet target = Workbooks.Open("{TARGET}")',
    "open_disabled": 'Application.EnableEvents = False\nSet target = Workbooks.Open("{TARGET}")',
    "open_twice": 'Application.EnableEvents = True\nSet target = Workbooks.Open("{TARGET}")\nEventLog = ""\nSet target = Workbooks.Open("{TARGET}")',
    "switch_books": 'Application.EnableEvents = True\nSet target = Workbooks.Open("{TARGET}")\nEventLog = ""\nThisWorkbook.Activate\ntarget.Activate',
}


def add_source(lines: list[str], code: str) -> None:
    lines.append('text = ""')
    lines.extend('text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines())
    lines.append('part.CodeModule.AddFromString text')


def main() -> None:
    rows: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="pyopenvba-open-") as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
            for name, action in ACTIONS.items():
                excel.new_document()
                excel.add_module('EventLogModule', LOG)
                lines = ['Function Install() As String', 'Dim part As Object, text As String, target As Workbook',
                         'Application.EnableEvents = False', 'Application.DisplayAlerts = False',
                         f'ThisWorkbook.SaveAs "{Path(folder) / "caller.xlsm"}", 52',
                         'Set part = ThisWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "OpenSink"']
                add_source(lines, SINK)
                lines += ['Set target = Workbooks.Add', 'Set part = target.VBProject.VBComponents("ThisWorkbook")']
                add_source(lines, BOOK)
                lines += ['Set part = target.VBProject.VBComponents("Sheet1")']
                add_source(lines, SHEET)
                lines += [f'target.SaveAs "{Path(folder) / "target.xlsm"}", 52', 'target.Close False',
                          'Install = Application.Version & "|" & Application.Build', 'End Function']
                installed = excel.run_vba('\n'.join(lines), 'Install', timeout=30.0)
                assert installed.ok, installed
                source = ('Function Probe() As String\nDim sink As New OpenSink, target As Workbook\n'
                          'Set sink.Source = Application\nEventLog = ""\n' + action
                          + '\nProbe = EventLog & "#value:" & CStr(target.Worksheets(1).Range("A1").Value)\n'
                          'Application.EnableEvents = False\ntarget.Close False\nEnd Function')
                result = excel.run_vba(source.replace('{TARGET}', str(Path(folder) / "target.xlsm")), 'Probe', timeout=30.0)
                assert result.ok, (name, result)
                rows.append({"name": name, "source": source, "result": result.value})
                print(name, result.value, flush=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "excel_version_build": installed.value,
                               "log": LOG, "sink": SINK, "book": BOOK, "sheet": SHEET, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
