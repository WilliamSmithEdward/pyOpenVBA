"""Record cancellable save/close events in isolated Excel workbooks."""

from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/workbook_lifecycle.json"
LOG = '''Public Log As String
Public CancelInDocument As Boolean
'''
SINK = '''Public WithEvents Book As Workbook
Public WithEvents App As Application
Public CancelSave As Boolean
Public CancelClose As Boolean
Public AppCancelSave As Boolean
Public AppCancelClose As Boolean
Public UndoCancel As Boolean
Public SwitchEventsOff As Boolean
Private Sub Book_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Log = Log & "book-before-save:" & CStr(SaveAsUI) & ":" & CStr(Cancel) & "|"
If CancelSave Then Cancel = True
If SwitchEventsOff Then Application.EnableEvents = False
End Sub
Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
If Not Wb Is Book Then Exit Sub
Log = Log & "app-before-save:" & CStr(SaveAsUI) & ":" & CStr(Cancel) & "|"
If AppCancelSave Then Cancel = True
If UndoCancel Then Cancel = False
End Sub
Private Sub Book_AfterSave(ByVal Success As Boolean)
Log = Log & "book-after-save:" & CStr(Success) & "|"
End Sub
Private Sub App_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
If Not Wb Is Book Then Exit Sub
Log = Log & "app-after-save:" & CStr(Success) & "|"
End Sub
Private Sub Book_BeforeClose(Cancel As Boolean)
Log = Log & "book-before-close:" & CStr(Cancel) & "|"
If CancelClose Then Cancel = True
End Sub
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
If Not Wb Is Book Then Exit Sub
Log = Log & "app-before-close:" & CStr(Cancel) & "|"
If AppCancelClose Then Cancel = True
End Sub
'''
BOOK = '''Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Log = Log & "own-before-save:" & CStr(SaveAsUI) & ":" & CStr(Cancel) & "|"
If CancelInDocument Then Cancel = True
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
Log = Log & "own-after-save:" & CStr(Success) & "|"
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Log = Log & "own-before-close:" & CStr(Cancel) & "|"
If CancelInDocument Then Cancel = True
End Sub
'''
SAVE = 'target.SaveAs Filename:="{OUTPUT}", FileFormat:=52'
ACTIONS = {
    "save_as": SAVE,
    "cancel_save_sink": 'sink.CancelSave = True\n' + SAVE,
    "cancel_save_unnamed": 'sink.CancelSave = True\ntarget.Save',
    "cancel_saveas_dialog": 'sink.CancelSave = True\ntarget.SaveAs',
    "cancel_save_document": 'CancelInDocument = True\n' + SAVE,
    "cancel_save_app": 'sink.AppCancelSave = True\n' + SAVE,
    "reverse_save_cancel": 'sink.CancelSave = True\nsink.UndoCancel = True\n' + SAVE,
    "save_events_disabled": 'Application.EnableEvents = False\n' + SAVE,
    "save_disable_events_in_sink": 'sink.SwitchEventsOff = True\n' + SAVE,
    "save_existing": 'Application.EnableEvents = False\n' + SAVE + '\nApplication.EnableEvents = True\ntarget.Sheets(1).Range("A1").Value = 2\nLog = ""\ntarget.Save',
    "save_failure": 'On Error Resume Next\ntarget.SaveAs Filename:="{OUTPUT}.missing/failed.xlsm", FileFormat:=52\nLog = Log & "error:" & CStr(Err.Number) & "|"',
    "save_default_format": 'On Error Resume Next\ntarget.SaveAs Filename:="{OUTPUT}"\nLog = Log & "error:" & CStr(Err.Number) & "|"',
    "save_mismatched_format": 'On Error Resume Next\ntarget.SaveAs Filename:="{OUTPUT}.xlsx", FileFormat:=52\nLog = Log & "error:" & CStr(Err.Number) & "|"',
    "cancel_close_sink": 'sink.CancelClose = True\ntarget.Close SaveChanges:=False',
    "cancel_close_document": 'CancelInDocument = True\ntarget.Close SaveChanges:=False',
    "cancel_close_app": 'sink.AppCancelClose = True\ntarget.Close SaveChanges:=False',
    "cancel_close_before_save": 'sink.CancelClose = True\ntarget.Close SaveChanges:=True',
    "cancel_close_invalid_filename": 'sink.CancelClose = True\nOn Error Resume Next\ntarget.Close SaveChanges:=True, Filename:="{OUTPUT}"\nLog = Log & "error:" & CStr(Err.Number) & "|"',
    "close_other_book": 'Application.EnableEvents = False\nSet target = Workbooks.Add\nSet sink.Book = target\nApplication.EnableEvents = True\ntarget.Close SaveChanges:=False',
    "close_other_save_filename": 'Application.EnableEvents = False\nSet target = Workbooks.Add\nSet sink.Book = target\ntarget.Sheets(1).Range("A1").Value = 2\nApplication.EnableEvents = True\ntarget.Close SaveChanges:=True, Filename:="{OUTPUT}.xlsx"',
    "cancel_other_close_save_filename": 'Application.EnableEvents = False\nSet target = Workbooks.Add\nSet sink.Book = target\ntarget.Sheets(1).Range("A1").Value = 2\nsink.CancelClose = True\nApplication.EnableEvents = True\ntarget.Close SaveChanges:=True, Filename:="{OUTPUT}.xlsx"',
    "close_cancel_save": 'Application.EnableEvents = False\nSet target = Workbooks.Add\nSet sink.Book = target\ntarget.SaveAs Filename:="{OUTPUT}.xlsx", FileFormat:=51\ntarget.Sheets(1).Range("A1").Value = 2\nsink.CancelSave = True\nApplication.EnableEvents = True\nOn Error Resume Next\ntarget.Close SaveChanges:=True\nLog = Log & "error:" & CStr(Err.Number) & "|"',
}


def main() -> None:
    lines = ['Function Install() As String', 'Dim text As String, part As Object']
    for owner, code in ((None, SINK), ('"ThisWorkbook"', BOOK)):
        if owner is None:
            lines += ['Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "LifecycleSink"']
        else:
            lines += [f'Set part = ActiveWorkbook.VBProject.VBComponents({owner})']
        lines += ['text = ""']
        lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
        lines += ['part.CodeModule.AddFromString text']
    lines += ['Install = Application.Version & "|" & Application.Build', 'End Function']
    rows = []
    with tempfile.TemporaryDirectory(prefix="pyopenvba-lifecycle-") as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
            for name, action in ACTIONS.items():
                excel.new_document()
                excel.add_module("EventLog", LOG)
                installed = excel.run_vba('\n'.join(lines), 'Install', timeout=30.0)
                assert installed.ok, installed
                source = '''Function Probe() As String
Dim target As Workbook, sink As New LifecycleSink
Dim startCount As Long
startCount = Workbooks.Count
Set target = ThisWorkbook
Set sink.Book = target
Set sink.App = Application
Application.EnableEvents = True
Log = ""
''' + action + '\nProbe = Log & "#count:" & CStr(Workbooks.Count - startCount) & "#format:" & CStr(ThisWorkbook.FileFormat) & ":" & TypeName(ThisWorkbook.FileFormat)\nApplication.EnableEvents = False\nEnd Function\n'
                code = source.replace('{OUTPUT}', str(Path(folder) / f'{name}.xlsm'))
                result = excel.run_vba(code, 'Probe', timeout=30.0)
                assert result.ok, (name, result)
                base = f'{name}.xlsm'
                outputs = [path.name[len(base):] for path in Path(folder).glob(base + '*') if path.is_file()]
                rows.append({"name": name, "source": source, "result": result.value, "output_suffixes": outputs})
                print(name, result.value, flush=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "excel_version_build": installed.value, "log": LOG, "sink": SINK, "book": BOOK, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
