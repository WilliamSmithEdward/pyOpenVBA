"""Measure Excel document handlers alongside Worksheet/Workbook/Application sinks."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/excel_event_sinks.json"
SINK = '''Public WithEvents Sheet As Worksheet
Public WithEvents Book As Workbook
Public WithEvents App As Application
Private Sub Sheet_Change(ByVal Target As Range)
Log = Log & "sheet-sink:" & Target.Address(False, False) & "|"
End Sub
Private Sub Book_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "book-sink:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "app-sink:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
'''
SHEET = '''Private Sub Worksheet_Change(ByVal Target As Range)
Log = Log & "sheet-own:" & Target.Address(False, False) & "|"
End Sub
'''
BOOK = '''Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "book-own:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
'''
for event, signature, suffix in (
    ("Calculate", "", '"|"'),
    ("Activate", "", '"|"'),
    ("Deactivate", "", '"|"'),
    ("SelectionChange", "ByVal Target As Range", '":" & Target.Address(False, False) & "|"'),
):
    SINK += f'Private Sub Sheet_{event}({signature})\nLog = Log & "sheet-sink-{event}" & {suffix}\nEnd Sub\n'
    SHEET += f'Private Sub Worksheet_{event}({signature})\nLog = Log & "sheet-own-{event}" & {suffix}\nEnd Sub\n'
    combined = "ByVal Sh As Object" + (", " + signature if signature else "")
    for prefix, label in (("Book", "book-sink"), ("App", "app-sink")):
        SINK += f'Private Sub {prefix}_Sheet{event}({combined})\nLog = Log & "{label}-{event}:" & Sh.Name & {suffix}\nEnd Sub\n'
    BOOK += f'Private Sub Workbook_Sheet{event}({combined})\nLog = Log & "book-own-{event}:" & Sh.Name & {suffix}\nEnd Sub\n'
SINK += '''Private Sub Book_NewSheet(ByVal Sh As Object)
Log = Log & "book-sink-NewSheet:" & Sh.Name & "|"
End Sub
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
Log = Log & "app-sink-NewSheet:" & Sh.Name & "|"
End Sub
'''
BOOK += '''Private Sub Workbook_NewSheet(ByVal Sh As Object)
Log = Log & "book-own-NewSheet:" & Sh.Name & "|"
End Sub
'''
ACTIONS = {
    "all_sinks": 'Range("A1").Value = 5',
    "events_disabled": 'Application.EnableEvents = False\nRange("A1").Value = 5',
    "sheet_sink_detached": 'Set sink.Sheet = Nothing\nRange("A1").Value = 5',
    "book_sink_detached": 'Set sink.Book = Nothing\nRange("A1").Value = 5',
    "app_sink_detached": 'Set sink.App = Nothing\nRange("A1").Value = 5',
    "multiple_areas": 'Range("A1,C3").Value = 5',
    "calculate_after_edit": 'Range("D1").Formula = "=A1*2"\nLog = ""\nRange("A1").Value = 5',
    "calculate_all": 'Range("D1").Formula = "=A1*2"\nLog = ""\nApplication.Calculate',
    "calculate_manual": 'Range("D1").Formula = "=A1*2"\nApplication.Calculation = -4135\nRange("A1").Value = 5\nLog = ""\nApplication.Calculate',
    "calculate_volatile": 'Range("D1").Formula = "=RAND()"\nLog = ""\nApplication.Calculate',
    "calculate_manual_volatile": 'Range("D1").Formula = "=RAND()"\nApplication.Calculation = -4135\nLog = ""\nApplication.Calculate',
    "selection_change": 'Range("B2").Select',
    "new_sheet": 'Worksheets.Add',
    "activation": 'Worksheets.Add\nLog = ""\nWorksheets("Sheet1").Activate',
}


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        lines = ['Function Install() As String', 'Dim text As String, part As Object']
        for owner, code in ((None, SINK), ('Worksheets(1).CodeName', SHEET), ('"ThisWorkbook"', BOOK)):
            if owner is None:
                lines += ['Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"']
            else:
                lines += [f'Set part = ActiveWorkbook.VBProject.VBComponents({owner})']
            lines += ['text = ""']
            lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
            lines += ['part.CodeModule.AddFromString text']
        lines += ['Install = Worksheets(1).CodeName & "|" & Application.Version & "|" & Application.Build', 'End Function']
        rows = []
        for name, action in ACTIONS.items():
            excel.new_document()
            excel.add_module("EventLog", "Public Log As String\n")
            installed = excel.run_vba('\n'.join(lines), 'Install', timeout=30.0)
            assert installed.ok, installed
            code = '''Function Probe() As String
Application.EnableEvents = False
Application.DisplayAlerts = False
Dim i As Long
For i = Worksheets.Count To 1 Step -1
If Worksheets(i).Name <> "Sheet1" Then Worksheets(i).Delete
Next
Worksheets("Sheet1").Activate
Cells.Clear
Range("A1").Select
Dim sink As New ExcelSink
Set sink.Sheet = ActiveSheet
Set sink.Book = ActiveWorkbook
Set sink.App = Application
Application.EnableEvents = True
Log = ""
''' + action + '\nProbe = Log\nApplication.EnableEvents = False\nEnd Function\n'
            result = excel.run_vba(code, 'Probe', timeout=30.0)
            assert result.ok, (name, result)
            rows.append({"name": name, "source": code, "result": result.value})
            print(name, result.value, flush=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "sheet_codename_excel_version_build": installed.value, "sink": SINK, "sheet": SHEET, "book": BOOK, "probes": rows}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
