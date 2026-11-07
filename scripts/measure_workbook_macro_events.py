"""Which of Workbook's unbuilt events a macro can fire, in live Excel.

A class module's WithEvents Workbook logs each event below while a macro refreshes a pivot table, moves a field, edits a value, adds a range
to the data model, builds a pivot table and a table over the model, and refreshes them. An event no action fires is not evidence that no
macro can fire it, only that these did not; the fixture records which action each event answered to. Each action runs in a new workbook,
in memory: no file, server or other program is named. Events stay off outside the action, as the harness leaves them.

    python scripts/measure_workbook_macro_events.py

writes tests/fixtures/workbook_macro_events.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_macro_events.json'
#: Each event and the arguments its handler takes, from the type library's signatures in docs/excel_checklist.csv. SheetPivotTableUpdate,
#: which is built, is the control: it should fire on a refresh.
EVENTS = {
    'ModelChange': 'ByVal Changes As ModelChanges',
    'PivotTableCloseConnection': 'ByVal Target As PivotTable',
    'PivotTableOpenConnection': 'ByVal Target As PivotTable',
    'RowsetComplete': 'ByVal Description As String, ByVal Sheet As String, ByVal Success As Boolean',
    'SheetPivotTableAfterValueChange': 'ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal TargetRange As Range',
    'SheetPivotTableBeforeAllocateChanges': 'ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, '
                                            'ByVal ValueChangeEnd As Long, Cancel As Boolean',
    'SheetPivotTableBeforeCommitChanges': 'ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, '
                                          'ByVal ValueChangeEnd As Long, Cancel As Boolean',
    'SheetPivotTableBeforeDiscardChanges': 'ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal ValueChangeStart As Long, '
                                           'ByVal ValueChangeEnd As Long',
    'SheetPivotTableChangeSync': 'ByVal Sh As Object, ByVal Target As PivotTable',
    'SheetPivotTableUpdate': 'ByVal Sh As Object, ByVal Target As PivotTable',
    'SheetTableUpdate': 'ByVal Sh As Object, ByVal Target As TableObject',
}
#: The two pivot events also on the sheet and the application, to place each sink in the order.
PIVOT = {'PivotTableChangeSync': 'ByVal Target As PivotTable', 'PivotTableUpdate': 'ByVal Target As PivotTable'}
SINK = 'Public WithEvents Book As Workbook\nPublic WithEvents Sheet As Worksheet\nPublic WithEvents App As Application\n' + ''.join(
    f'Private Sub Book_{name}({arguments})\n    Log = Log & "{name}|"\nEnd Sub\n' for name, arguments in EVENTS.items()) + ''.join(
    f'Private Sub Sheet_{name}({arguments})\n    Log = Log & "Sheet:{name}|"\nEnd Sub\n' for name, arguments in PIVOT.items()) + ''.join(
    f'Private Sub App_Sheet{name}(ByVal Sh As Object, {arguments})\n    Log = Log & "App:Sheet{name}|"\nEnd Sub\n' for name, arguments in PIVOT.items())
#: The data and a pivot table over it, made before events are turned on.
SETUP = '''    Set ws = ActiveSheet
    ws.Range("A1").Value = "Item": ws.Range("B1").Value = "Amount"
    ws.Range("A2").Value = "a": ws.Range("B2").Value = 1
    ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    ws.Range("A4").Value = "a": ws.Range("B4").Value = 3
    Set pt = ActiveWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:B4")).CreatePivotTable(ws.Range("E1"))
    pt.PivotFields("Item").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
'''
MODEL = '''ActiveWorkbook.Connections.Add2 "Probe range", "", "WORKSHEET;" & ActiveWorkbook.FullName, ws.Name & "!$A$1:$B$4", 7, True, False
'''
ACTIONS = {
    'pivot_refresh': 'pt.RefreshTable',
    'pivot_source_refresh': 'ws.Range("B2").Value = 10\npt.PivotCache.Refresh',
    'pivot_field_move': 'pt.PivotFields("Item").Orientation = xlColumnField',
    'pivot_value_edit': 'pt.DataBodyRange.Cells(1, 1).Value = 99',
    'pivot_create': 'Set mp = ActiveWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:B4")).CreatePivotTable(ws.Range("J1"))',
    'pivot_add_data': 'pt.AddDataField pt.PivotFields("Amount"), "Count", xlCount',
    'pivot_function': 'pt.DataFields(1).Function = xlCount',
    'pivot_clear': 'pt.ClearTable',
    'model_add': MODEL,
    'model_pivot': MODEL + '''Set mp = ActiveWorkbook.PivotCaches.Create(xlExternal, ActiveWorkbook.Connections("ThisWorkbookDataModel"), 6).CreatePivotTable(ws.Range("J1"))
mp.PivotCache.Refresh''',
    'model_table': MODEL + '''Set lo = ws.ListObjects.Add(SourceType:=4, Source:=ActiveWorkbook.Connections("Probe range"), Destination:=ws.Range("N1"))
lo.TableObject.Refresh''',
}


def install() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object', 'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)',
             'part.Name = "BookSink"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in SINK.splitlines()]
    lines += ['part.CodeModule.AddFromString text', 'Install = "installed"', 'End Function']
    return '\n'.join(lines)


def probe(action: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Worksheet, pt As PivotTable, mp As PivotTable, lo As ListObject, e As Long, d As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
{SETUP}    Dim sink As New BookSink
    Set sink.Book = ActiveWorkbook
    Set sink.Sheet = ws
    Set sink.App = Application
    Log = ""
    Application.EnableEvents = True
    On Error Resume Next
{action}
    e = Err.Number: d = Err.Description
    On Error GoTo 0
    Application.EnableEvents = False
    Application.DisplayAlerts = True
    Probe = "e" & e & IIf(e <> 0, " " & d, "") & "~" & Log
End Function
'''


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        for name, action in ACTIONS.items():
            excel.new_document()
            excel.add_module('EventLog', 'Public Log As String\n')
            installed = excel.run_vba(install(), 'Install', timeout=60, line_numbers=False)
            assert installed.ok, (installed.error, installed.dialogs)
            result = excel.run_vba(probe(action), 'Probe', timeout=180, line_numbers=False)
            answers[name] = str(result.value) if result.ok else f'harness: {result.error} {[d.message for d in result.dialogs]}'[:300]
            print(name, '->', ascii(answers[name])[:300], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'events': list(EVENTS), 'answers': answers}, indent=1) + '\n',
                   encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
