"""Sheets.FillAcrossSheets: what Excel copies from a range to the same cells of every sheet of the group.

Each case builds a workbook of three sheets, Sheet1 holding values, a formula and formats in A1:B2 and Sheet2
and Sheet3 other content, runs one FillAcrossSheets, and records each sheet's A1:C3 -- formula, number format,
bold, fill colour -- and column A's width, or the error raised.

    python scripts/measure_fill_across_sheets.py

writes tests/fixtures/fill_across_sheets.json, which tests/test_excel_fill_across_sheets.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/fill_across_sheets.json'

#: Each case: the statement run on the workbook wb.
CASES = [
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), xlFillWithAll',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), xlFillWithContents',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), xlFillWithFormats',
    'wb.Sheets(Array("Sheet1", "Sheet2", "Sheet3")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Worksheets.FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Sheets(Array("Sheet2", "Sheet3")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet2").Range("A1:C3")',
    'wb.Sheets(Array("Sheet1")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:A2,B1")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), 9',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("C3")',
    'wb.Sheets(Array("Sheet3", "Sheet1")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2")',
    'wb.Sheets(Array("Sheet1", "Sheet2")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), "x"',
]


def _probe() -> str:
    lines = ['Private Function Snapshot(ByVal wb As Workbook) As String', '    Dim ws As Worksheet, c As Range, s As String',
             '    For Each ws In wb.Worksheets', '        s = s & "~" & ws.Name & ":" & ws.Columns(1).ColumnWidth',
             '        For Each c In ws.Range("A1:C3")',
             '            s = s & ";" & c.Formula & "," & c.NumberFormat & "," & c.Font.Bold & "," & c.Interior.Color',
             '        Next', '    Next', '    Snapshot = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal which As Long) As String', '    Dim wb As Workbook',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    wb.Worksheets(1).Name = "Sheet1"',
             '    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Sheet2"',
             '    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Sheet3"',
             '    With wb.Worksheets("Sheet1")', '        .Range("A1").Value = 1', '        .Range("B1").Formula = "=A1*2"',
             '        .Range("A2").Value = "text"', '        .Range("B2").Value = 0.5',
             '        .Range("B2").NumberFormat = "0.0%"', '        .Range("A1").Font.Bold = True',
             '        .Range("A2").Interior.Color = RGB(255, 255, 0)', '        .Columns(1).ColumnWidth = 20',
             '    End With', '    With wb.Worksheets("Sheet2")', '        .Range("A1:C3").Value = "old"',
             '        .Range("C3").Font.Bold = True', '        .Range("B2").Interior.Color = RGB(0, 0, 255)',
             '    End With', '    wb.Worksheets("Sheet3").Range("A1").Value = 99', '    wb.Worksheets("Sheet1").Activate',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines.append(f'    Case {index}: {case}')
    lines += ['    End Select', '    Tried = Snapshot(wb)', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description & "~" & Snapshot(wb)', '    Resume Done',
              'End Function', '', 'Public Function Probe() As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "|" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


#: A sink for the Change events a fill raises; the harness runs with events off, so the macro turns them on and puts
#: the setting back.
SINK = '''Public WithEvents App As Application
Public Log As String

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "|change " & Sh.Name & "!" & Target.Address(False, False)
End Sub
'''

EVENTS = '''Private sink As FillSink

Public Function Events() As String
    Dim wb As Workbook, s As String, wasOn As Boolean, t As Variant
    wasOn = Application.EnableEvents
    Application.EnableEvents = True
    Application.DisplayAlerts = False
    For Each t In Array(xlFillWithAll, xlFillWithContents, xlFillWithFormats)
        Set wb = Workbooks.Add(xlWBATWorksheet)
        wb.Worksheets(1).Name = "Sheet1"
        wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Sheet2"
        wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Sheet3"
        wb.Worksheets("Sheet1").Activate
        wb.Worksheets("Sheet1").Range("A1:B2").Value = 1
        wb.Worksheets("Sheet1").Range("D4").Select
        Set sink = New FillSink
        Set sink.App = Application
        wb.Sheets(Array("Sheet1", "Sheet2", "Sheet3")).FillAcrossSheets wb.Worksheets("Sheet1").Range("A1:B2"), t
        s = s & "#" & t & ":" & sink.Log & " active " & ActiveSheet.Name & " selection " & Selection.Address(False, False)
        Set sink.App = Nothing
        Set sink = Nothing
        wb.Close False
    Next
    Application.DisplayAlerts = True
    Application.EnableEvents = wasOn
    Events = Mid(s, 2)
End Function
'''


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
        excel.add_module('FillSink', SINK, kind='class')
        events = excel.run_vba(EVENTS, 'Events', timeout=120.0, module_name='FillEvents')
        assert events.ok and not events.dialogs, events
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'sink': SINK,
              'events_source': EVENTS, 'events': str(events.value), 'cases': dict(zip(CASES, answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(case, '\n   ', answer.replace('~', '\n    ')[:900])


if __name__ == '__main__':
    main()
