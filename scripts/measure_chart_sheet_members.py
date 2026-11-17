"""Measure Chart.PrintedCommentPages and Chart.SaveAs on a chart sheet and an embedded chart, and a chart sheet's Activate
and Deactivate events with the workbook's SheetActivate, SheetDeactivate and NewSheet, through WithEvents sinks. Each
probe works in a workbook of its own, which it closes. Needs Excel:

    python scripts/measure_chart_sheet_members.py

writes tests/fixtures/chart_sheet_members.json (see vba_probe_fixture.py).
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_sheet_members.json'

CLASSES = {
    'CChartSink': '''Option Explicit
Public WithEvents Ch As Chart
Public Log As String
Private Sub Ch_Activate()
    Log = Log & "chart act:" & Ch.Name & ";"
End Sub
Private Sub Ch_Deactivate()
    Log = Log & "chart deact:" & Ch.Name & ";"
End Sub
''',
    'CSheetSink': '''Option Explicit
Public WithEvents Ws As Worksheet
Public Log As String
Private Sub Ws_Activate()
    Log = Log & "sheet act;"
End Sub
Private Sub Ws_Deactivate()
    Log = Log & "sheet deact;"
End Sub
''',
    'CBookSink': '''Option Explicit
Public WithEvents Wb As Workbook
Public Log As String
Private Sub Wb_SheetActivate(ByVal Sh As Object)
    Log = Log & "book sheetact:" & TypeName(Sh) & ":" & Sh.Name & ";"
End Sub
Private Sub Wb_SheetDeactivate(ByVal Sh As Object)
    Log = Log & "book sheetdeact:" & TypeName(Sh) & ":" & Sh.Name & ";"
End Sub
Private Sub Wb_NewSheet(ByVal Sh As Object)
    Log = Log & "book new:" & TypeName(Sh) & ";"
End Sub
''',
}

#: A workbook of one sheet, One, with data and a chart sheet of it, Chart1, made active.
BOOK = '''Dim wb As Workbook, ws As Worksheet, ch As Chart
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Name = "One"
ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
ws.Range("A2:C2").Value = Array("a", 10, 4)
ws.Range("A3:C3").Value = Array("b", 20, 6)
Set ch = wb.Charts.Add
ch.SetSourceData ws.Range("A1:C3")
'''
#: The same with three sinks hooked, events on.
SINKS = '''Dim c As New CChartSink, s As New CSheetSink, b As New CBookSink
Set c.Ch = ch
Set s.Ws = ws
Set b.Wb = wb
Application.EnableEvents = True
'''
LOGS = 'Probe = "[" & c.Log & "][" & s.Log & "][" & b.Log & "]"\nApplication.EnableEvents = False\nwb.Close False'
EMBEDDED = 'Dim e As Object\nSet e = ws.ChartObjects.Add(48, 96, 240, 144).Chart\n'
PATH = 'Dim path As String\npath = Environ("TEMP") & "\\pyov_chart_sheet_save_as.xlsx"\nIf Dir(path) <> "" Then Kill path\n'

CASES = {
    'pages_sheet': BOOK + 'Probe = TypeName(ch.PrintedCommentPages) & " " & ch.PrintedCommentPages\nwb.Close False',
    'pages_embedded': BOOK + EMBEDDED + 'On Error Resume Next\nDim v As Variant\nv = e.PrintedCommentPages\n'
                      'Probe = Err.Number & " " & TypeName(v) & " " & v\nOn Error GoTo 0\nwb.Close False',
    'save_as_sheet': BOOK + PATH + 'ch.SaveAs path, xlOpenXMLWorkbook\n'
                     'Probe = wb.Name & "|" & (Dir(path) <> "") & "|" & wb.Sheets.Count & "|" & ActiveSheet.Name\n'
                     'wb.Close False\nKill path',
    'save_as_embedded': BOOK + EMBEDDED + PATH + 'On Error Resume Next\ne.SaveAs path, xlOpenXMLWorkbook\n'
                        'Probe = Err.Number & "|" & (Dir(path) <> "") & "|" & (wb.Path = "")\nOn Error GoTo 0\n'
                        'wb.Close False\nIf Dir(path) <> "" Then Kill path',
    'events_chart_to_sheet': BOOK + SINKS + 'ws.Activate\n' + LOGS,
    'events_sheet_to_chart': BOOK + 'ws.Activate\n' + SINKS + 'ch.Activate\n' + LOGS,
    'events_chart_select': BOOK + 'ws.Activate\n' + SINKS + 'ch.Select\n' + LOGS,
    'events_chart_to_chart': BOOK + 'Dim ch2 As Chart\nSet ch2 = wb.Charts.Add\n' + SINKS + 'ch.Activate\nch2.Activate\n' + LOGS,
    'events_charts_add': BOOK + 'ws.Activate\n' + SINKS + 'wb.Charts.Add\n' + LOGS,
    'events_off': BOOK + SINKS + 'Application.EnableEvents = False\nws.Activate\nch.Activate\n' + LOGS,
}


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, CLASSES)


if __name__ == '__main__':
    main()
