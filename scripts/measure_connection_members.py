"""The rest of Workbook.Connections and its OLEDBConnection, Scenarios.CreateSummary and the file-naming members of
Workbooks. A connection's members are read on a copy of tests/fixtures/power_query/loaded_to_sheet.xlsx, whose Power Query
connection is an OLE DB one, and on connections a macro adds. Nothing connects to anything: the sources named are not
there. Workbooks.Close would end the harness's own workbook, so it logs to a file and is run on its own.

    python scripts/measure_connection_members.py

writes tests/fixtures/connection_members.json, which tests/test_excel_connection_members.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/connection_members.json'
FIXTURE = 'power_query/loaded_to_sheet.xlsx'
O = 'c.OLEDBConnection'
ADD = 'Set a = n.Connections.Add2("Mine", "About", "OLEDB;Provider=None;Data Source=nowhere", "Select 1", xlCmdSql)'
SUMMARY = ('ws.Range("A1").Value = 1: ws.Range("A2").Formula = "=A1*2": ws.Scenarios.Add "One", ws.Range("A1"), Array(5): '
           'ws.Scenarios.Add "Two", ws.Range("A1"), Array(9): ')
CASES = [
    # An OLE DB connection's settings.
    *[f'r = {O}.{name}' for name in ('LocaleID', 'MaxDrillthroughRecords', 'OLAP', 'RetrieveInOfficeUILang', 'Refreshing', 'RobustConnect', 'ServerFillColor', 'ServerFontStyle', 'ServerNumberFormat', 'ServerTextColor', 'ServerSSOApplicationID', 'SourceDataFile', 'RefreshDate')],
    f'r = TypeName({O}.ADOConnection)', f'r = TypeName({O}.CalculatedMembers)', f'{O}.CancelRefresh', f'r = {O}.CancelRefresh', f'{O}.Reconnect', f'{O}.SaveAsODC "nope"', f'{O}.MakeConnection',
    # The connection's.
    'r = TypeName(c.DataFeedConnection)', 'r = TypeName(c.Ranges)', 'r = c.Ranges.Count', 'r = c.Ranges(1).Address', 'r = c.Ranges(1).Parent.Name', 'r = TypeName(c.ModelTables)', 'r = c.ModelTables.Count', 'r = c', 'r = c.[_Default]',
    'c.Delete: r = wb.Connections.Count', 'c.Delete: r = wb.Queries.Count', 'r = c.Delete', 'c.Delete: r = wb.Worksheets(1).ListObjects.Count & "/" & wb.Worksheets(1).ListObjects(1).SourceType', 'c.Delete: r = wb.Worksheets(1).Range("A1").Value',
    # Queries.
    'r = wb.Queries.FastCombine', 'wb.Queries.FastCombine = True: r = wb.Queries.FastCombine', 'r = wb.Queries.Count', 'For Each x In wb.Queries: r = TypeName(x): Next', 'r = wb.Queries.[_Default](1).Name', 'r = TypeName(wb.Queries.Item(1))', 'For Each x In wb.Connections: r = TypeName(x): Next', 'r = wb.Connections.[_Default](1).Name',
    # Adding.
    f'{ADD}: r = TypeName(a) & "/" & n.Connections.Count & "/" & a.Name & "/" & a.Description & "/" & a.Type', f'{ADD}: r = TypeName(a.OLEDBConnection) & "/" & a.OLEDBConnection.Connection & "/" & a.OLEDBConnection.CommandText & "/" & a.OLEDBConnection.CommandType', f'{ADD}: a.Delete: r = n.Connections.Count', f'{ADD}: r = a.Ranges.Count', f'{ADD}: r = a.OLEDBConnection.BackgroundQuery & "/" & a.OLEDBConnection.RefreshOnFileOpen & "/" & a.OLEDBConnection.EnableRefresh & "/" & a.RefreshWithRefreshAll',
    f'{ADD}: {ADD}', f'{ADD}: Set a = n.Connections.Add2("mine", "", "OLEDB;Provider=None", "", xlCmdSql): r = a.Name', 'Set a = n.Connections.Add2("", "", "OLEDB;Provider=None", "", xlCmdSql): r = a.Name', 'Set a = n.Connections.Add2("Mine", "", "nonsense", "", xlCmdSql): r = a.Type', 'Set a = n.Connections.Add2("Mine", "", "", "", xlCmdSql): r = a.Type', 'Set a = n.Connections.Add2("Mine", "", "OLEDB;Provider=None", "", 99): r = a.OLEDBConnection.CommandType', 'Set a = n.Connections.Add2("Mine"): r = a.Name', 'Set a = n.Connections.Add2("Mine", "", "OLEDB;Provider=None"): r = a.OLEDBConnection.CommandType', 'Set a = n.Connections.Add2(): r = a.Name',
    'Set a = n.Connections.Add2("Mine", "", "TEXT;nowhere.csv", "", xlCmdSql): r = a.Type', 'Set a = n.Connections.Add2("Mine", "", "ODBC;DSN=nowhere", "", xlCmdSql): r = a.Type', f'{ADD}: a.Refresh', f'{ADD}: r = a.OLEDBConnection.IsConnected', f'{ADD}: a.Name = "Other": r = a.Name & "/" & n.Connections("Other").Name', f'{ADD}: a.Description = "Why": r = a.Description',
    'r = TypeName(n.Connections.AddFromFile("nope.odc"))', 'n.Connections.AddFromFile', 'r = TypeName(n.Connections.AddFromFile("nope.odc", True, False))',
    # Workbooks' files.
    'r = w.CanCheckOut(path)', 'w.CheckOut "nope.xlsx"', 'w.CheckOut path', 'w.OpenXML "nope.xml"', 'w.OpenDatabase "nope.accdb"', 'w.OpenXML', 'w.OpenDatabase',
    # A summary of scenarios.
    'ws.Scenarios.CreateSummary', 'r = ws.Scenarios.CreateSummary(xlStandardSummary)', f'{SUMMARY}ws.Scenarios.CreateSummary xlStandardSummary, ws.Range("A2"): r = Dumped(n)', f'{SUMMARY}ws.Scenarios.CreateSummary: r = Dumped(n)', f'{SUMMARY}ws.Scenarios.CreateSummary ResultCells:=ws.Range("A2"): r = Dumped(n)', f'{SUMMARY}ws.Scenarios.CreateSummary 99', f'{SUMMARY}r = TypeName(ws.Scenarios.CreateSummary(xlStandardSummary, ws.Range("A2")))', f'{SUMMARY}ws.Scenarios.CreateSummary xlStandardSummary, "x"',
    f'{SUMMARY}ws.Scenarios.CreateSummary xlStandardSummary, ws.Range("A2"): ws.Scenarios.CreateSummary xlStandardSummary, ws.Range("A2"): r = n.Sheets(1).Name & "/" & n.Sheets(2).Name & "/" & n.Sheets(3).Name & "/" & ActiveSheet.Name', f'{SUMMARY}ws.Scenarios.CreateSummary xlPivotTable, ws.Range("A2"): r = Dumped(n)',
    'ws.Name = "My Data": ws.Range("B2").Value = 1: ws.Range("B3").Formula = "=B2": ws.Names.Add "Input", ws.Range("B2"): n.Names.Add "Result", ws.Range("B3"): ws.Scenarios.Add "One", ws.Range("B2"), Array(5), , , True: ws.Scenarios.CreateSummary xlStandardSummary, ws.Range("B3"): r = Dumped(n)',
]

MODULE = r'''Option Explicit

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
    ElseIf IsArray(v) Then
        Shown = "array"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dumped(ByVal n As Object) As String
    Dim s As String, sheet As Object, cell As Object
    For Each sheet In n.Sheets
        s = s & "{" & sheet.Name & " " & sheet.UsedRange.Address & " out=" & sheet.Outline.SummaryRow & "/" & sheet.Outline.SummaryColumn
        For Each cell In sheet.UsedRange.Cells
            If cell.Formula <> "" Or cell.Interior.ColorIndex <> xlNone Then
                s = s & "~" & cell.Address(False, False) & "=" & cell.Formula & ";" & cell.NumberFormat & ";" & cell.Font.Bold & ";" & cell.Font.Color & ";" & cell.Interior.Color & ";" & cell.Interior.ColorIndex & ";" & cell.HorizontalAlignment & ";" & cell.Rows.OutlineLevel & ";" & cell.Borders(xlEdgeBottom).LineStyle & ";" & cell.Borders(xlEdgeTop).LineStyle
            End If
        Next
        s = s & "}"
    Next
    Dumped = s & "active " & ActiveSheet.Name & "!" & ActiveCell.Address(False, False)
End Function

Public Function Tried(ByVal which As Long, ByVal path As String) As String
    Dim wb As Object, c As Object, n As Object, ws As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Set n = Workbooks.Add(xlWBATWorksheet)
    Set ws = n.Worksheets(1)
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    Set c = wb.Connections(1)
    n.Activate
    On Error GoTo Failed
    Run which, wb, c, n, ws, path, r
    s = "ok " & Shown(r)
Answered:
    On Error Resume Next
    Application.DisplayAlerts = False
    Tried = s
    wb.Close False
    n.Close False
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function

Public Function CloseAll(ByVal log As String, ByVal folder As String) As String
    Dim file As Integer, a As Workbook, b As Workbook
    Application.DisplayAlerts = False
    Set a = Workbooks.Add(xlWBATWorksheet)
    Set b = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Range("A1").Value = 1
    file = FreeFile
    Open log For Append As #file
    Print #file, "before " & Workbooks.Count
    Close #file
    Workbooks.Close
    file = FreeFile
    Open log For Append As #file
    Print #file, "after " & Workbooks.Count
    Close #file
    CloseAll = "returned"
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, wb As Object, c As Object, n As Object, ws As Object, ByVal path As String, r As Variant)',
             '    Dim x As Object, a As Object, w As Object', '    Set w = Application.Workbooks', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines)


def main() -> None:
    source = _source()
    answers: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        copy = Path(folder) / 'book.xlsx'
        shutil.copy(ROOT / 'tests/fixtures' / FIXTURE, copy)
        index = 0
        while index < len(CASES):
            with ExcelSession(HarnessConfig(exclusive=False)) as excel:
                excel.new_document()
                while index < len(CASES):
                    result = excel.run_vba(source, 'Tried', (index, str(copy)), timeout=25.0)
                    index += 1
                    if not result.ok or result.dialogs:
                        text = '; '.join(getattr(dialog, 'message', '') for dialog in result.dialogs)
                        answers.append(f'breaks: {result.outcome} {text}')
                        print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
                        break
                    answer = re.sub(re.escape(folder), '<folder>', str(result.value), flags=re.IGNORECASE)
                    answers.append(re.sub(r'[A-Z]:\\Users\\[^\\]+\\[^\\]+', '<documents>', answer))
                    print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
        log = Path(folder) / 'close.log'
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'CloseAll', (str(log), folder), timeout=25.0)
            closing = {'outcome': str(result.outcome), 'value': str(result.value), 'log': log.read_text() if log.exists() else ''}
            print(closing, flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'file': FIXTURE, 'source': source, 'workbooks_close': closing,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
