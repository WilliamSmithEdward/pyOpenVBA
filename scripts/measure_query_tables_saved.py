"""A saved query table read back by Excel: what its members answer after opening the file, what a refresh does, and what
Excel saves again. Each case opens a workbook made as scripts/make_query_table_workbook.py makes it, runs its statements with
q on the first table found and p on the one named, records r, the cells and the parts, and saves a copy.

    python scripts/measure_query_tables_saved.py

writes tests/fixtures/query_tables_saved.json, which tests/test_excel_query_tables_saved.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from make_query_table_workbook import SOURCE as MAKER
from measure_query_tables import READS_OF_ONE, parts

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/query_tables_saved.json'
FILES = {'d.csv': 'a,b,c\r\n1,2,3\r\n4,x,6\r\n', 's.txt': 'skip\r\nh;i\r\n7;8\r\n', 'd2.csv': 'p,q,r\r\n9,8,7\r\n6,5,4\r\n3,2,1\r\n'}
FIRST = 'Set q = ws.QueryTables("ExternalData_1"): '
SECOND = 'Set q = ws.QueryTables(1): Set q = ws.Range("F1").QueryTable: '
CASES: list[str] = ['r = ws.QueryTables.Count', 'r = ws.QueryTables(1).Name & " / " & ws.QueryTables(2).Name', 'r = ws.Range("A1").QueryTable.Name & " / " & ws.Range("F1").QueryTable.Name', 'r = TypeName(ws.Range("A20").QueryTable)', 'r = TypeName(ws.QueryTables("second one"))', 'r = TypeName(ws.QueryTables("second_one"))', 'r = ws.Names(1).Name & " " & ws.Names(2).Name']
CASES += [f'{FIRST}r = q.{name}' for name in READS_OF_ONE]
CASES += [f'{SECOND}r = q.{name}' for name in READS_OF_ONE]
CASES += [f'{FIRST}r = TypeName(q.{name})' for name in ('Parent', 'Destination', 'ResultRange', 'ListObject', 'Recordset', 'TextFileColumnDataTypes', 'TextFileFixedColumnWidths')]
RETARGET = 'q.Connection = "TEXT;" & folder & "\\{0}": '
CASES += ['r = "nothing"', f'{FIRST}{RETARGET.format("d.csv")}r = q.Refresh(False) & " " & q.ResultRange.Address', f'{FIRST}{RETARGET.format("d2.csv")}q.Refresh False: r = q.ResultRange.Address & " " & q.Destination.Address',
          f'{SECOND}{RETARGET.format("s.txt")}q.Refresh False: r = q.ResultRange.Address', f'{SECOND}{RETARGET.format("d2.csv")}q.TextFileSemicolonDelimiter = False: q.TextFileCommaDelimiter = True: q.Refresh False: r = q.ResultRange.Address', f'{FIRST}q.Refresh False: r = q.ResultRange.Address',
          f'{FIRST}q.AdjustColumnWidth = False: r = q.AdjustColumnWidth', f'{FIRST}q.Name = "renamed": r = q.Name & " " & ws.Names(1).Name & " " & ws.Names(2).Name', f'{SECOND}q.TextFileStartRow = 1: r = q.TextFileStartRow', f'{FIRST}q.Delete: r = ws.QueryTables.Count & " " & ws.Names.Count & " " & ws.Parent.Connections.Count', f'{SECOND}q.Delete: r = ws.QueryTables.Count & " " & ws.Parent.Connections.Count',
          f'{FIRST}q.Delete: ws.QueryTables(1).Delete: r = ws.QueryTables.Count & " " & ws.Parent.Connections.Count', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("K1")): r = q.Name & " " & ws.QueryTables.Count', 'Set q = ws.QueryTables.Add("TEXT;" & folder & "\\d.csv", ws.Range("K1")): q.TextFileCommaDelimiter = True: q.Refresh False: r = q.Name & " " & ws.QueryTables(1).Name',
          'ws.Rows(1).Insert: r = ws.Range("A2").QueryTable.Name & " " & ws.Range("A2").QueryTable.Destination.Address', 'ws.Range("A21").Value = "edited": r = 1', 'Set w2 = ws.Parent.Worksheets(2): r = w2.QueryTables.Count', f'{FIRST}ws.Range("A1").Value = "h": r = ws.Range("A1").Value']

# Copies of the sheet, in the workbook and out of it.
INNER = 'ws.Copy After:=ws: Set w2 = ws.Parent.Worksheets(2): '
OUTER = 'ws.Copy: Set w2 = ActiveSheet: '
for copy in (INNER, OUTER):
    CASES += [f'{copy}r = w2.Name & " " & w2.QueryTables.Count & " " & w2.QueryTables(1).Name & " " & w2.QueryTables(2).Name', f'{copy}r = w2.Parent.Connections.Count & " " & w2.Parent.Connections(1).Name & " " & w2.Parent.Connections(w2.Parent.Connections.Count).Name',
              f'{copy}r = w2.Names.Count & " " & w2.Names(1).Name & " " & w2.Names(1).RefersTo', f'{copy}r = w2.QueryTables("ExternalData_1").Connection & " " & w2.QueryTables("ExternalData_1").Destination.Address & " " & w2.QueryTables("ExternalData_1").ResultRange.Address',
              f'{copy}r = w2.QueryTables("second one").TextFileStartRow & " " & w2.QueryTables("second one").RefreshStyle & " " & w2.QueryTables("second one").BackgroundQuery & " " & w2.Range("F1").QueryTable.Name', f'{copy}w2.QueryTables("ExternalData_1").Refresh False: r = w2.QueryTables("ExternalData_1").ResultRange.Address & " " & ws.QueryTables("ExternalData_1").ResultRange.Address',
              f'{copy}w2.QueryTables("ExternalData_1").Name = "copied": r = w2.QueryTables(2).Name & " " & ws.QueryTables("ExternalData_1").Name', f'{copy}w2.QueryTables("ExternalData_1").Delete: r = w2.QueryTables.Count & " " & ws.QueryTables.Count']
CASES += [f'{OUTER}w2.Parent.SaveCopyAs folder & "\\o.xlsx": r = "saved"', f'{INNER}ws.Parent.SaveCopyAs folder & "\\i.xlsx": r = "saved"']


def _probe() -> str:
    lines = ['Private Sub Act(ByVal which As Long, ByVal folder As String, ws As Worksheet, r As Variant)', '    Dim q As Object, w2 As Object', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '', 'Private Function Dumped(ws As Worksheet) As String', '    Dim c As Range, s As String', '    For Each c In ws.UsedRange.Cells', '        If Not IsEmpty(c.Value) Then s = s & c.Address(False, False) & "=" & TypeName(c.Value) & ":" & c.Text & ";"', '    Next', '    Dumped = s', 'End Function', '',
              'Public Function Probe(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, s As String, v As String, r As Variant', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(folder & "\\saved.xlsx")', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Act which, folder, ws, r', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v & "|" & Dumped(ws)',
              '    Dim other As Workbook', '    If Workbooks.Count > 1 Then', '        For Each other In Workbooks', '            If Not other Is ThisWorkbook And Not other Is wb Then other.SaveCopyAs folder & "\\o" & which & ".xlsx": other.Close False: Exit For', '        Next', '    End If', '    On Error GoTo 0', '    wb.SaveCopyAs folder & "\\c" & which & ".xlsx"', '    wb.Close False', '    Application.DisplayAlerts = True', '    Probe = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for name, text in FILES.items():
            (Path(folder) / name).write_bytes(text.encode('ascii'))
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKER, 'Make', (folder,), timeout=120.0)
            assert made.ok and not made.dialogs, made
            for index in range(len(CASES)):
                result = excel.run_vba(source, 'Probe', (index, folder), timeout=120.0)
                answers.append(str(result.value) if result.ok and not result.dialogs else f'STOPPED {result.outcome}')
        saved = [parts(Path(folder) / f'c{index}.xlsx', folder) for index in range(len(CASES))]
        for index, found in enumerate(saved):
            # A copy in a workbook of its own, under names that say so.
            found.update({f'copy/{name}': text for name, text in parts(Path(folder) / f'o{index}.xlsx', folder).items()})
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'files': FILES, 'source': source,
              'cases': [[statement, re.sub(re.escape(folder), '<folder>', answer, flags=re.IGNORECASE), part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer, _ in record['cases']:
        print(statement, '=>', answer)


if __name__ == '__main__':
    main()
