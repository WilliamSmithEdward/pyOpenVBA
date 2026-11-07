"""What Power Query makes of a list or table of a vast size, evaluated in live Excel's own Power Query: whether it builds it lazily, refuses it, or works.

Each case is a query of its own, `let x = EXPRESSION in Text.From(x)`, loaded to a sheet and refreshed in the foreground, smallest risk first:
reads that need only the first item, then counts. A case that runs past its time is recorded as such and Excel closed. No file, server or web
source is named.

    python scripts/measure_m_vast_lists.py

writes tests/fixtures/m_vast_lists.json, keeping each case as it finishes and asking only the rest.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/m_vast_lists.json'
TABLE = '#table({"a"}, {{1}})'
CASES = [
    'List.First(List.Numbers(1, 1e12))', 'List.First(List.Numbers(1, 3e9))', 'List.First({1..1e12})', 'List.First({1..3e9})', 'List.First({1..2147483648})',
    'List.First({1..2147483647})', 'Table.FirstN(Table.Repeat(' + TABLE + ', 1e12), 1)[a]{0}', 'List.First(List.Numbers(1, 1.8e308))', 'List.First({1..1.8e308})',
    'Table.FirstN(Table.Repeat(' + TABLE + ', 1.8e308), 1)[a]{0}', 'List.Count(List.Numbers(1, 1e7))', 'List.Count({1..1e7})',
    'Table.RowCount(Table.Repeat(' + TABLE + ', 1e7))', 'List.Count(List.Numbers(1, 1e12))', 'List.Count({1..1e12})', 'Table.RowCount(Table.Repeat(' + TABLE + ', 1e12))',
]
TIMEOUT = 90.0
SOURCE = '''Function Probe(ByVal formula As String) As String
    Dim ws As Worksheet, lo As ListObject, e As Long, d As String, shown As String, q As Object, t As Single
    Application.DisplayAlerts = False
    Set ws = Worksheets.Add
    For Each q In ThisWorkbook.Queries
        q.Delete
    Next
    ThisWorkbook.Queries.Add Name:="Q", Formula:=formula
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Q;Extended Properties=""""", Destination:=ws.Range("$A$1"))
    lo.QueryTable.CommandType = xlCmdSql
    lo.QueryTable.CommandText = Array("SELECT * FROM [Q]")
    t = Timer
    On Error Resume Next
    lo.QueryTable.Refresh BackgroundQuery:=False
    e = Err.Number
    d = Err.Description
    Err.Clear
    On Error GoTo 0
    If e = 0 Then shown = CStr(ws.Range("A2").Value) Else shown = "error " & e & ": " & d
    Probe = shown & " | " & Format(Timer - t, "0.0") & " s"
    lo.Delete
    ThisWorkbook.Queries("Q").Delete
    ws.Delete
    Application.DisplayAlerts = True
End Function
'''


def formula(expression: str) -> str:
    return f'let x = {expression} in Text.From(x)'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    for case in CASES:
        if case in answers:
            continue
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                result = excel.run_vba(SOURCE, 'Probe', (formula(case),), timeout=TIMEOUT, line_numbers=False)
            if result.dialogs:
                answers[case] = 'dialog: ' + ' '.join(result.dialogs[0].message.split())[:200]
            elif result.ok:
                answers[case] = str(result.value)
            else:
                # The harness's message names a screenshot on this machine, which no committed file may.
                answers[case] = f'ran past {TIMEOUT:.0f} s'
        except HarnessError as failed:
            answers[case] = f'harness: {str(failed)[:120]}'
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8',
                       newline='\n')
        print(ascii(case[:60]), '->', ascii(answers[case][:160]), flush=True)


if __name__ == '__main__':
    main()
