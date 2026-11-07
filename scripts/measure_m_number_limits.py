"""What Power Query makes of number literals near and past the largest Double, hex and decimal, evaluated in live Excel's own Power Query.

Each literal is a query of its own, `let x = LITERAL in Text.From(x)`, loaded to a sheet and refreshed in the foreground; the cell's text, or the
refresh's error message, is recorded. The queries read nothing outside the workbook: no file, server or web source is named.

    python scripts/measure_m_number_limits.py

writes tests/fixtures/m_number_limits.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/m_number_limits.json'
LITERALS = [
    '0x1FFFFFFFFFFFFF', '0x20000000000001', '0xFFFFFFFFFFFFFFFF', '0x' + 'F' * 255, '0x' + 'F' * 256, '0x' + '1' + '0' * 256, '0xe' + '1' * 300,
    '1e14', '1e15', '1e16', '123456789012345', '1234567890123456', '12345678901234567', '9007199254740993', '2e20', '-1.8e308', '0.1', '1e-5', '1e-7',
    '1e308', '1.7976931348623157e308', '1.7976931348623158e308', '1.8e308', '1e309', '1' + '0' * 400, '0.' + '0' * 400 + '1', '1e-400',
]
SOURCE = '''Function Probe(ByVal formula As String) As String
    Dim ws As Worksheet, lo As ListObject, e As Long, d As String, shown As String, q As Object
    Application.DisplayAlerts = False
    Set ws = Worksheets.Add
    For Each q In ThisWorkbook.Queries
        q.Delete
    Next
    ThisWorkbook.Queries.Add Name:="Q", Formula:=formula
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Q;Extended Properties=""""", Destination:=ws.Range("$A$1"))
    lo.QueryTable.CommandType = xlCmdSql
    lo.QueryTable.CommandText = Array("SELECT * FROM [Q]")
    On Error Resume Next
    lo.QueryTable.Refresh BackgroundQuery:=False
    e = Err.Number
    d = Err.Description
    Err.Clear
    On Error GoTo 0
    If e = 0 Then shown = CStr(ws.Range("A2").Value) Else shown = "error " & e & ": " & d
    Probe = shown
    lo.Delete
    ThisWorkbook.Queries("Q").Delete
    ws.Delete
    Application.DisplayAlerts = True
End Function
'''


def formula(literal: str) -> str:
    return f'let x = {literal} in Text.From(x)'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    for literal in LITERALS:
        if literal in answers:
            continue
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(SOURCE, 'Probe', (formula(literal),), timeout=120.0, line_numbers=False)
        if not result.ok or result.dialogs:
            raise SystemExit(f'{literal[:40]} stopped: {result.message} {[d.message[:80] for d in result.dialogs]}')
        answers[literal] = str(result.value)
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8',
                       newline='\n')
        print(ascii(literal[:44]), '->', ascii(answers[literal][:120]), flush=True)


if __name__ == '__main__':
    main()
