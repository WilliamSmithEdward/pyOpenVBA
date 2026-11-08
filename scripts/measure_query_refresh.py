"""What a query table's refresh from a database puts on the sheet, from a source made for the run: an Access file, reached through Excel's own
ACE OLEDB provider. No network source is ever named.

Part one, types: one column for each ADO type, read back cell by cell, value with its type, number format and text, so the model can lay
out what a host returns as Excel does. Part two, BackgroundQuery: what a macro reads right after Refresh with it True and with it False,
and once DoEvents has run. Part three, what Excel writes back of a connection string that holds a password.

Predictions, written before the run:
- Text, Memo: String, General. Text that looks like a number ("00123", "1E3") stays a String, General.
- Long, Double: Double, General. Currency: Currency, a currency number format. Decimal: a Double, or Variant Decimal, General.
- Date with no time: Date, a short date format; with a time: Date, a date and time format. YesNo: Boolean, General.
- Null: Empty.
- BackgroundQuery True: right after Refresh the sheet is still empty and Refreshing is True; after DoEvents it is filled.
- The connection string in the file keeps the password only when SavePassword is True.

    python scripts/measure_query_refresh.py

writes tests/fixtures/query_refresh.json.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/query_refresh.json"
PROVIDER = "Microsoft.ACE.OLEDB.16.0"
#: Each column: its name, its SQL type, and two values as SQL literals.
COLUMNS = [("T", "TEXT(20)", "'abc'", "'00123'"), ("N", "TEXT(20)", "'1E3'", "'12.50'"), ("M", "MEMO", "'long text'", "NULL"),
           ("L", "LONG", "7", "-3"), ("D", "DOUBLE", "1.5", "2"), ("C", "CURRENCY", "12.34", "0"), ("X", "DECIMAL(10,2)", "3.25", "4"),
           ("DT", "DATETIME", "#2024-01-15#", "#2024-02-20 13:30:00#"), ("B", "YESNO", "True", "False"), ("E", "TEXT(5)", "NULL", "''")]
MAKE = r'''Function MakeDb(ByVal file As String, ByVal password As String) As String
    Dim cat As Object, cn As Object
    Set cat = CreateObject("ADOX.Catalog")
    If password = "" Then
        cat.Create "Provider={provider};Data Source=" & file
    Else
        cat.Create "Provider={provider};Data Source=" & file & ";Jet OLEDB:Database Password=" & password
    End If
    Set cn = cat.ActiveConnection
    cn.Execute "CREATE TABLE T ({columns})"
{inserts}    cn.Close
    MakeDb = "made"
End Function
'''
READ = r'''Private Function Cell(ByVal r As Range) As String
    Dim v As Variant
    v = r.Value
    Cell = TypeName(v) & ":" & CStr(v) & "|" & r.NumberFormat & "|" & r.Text
End Function
Function Types(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String, r As Long, c As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set q = ws.QueryTables.Add("OLEDB;Provider={provider};Data Source=" & file, ws.Range("A1"), "SELECT * FROM T")
    q.Refresh False
    s = q.ResultRange.Address
    For r = 1 To 3
        For c = 1 To {count}
            s = s & vbTab & Cell(ws.Cells(r, c))
        Next
    Next
    wb.Close False
    Types = s
End Function
Function Background(ByVal file As String, ByVal wait As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String, t As Double
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set q = ws.QueryTables.Add("OLEDB;Provider={provider};Data Source=" & file, ws.Range("A1"), "SELECT * FROM T")
    s = "default " & q.BackgroundQuery
    q.Refresh
    s = s & "|after refresh: A2=" & TypeName(ws.Range("A2").Value) & ":" & ws.Range("A2").Value & " refreshing=" & q.Refreshing
    On Error Resume Next
    s = s & " result=" & q.ResultRange.Address
    If Err.Number <> 0 Then s = s & " result=E" & Err.Number
    Err.Clear
    On Error GoTo 0
    If wait Then
        t = Timer
        Do While q.Refreshing And Timer - t < 10
            DoEvents
        Loop
        s = s & "|after DoEvents: A2=" & TypeName(ws.Range("A2").Value) & ":" & ws.Range("A2").Value & " refreshing=" & q.Refreshing
    End If
    wb.Close False
    Background = s
End Function
Function Saved(ByVal file As String, ByVal password As String, ByVal keep As Boolean, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set q = ws.QueryTables.Add("OLEDB;Provider={provider};Data Source=" & file & ";Jet OLEDB:Database Password=" & password, ws.Range("A1"), "SELECT T FROM T")
    q.SavePassword = keep
    q.Refresh False
    Saved = q.Connection & "|" & wb.Connections(1).OLEDBConnection.Connection
    wb.SaveAs out, 51
    wb.Close False
End Function
'''


def source() -> str:
    names = ", ".join(f"{name} {kind}" for name, kind, *_ in COLUMNS)
    inserts = "".join(f'    cn.Execute "INSERT INTO T VALUES ({", ".join(column[2 + row] for column in COLUMNS)})"\n' for row in range(2))
    return (MAKE.replace("{columns}", names).replace("{inserts}", inserts) + READ.replace("{count}", str(len(COLUMNS)))).replace("{provider}", PROVIDER)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, object] = {"provider": PROVIDER, "columns": COLUMNS}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_qr_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        code = source()

        def run(entry: str, *arguments: object) -> str:
            answer = excel.run_vba(code, entry, arguments, timeout=120.0, line_numbers=False)
            assert answer.ok, (entry, answer.outcome, answer.message)
            return str(answer.value).replace(folder, "<folder>")

        plain, locked = str(Path(folder) / "plain.accdb"), str(Path(folder) / "locked.accdb")
        run("MakeDb", plain, "")
        run("MakeDb", locked, "Secret1")
        types = run("Types", plain).split("\t")
        record["types"] = {"result": types[0], "cells": types[1:]}
        record["background"] = {"no wait": run("Background", plain, False), "wait": run("Background", plain, True)}
        saved: dict[str, object] = {}
        for keep in (False, True):
            out = str(Path(folder) / f"saved_{keep}.xlsx")
            said = run("Saved", locked, "Secret1", keep, out)
            with zipfile.ZipFile(out) as package:
                parts = {name: package.read(name).decode("utf-8") for name in package.namelist() if "connection" in name or "queryTable" in name}
            saved[f"SavePassword={keep}"] = {"read": said, "parts": {name: re.sub(re.escape(folder), "<folder>", text, flags=re.I) for name, text in parts.items()}}
        record["saved"] = saved
    record["measured_at"] = datetime.now(timezone.utc).isoformat()
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in record.items() if key != "saved"}, indent=1)[:3500])
    for key, value in saved.items():
        print(key, value["read"][:300])
        for name, text in value["parts"].items():
            print("  ", name, text[:700])


if __name__ == "__main__":
    main()
