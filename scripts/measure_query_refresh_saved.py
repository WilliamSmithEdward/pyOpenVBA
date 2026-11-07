"""Does a workbook the model saved with a database query table open in Excel, and does Excel's own refresh of it work? An Access file is
made through Excel's ACE OLEDB provider; the model refreshes from it through a host function built on pyOpenVBA's own Access engine, saves,
and Excel opens the save, reads the table, refreshes it and reads it again. No network source is ever named.

Prediction: the file opens with no repair, its cells read as the model wrote them, and Excel's refresh gives the same cells, with its own
fuller connection string afterwards (the named difference).

    python scripts/measure_query_refresh_saved.py

writes tests/fixtures/query_refresh_saved.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_query_refresh_layout import MAKE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/query_refresh_saved.json"
PROVIDER = "Microsoft.ACE.OLEDB.16.0"
MODEL = """Sub Make(ByVal file As String, ByVal out As String)
    Dim q As QueryTable
    Set q = ActiveSheet.QueryTables.Add("OLEDB;Provider={provider};Data Source=" & file, ActiveSheet.Range("B2"), "SELECT * FROM T WHERE Id <= 3 ORDER BY Id")
    q.Refresh False
    ActiveSheet.Range("E3").Formula = "=B3*10"
    ActiveWorkbook.SaveAs out, 51
End Sub
""".replace("{provider}", PROVIDER)
EXCEL = r'''Private Function Grid(ByVal ws As Worksheet) As String
    Dim r As Long, c As Long, s As String
    For r = 1 To 7
        For c = 2 To 5
            s = s & "|" & TypeName(ws.Cells(r, c).Value) & ":" & ws.Cells(r, c).Text
        Next
    Next
    Grid = s
End Function
Function Check(ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(out, UpdateLinks:=0)
    If Err.Number <> 0 Then Check = "open E" & Err.Number & " " & Err.Description: Exit Function
    Set ws = wb.Worksheets(1)
    s = "opened tables=" & ws.QueryTables.Count & " conns=" & wb.Connections.Count
    Set q = ws.QueryTables(1)
    s = s & " result=" & q.ResultRange.Address & " name=" & q.Name & " type=" & q.QueryType & " command=" & q.CommandText
    s = s & vbTab & "before" & Grid(ws)
    Err.Clear
    q.CommandText = "SELECT * FROM T WHERE Id <= 5 ORDER BY Id"
    q.Refresh False
    s = s & vbTab & "refresh E" & Err.Number & " " & Err.Description & " result=" & q.ResultRange.Address
    s = s & vbTab & "after" & Grid(ws)
    wb.Close False
    Check = s
End Function
'''


def main() -> None:
    from pyopenvba.access import AccessDatabase
    from pyopenvba.apps.excel import ExcelApplication
    from pyvbaharness import ExcelSession, HarnessConfig

    with tempfile.TemporaryDirectory(prefix="pyopenvba_qs_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file, out = str(Path(folder) / "rows.accdb"), str(Path(folder) / "saved.xlsx")
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message

        def from_access(request: object) -> list[list[object]]:
            source = getattr(request, "source")
            with AccessDatabase(file) as db:
                found = db.execute(source.command_text)
            assert isinstance(found, list)
            names = list(found[0]) if found else ["Id", "Name"]
            return [names, *[[row[name] for name in names] for row in found]]

        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.set_hook("Refresh", from_access)
        app.add_module(MODEL, name="M")
        app.run("Make", file, out)
        checked = excel.run_vba(EXCEL, "Check", (out,), timeout=120.0, line_numbers=False)
        answer = str(checked.value).replace(folder, "<folder>") if checked.ok else f"[{checked.outcome}] {checked.message}"
    print(answer.replace("\t", "\n"))
    OUT.write_text(json.dumps({"check": answer.split("\t"), "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
