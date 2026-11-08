"""Do books Excel saved with a query table over a database, opened, refreshed and saved again by the model, open in Excel, and does
Excel's own refresh of them work? The books are tests/fixtures/query_refresh_loaded/, with C:\\data written back as the run's folder,
which holds an Access file made through Excel's own ACE OLEDB provider and the CSV file. No network source is ever named.

Prediction: each opens with no repair, reads as the model wrote it, a plain query table or a ListObject as it was, and Excel's refresh
for three rows gives the rows it gave in the book it made (tests/fixtures/query_refresh_loaded.json).

    python scripts/measure_query_refresh_loaded_saved.py

writes tests/fixtures/query_refresh_loaded_saved.json.
"""
from __future__ import annotations

import io
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_query_refresh import PROVIDER  # noqa: E402
from measure_query_refresh_layout import MAKE  # noqa: E402
from measure_query_refresh_loaded import BOOKS, CSV, FOLDER, PROBE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/query_refresh_loaded_saved.json"
MODEL = """Sub Five(ByVal kind As String, ByVal source As String)
    Dim q As QueryTable
    If kind = "table" Then
        Set q = ActiveWorkbook.Worksheets(1).ListObjects(1).QueryTable
    Else
        Set q = ActiveWorkbook.Worksheets(1).QueryTables(1)
    End If
    q.CommandText = "SELECT * FROM " & source & " WHERE Id <= 5 ORDER BY Id"
    q.Refresh False
End Sub
"""
CHECK = PROBE + r'''
Function Check(ByVal path As String, ByVal kind As String, ByVal source As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    If Err.Number <> 0 Then Check = "open E" & Err.Number & " " & Err.Description: Exit Function
    Set ws = wb.Worksheets(1)
    s = "querytables=" & ws.QueryTables.Count & " listobjects=" & ws.ListObjects.Count & " connections=" & wb.Connections.Count
    If kind = "table" Then Set q = ws.ListObjects(1).QueryTable Else Set q = ws.QueryTables(1)
    If q Is Nothing Then Check = s & " no table E" & Err.Number: wb.Close False: Exit Function
    s = s & vbTab & "opened " & Place(wb, ws, q)
    Err.Clear
    q.CommandText = "SELECT * FROM " & source & " WHERE Id <= 3 ORDER BY Id"
    q.Refresh False
    s = s & vbTab & "three E" & Err.Number & " " & Err.Description & " " & Place(wb, ws, q)
    wb.Close False
    Check = s
End Function
'''


def unmasked(data: bytes, folder: str) -> bytes:
    found = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(found, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            target.writestr(info, source.read(info).replace(FOLDER.encode("utf-8"), folder.encode("utf-8")))
    return found.getvalue()


def main() -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyvbaharness import ExcelSession, HarnessConfig

    def numbered(request: object) -> list[list[object]]:
        source = getattr(request, "source")
        wanted = int(re.search(r"Id <= (\d+)", source.command_text).group(1))  # type: ignore[union-attr]
        return [["Id", "Name"], *[[float(at), f"n{at}"] for at in range(1, 7) if at <= wanted]]

    measured: dict[str, list[str]] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_qls_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        (Path(folder) / "rows.csv").write_bytes(CSV.encode("ascii"))
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for book in sorted(BOOKS.glob("*.xlsx")):
            case = book.stem
            kind, source = case.split("_")[1], ("T" if case.startswith("access") else "[rows.csv]")
            given, out = Path(folder) / f"{case}.xlsx", Path(folder) / f"{case}_model.xlsx"
            given.write_bytes(unmasked(book.read_bytes(), folder))
            app = ExcelApplication(sandbox=False)
            app.open_workbook(given)
            app.set_hook("Refresh", numbered)
            app.add_module(MODEL, name="M")
            app.run("Five", kind, source)
            app.workbook.vba_call("SaveAs", [str(out), 51])
            checked = excel.run_vba(CHECK, "Check", (str(out), kind, source), timeout=180.0, line_numbers=False)
            answer = str(checked.value) if checked.ok else f"[{checked.outcome}] {checked.message}"
            measured[case] = answer.replace(folder, FOLDER).split("\t")
            print(case, "\n  ".join(measured[case]))
    text = json.dumps({"cases": measured, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n"
    for leak in ("Users\\", "AppData", "William"):
        assert leak not in text, leak
    OUT.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
