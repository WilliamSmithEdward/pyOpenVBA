"""Query tables over a database in a workbook Excel saved, opened again: what Excel reads back of each, and what its refresh then does.
The sources are an Access file made for the run through Excel's own ACE OLEDB provider and a CSV file beside it. No network source is
ever named.

Excel makes each book, refreshes it once and saves it. The saved books, with the run's folder written C:\\data, are kept as fixtures in
tests/fixtures/query_refresh_loaded/. Excel then opens each book again, reads the table's settings, refreshes it for more rows and for
fewer, reads the grid after each, and saves it again, and the parts of that save are recorded.

Predictions, written before the run:
- Every setting reads back as it was saved: CommandText, CommandType xlCmdSql, QueryType xlOLEDBQuery (5) or xlODBCQuery (1),
  BackgroundQuery False, RefreshStyle xlInsertDeleteCells, FieldNames True.
- Connection reads back as the provider's fuller string, as it does in the session that made it, without the password.
- The refreshes lay out the rows as in a book never saved (tests/fixtures/query_refresh_layout.json).
- A ListObject over the database has a QueryTable whose ResultRange is the table's range; the table grows and shrinks with it.
- Saving again writes the connection as it was read, with the same id and name, and the queryTable part with a higher nextId.

    python scripts/measure_query_refresh_loaded.py

writes tests/fixtures/query_refresh_loaded.json and the books in tests/fixtures/query_refresh_loaded/.
"""
from __future__ import annotations

import io
import json
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_query_refresh import PROVIDER  # noqa: E402
from measure_query_refresh_layout import MAKE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/query_refresh_loaded.json"
BOOKS = ROOT / "tests/fixtures/query_refresh_loaded"
#: What stands for the run's folder in the fixtures.
FOLDER = "C:\\data"
CSV = "Id,Name\r\n" + "".join(f"{at},n{at}\r\n" for at in range(1, 7))
#: How each book's table is made, given the Access file's path, the folder and the provider; its connection and command.
SOURCES = {
    "access": ('"OLEDB;Provider={provider};Data Source=" & file', "T", "Id"),
    "csv": ('"OLEDB;Provider={provider};Data Source=" & folder & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"""', "[rows.csv]", "Id"),
}
BUILD = r'''Private Function Rows_(ByVal source As String, ByVal n As Long) As String
    Rows_ = "SELECT * FROM " & source & " WHERE Id <= " & n & " ORDER BY Id"
End Function
Function Build(ByVal file As String, ByVal folder As String, ByVal out As String, ByVal kind As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, lo As ListObject, connection As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    connection = {connection}
    ws.Range("A1").Value = "marker"
    ws.Range("B9").Value = "below"
    If kind = "table" Then
        Set lo = ws.ListObjects.Add(xlSrcExternal, Array(connection), , xlYes, ws.Range("B2"))
        Set q = lo.QueryTable
        q.CommandType = xlCmdSql
        q.CommandText = Rows_("{source}", 3)
    Else
        Set q = ws.QueryTables.Add(connection, ws.Range("B2"), Rows_("{source}", 3))
    End If
    q.Refresh False
    wb.SaveAs out, 51
    wb.Close False
    Build = "built"
End Function
'''
PROBE = r'''Private Function Cell(ByVal r As Range) As String
    Dim v As Variant
    v = r.Value
    Cell = TypeName(v) & ":" & CStr(v)
    If r.HasFormula Then Cell = Cell & "=" & r.Formula
    If r.Font.Bold = True Then Cell = Cell & "!b"
End Function
Private Function Grid(ByVal ws As Worksheet) As String
    Dim r As Long, c As Long, s As String
    For r = 1 To 10
        For c = 1 To 5
            s = s & "|" & Cell(ws.Cells(r, c))
        Next
    Next
    Grid = s
End Function
Private Function Place(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal q As QueryTable) As String
    Dim s As String, nm As Name
    On Error Resume Next
    s = "result=" & q.ResultRange.Address
    If Err.Number <> 0 Then s = "result=E" & Err.Number
    Err.Clear
    s = s & " tables=" & ws.ListObjects.Count
    If ws.ListObjects.Count > 0 Then s = s & " " & ws.ListObjects(1).Range.Address & " " & ws.ListObjects(1).Name
    For Each nm In wb.Names
        s = s & " " & nm.Name & "=" & nm.RefersTo & "/" & nm.Visible
    Next
    Place = s & Grid(ws)
End Function
Private Function Settings(ByVal q As QueryTable) As String
    Dim s As String, names As Variant, i As Long
    names = Array("Name", "Connection", "CommandText", "CommandType", "QueryType", "BackgroundQuery", "RefreshStyle", "FieldNames", _
        "RowNumbers", "FillAdjacentFormulas", "PreserveFormatting", "AdjustColumnWidth", "SavePassword", "SaveData", "EnableRefresh", _
        "RefreshOnFileOpen", "EnableEditing", "MaintainConnection", "PreserveColumnInfo", "RefreshPeriod", "Refreshing", "RobustConnect")
    On Error Resume Next
    For i = 0 To UBound(names)
        Err.Clear
        s = s & vbTab & names(i) & "=" & TypeName(CallByName(q, names(i), VbGet)) & ":" & CStr(CallByName(q, names(i), VbGet))
        If Err.Number <> 0 Then s = s & vbTab & names(i) & "=E" & Err.Number
    Next
    Err.Clear
    s = s & vbTab & "Destination=" & q.Destination.Address
    If Err.Number <> 0 Then s = s & "E" & Err.Number
    Err.Clear
    s = s & vbTab & "WorkbookConnection=" & q.WorkbookConnection.Name & "/" & q.WorkbookConnection.Type
    If Err.Number <> 0 Then s = s & "E" & Err.Number
    Err.Clear
    s = s & vbTab & "Parent=" & TypeName(q.Parent)
    Err.Clear
    s = s & vbTab & "ListObject=" & q.ListObject.Name
    If Err.Number <> 0 Then s = s & "E" & Err.Number
    Settings = s
End Function
Function Probe(ByVal path As String, ByVal out As String, ByVal kind As String, ByVal source As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, s As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    If Err.Number <> 0 Then Probe = "open E" & Err.Number & " " & Err.Description: Exit Function
    Set ws = wb.Worksheets(1)
    s = "querytables=" & ws.QueryTables.Count & " listobjects=" & ws.ListObjects.Count & " connections=" & wb.Connections.Count
    If kind = "table" Then
        Set q = ws.ListObjects(1).QueryTable
        s = s & " sourcetype=" & ws.ListObjects(1).SourceType
    Else
        Set q = ws.QueryTables(1)
    End If
    If q Is Nothing Then Probe = s & " no table E" & Err.Number: wb.Close False: Exit Function
    s = s & Settings(q)
    s = s & vbTab & "opened " & Place(wb, ws, q)
    Err.Clear
    q.CommandText = "SELECT * FROM " & source & " WHERE Id <= 5 ORDER BY Id"
    q.Refresh False
    s = s & vbTab & "five E" & Err.Number & " " & Place(wb, ws, q)
    Err.Clear
    q.CommandText = "SELECT * FROM " & source & " WHERE Id <= 2 ORDER BY Id"
    q.Refresh False
    s = s & vbTab & "two E" & Err.Number & " " & Place(wb, ws, q)
    s = s & vbTab & "after" & Settings(q)
    Err.Clear
    wb.SaveAs out, 51
    s = s & vbTab & "saved E" & Err.Number
    wb.Close False
    Probe = s
End Function
'''


def masked_book(data: bytes, folder: str) -> bytes:
    """The book with the run's folder written C:\\data in every part, its entries in their order."""
    spellings = {folder: FOLDER, folder.replace("\\", "/"): FOLDER.replace("\\", "/")}
    found = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(found, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            body = source.read(info)
            for old, new in spellings.items():
                body = body.replace(old.encode("utf-8"), new.encode("utf-8"))
            target.writestr(info, body)
    return found.getvalue()


def parts(data: bytes) -> dict[str, str]:
    """The parts a save wrote that say what the query tables are."""
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return {name: package.read(name).decode("utf-8") for name in package.namelist()
                if "connections" in name or "queryTable" in name or "tables/" in name or name == "xl/workbook.xml"}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    BOOKS.mkdir(exist_ok=True)
    measured: dict[str, object] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_ql_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        (Path(folder) / "rows.csv").write_bytes(CSV.encode("ascii"))
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for source, (connection, table, _) in SOURCES.items():
            for kind in ("query", "table"):
                case = f"{source}_{kind}"
                saved, again = Path(folder) / f"{case}.xlsx", Path(folder) / f"{case}_again.xlsx"
                code = BUILD.replace("{connection}", connection.replace("{provider}", PROVIDER)).replace("{source}", table)
                built = excel.run_vba(code, "Build", (file, folder, str(saved), kind), timeout=120.0, line_numbers=False)
                if not built.ok or built.value != "built":
                    measured[case] = {"built": f"[{built.outcome}] {built.message} {built.value}"}
                    print(case, measured[case])
                    continue
                kept = masked_book(saved.read_bytes(), folder)
                (BOOKS / f"{case}.xlsx").write_bytes(kept)
                probed = excel.run_vba(PROBE, "Probe", (str(saved), str(again), kind, table), timeout=180.0, line_numbers=False)
                answer = str(probed.value) if probed.ok else f"[{probed.outcome}] {probed.message}"
                answer = answer.replace(folder, FOLDER)
                measured[case] = {"probe": answer.split("\t"), "saved": parts(kept),
                                  "again": parts(masked_book(again.read_bytes(), folder)) if again.exists() else None}
                print(case, answer.replace("\t", "\n  "))
    text = json.dumps({"cases": measured, "provider": PROVIDER, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n"
    for leak in ("Users\\", "Users/", "AppData", "William"):
        assert leak not in text, leak
    for book in BOOKS.iterdir():
        assert b"AppData" not in book.read_bytes(), book
    OUT.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
