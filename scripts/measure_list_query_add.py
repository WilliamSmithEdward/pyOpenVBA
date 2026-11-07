"""ListObjects.Add over a database: the table it makes before and after a refresh, what its QueryTable holds, how it is refreshed, and
what Excel saves of it. The source is an Access file made for the run through Excel's own ACE OLEDB provider. No network source is ever
named, and nothing here touches encryption, signing or a certificate.

Each case starts a new workbook, runs its steps and records the state at each Info call: the table's range, name, SourceType and
headers, the QueryTable's settings, the sheet's names and query tables, the workbook's connections, and the grid B1 to E10. Cases with a
save write the book to the run's folder and record the parts that describe the table.

Predictions, written before the run:
- Add(xlSrcExternal, Array(connection), , xlYes, B2) makes a table before any refresh: Range $B$2:$B$3, header Column1, named
  Table_ExternalData_1, SourceType 3, with one workbook connection named Connection and no sheet query table.
- Its QueryTable has CommandText "", CommandType xlCmdSql (2), BackgroundQuery True, RefreshStyle xlInsertDeleteCells (1), and Name and
  FieldNames raise 1004.
- A String Source works as the one-element Array does.
- No Destination raises 1004; xlNo for the headers still writes field names.
- CommandType xlCmdTable with CommandText Array("T") gives the same rows as SELECT *.
- ListObject.Refresh refreshes the query, as QueryTable.Refresh False does.
- A second table is Table_ExternalData_2 with ExternalData_2 and connection Connection1.
- Refreshing with no CommandText raises 1004 and leaves the table as it was.
- Refresh with no argument, BackgroundQuery True, leaves the table unchanged until the macro ends.
- A refresh with one more column widens the table, and one with fewer narrows it; the band below shifts in the table's columns only.
- The headers are not bold, and the style is TableStyleMedium2.
- Unlist leaves a plain query table on the sheet; Delete removes the table and its connection.
- The save has the connection with keepAlive="1", a queryTable part with a tableColumnId per field, and the table with
  tableType="queryTable", as in the books of tests/fixtures/query_refresh_loaded/.

Predictions for the second round, written after the first and before it ran (the first's answers are above them in the fixture):
- Unlist and Delete before any refresh drop the connection and the ExternalData_ name, as after one.
- The placeholder column is id 1; a first refresh gives the fields ids from 1 and the columns ids from 1 + the number of fields
  (3 and 4 for two). A later refresh keeps the ids of columns whose field names stay, and gives a new name the next free ids.
- A cell right of or below the two placeholder cells moves right by one as the table is added, as in over_cells; a cell not
  touching them stays.

    python scripts/measure_list_query_add.py

writes tests/fixtures/list_query_add.json.
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
OUT = ROOT / "tests/fixtures/list_query_add.json"
#: What stands for the run's folder in the fixture.
FOLDER = "C:\\data"
CASE = r'''Private Function Cell(ByVal r As Range) As String
    Dim v As Variant
    v = r.Value
    Cell = TypeName(v) & ":" & CStr(v)
    If r.HasFormula Then Cell = Cell & "=" & r.Formula
    If r.Font.Bold = True Then Cell = Cell & "!b"
End Function
Private Function Rows_(ByVal n As Long) As String
    Rows_ = "SELECT * FROM T WHERE Id <= " & n & " ORDER BY Id"
End Function
Private Function Got(ByVal o As Object, ByVal name As String) As String
    On Error Resume Next
    Dim v As Variant
    v = CallByName(o, name, VbGet)
    If Err.Number <> 0 Then
        Got = name & "=E" & Err.Number
    ElseIf IsArray(v) Then
        Got = name & "=Array:" & Join(v, "~")
    Else
        Got = name & "=" & TypeName(v) & ":" & CStr(v)
    End If
End Function
Private Function Info(ByVal wb As Workbook, ByVal ws As Worksheet) As String
    Dim s As String, lo As ListObject, q As QueryTable, nm As Name, c As Object, names As Variant, i As Long, r As Long, k As Long
    On Error Resume Next
    s = "listobjects=" & ws.ListObjects.Count & " querytables=" & ws.QueryTables.Count & " connections=" & wb.Connections.Count
    For Each c In wb.Connections
        s = s & " [" & c.Name & "/" & c.Type & "]"
    Next
    For Each nm In wb.Names
        s = s & " " & nm.Name & "=" & nm.RefersTo & "/" & nm.Visible
    Next
    For Each lo In ws.ListObjects
        s = s & vbTab & "table " & lo.Name & " " & lo.Range.Address & " " & Got(lo, "SourceType") & " " & Got(lo, "ShowHeaders") & _
            " " & Got(lo, "ShowAutoFilter") & " " & Got(lo, "TableStyle") & " " & Got(lo, "DisplayName")
        Err.Clear
        Set q = Nothing
        Set q = lo.QueryTable
        If q Is Nothing Then
            s = s & " query=E" & Err.Number
        Else
            names = Array("Name", "Connection", "CommandText", "CommandType", "QueryType", "BackgroundQuery", "RefreshStyle", "FieldNames", _
                "RowNumbers", "FillAdjacentFormulas", "PreserveFormatting", "AdjustColumnWidth", "SavePassword", "SaveData", "EnableRefresh", _
                "RefreshOnFileOpen", "EnableEditing", "MaintainConnection", "PreserveColumnInfo", "RefreshPeriod", "Refreshing")
            For i = 0 To UBound(names)
                s = s & " " & Got(q, names(i))
            Next
            Err.Clear
            s = s & " result=" & q.ResultRange.Address
            If Err.Number <> 0 Then s = s & "E" & Err.Number
            Err.Clear
            s = s & " destination=" & q.Destination.Address
            If Err.Number <> 0 Then s = s & "E" & Err.Number
            Err.Clear
            s = s & " connection=" & q.WorkbookConnection.Name
            If Err.Number <> 0 Then s = s & "E" & Err.Number
        End If
    Next
    For Each q In ws.QueryTables
        s = s & vbTab & "sheet query " & Got(q, "Name") & " " & Got(q, "CommandText")
        Err.Clear
        s = s & " result=" & q.ResultRange.Address
        If Err.Number <> 0 Then s = s & "E" & Err.Number
    Next
    s = s & vbTab
    For r = 1 To 10
        For k = 2 To 5
            s = s & "|" & Cell(ws.Cells(r, k))
        Next
    Next
    Err.Clear
    Info = s
End Function
Function Go(ByVal file As String, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, lo As ListObject, lo2 As ListObject, s As String, connection As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    connection = "OLEDB;Provider={provider};Data Source=" & file
    On Error Resume Next
{steps}
    If Err.Number <> 0 Then s = s & vbTab & "E" & Err.Number & ":" & Err.Description
    Err.Clear
    wb.Close False
    Go = s
End Function
'''
ADD = 'Set lo = ws.ListObjects.Add(xlSrcExternal, Array(connection), , xlYes, ws.Range("B2"))'
STEP = 'Err.Clear: {step}: s = s & vbTab & "{label} E" & Err.Number & " " & Info(wb, ws)'
MARKERS = 'ws.Range("B9").Value = "below": ws.Range("E3").Value = "right"'


def steps(*pairs: tuple[str, str]) -> list[str]:
    """Each step run, then the state after it under its label."""
    return [STEP.format(step=step, label=label) for label, step in pairs]


CASES: dict[str, list[str]] = {
    "add": steps(("added", ADD), ("query", "Set q = lo.QueryTable"), ("sql", "q.CommandText = Rows_(3)"), ("refreshed", "q.Refresh False")),
    "add_string": steps(("added", ADD.replace("Array(connection)", "connection")), ("query", "Set q = lo.QueryTable"),
                        ("refreshed", "q.CommandText = Rows_(2): q.Refresh False")),
    "add_no_destination": steps(("added", 'Set lo = ws.ListObjects.Add(xlSrcExternal, Array(connection), , xlYes)')),
    "add_no_headers": steps(("added", ADD.replace("xlYes", "xlNo")), ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
    "add_guess_headers": steps(("added", ADD.replace("xlYes", "")), ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
    "command_table": steps(("added", ADD), ("table", 'Set q = lo.QueryTable: q.CommandType = xlCmdTable: q.CommandText = Array("T")'),
                           ("refreshed", "q.Refresh False")),
    "list_refresh": steps(("added", ADD), ("sql", "lo.QueryTable.CommandText = Rows_(3)"), ("refreshed", "lo.Refresh"),
                          ("again", "lo.QueryTable.CommandText = Rows_(1): lo.Refresh")),
    "two_tables": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False"),
                        ("second", 'Set lo2 = ws.ListObjects.Add(xlSrcExternal, Array(connection), , xlYes, ws.Range("E2"))'),
                        ("refreshed", "lo2.QueryTable.CommandText = Rows_(3): lo2.QueryTable.Refresh False")),
    "refresh_no_command": steps(("added", ADD), ("refreshed", "lo.QueryTable.Refresh False")),
    "refresh_background": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False"),
                                ("background", "lo.QueryTable.CommandText = Rows_(4): lo.QueryTable.Refresh")),
    "columns_change": steps(("added", ADD), ("one", 'lo.QueryTable.CommandText = "SELECT Id FROM T WHERE Id <= 3 ORDER BY Id": lo.QueryTable.Refresh False'),
                            ("markers", MARKERS),
                            ("two", "lo.QueryTable.CommandText = Rows_(3): lo.QueryTable.Refresh False"),
                            ("back", 'lo.QueryTable.CommandText = "SELECT Name FROM T WHERE Id <= 3 ORDER BY Id": lo.QueryTable.Refresh False')),
    "grow_shrink_markers": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(3): lo.QueryTable.Refresh False"), ("markers", MARKERS),
                                 ("grow", "lo.QueryTable.CommandText = Rows_(5): lo.QueryTable.Refresh False"),
                                 ("shrink", "lo.QueryTable.CommandText = Rows_(1): lo.QueryTable.Refresh False")),
    "over_cells": steps(("filled", 'ws.Range("B2:C4").Value = "old"'), ("added", ADD),
                        ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
    "unlist": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False"), ("unlisted", "lo.Unlist")),
    "delete": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False"), ("deleted", "lo.Delete")),
    "save": steps(("added", ADD + ": lo.QueryTable.CommandText = Rows_(3): lo.QueryTable.Refresh False"), ("saved", 'wb.SaveAs out & ".xlsx", 51')),
    "save_unrefreshed": steps(("added", ADD), ("saved", 'wb.SaveAs out & ".xlsx", 51')),
    "unlist_unrefreshed": steps(("added", ADD), ("unlisted", "lo.Unlist")),
    "delete_unrefreshed": steps(("added", ADD), ("deleted", "lo.Delete")),
    "beside_right": steps(("filled", 'ws.Range("C3").Value = "x"'), ("added", ADD),
                          ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
    "beside_below": steps(("filled", 'ws.Range("B4").Value = "x"'), ("added", ADD),
                          ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
    "beside_apart": steps(("filled", 'ws.Range("D2").Value = "x"'), ("added", ADD),
                          ("refreshed", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False")),
}
#: Copies saved after each refresh, numbered, for the ids a save gives the table's columns and the query's fields.
COPY = 'wb.SaveCopyAs out & "{n}.xlsx"'
CASES["ids"] = steps(
    ("added", ADD + ": " + COPY.format(n=0)),
    ("one", 'lo.QueryTable.CommandText = "SELECT Id FROM T WHERE Id <= 2 ORDER BY Id": lo.QueryTable.Refresh False: ' + COPY.format(n=1)),
    ("two", "lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False: " + COPY.format(n=2)),
    ("more", "lo.QueryTable.CommandText = Rows_(4): lo.QueryTable.Refresh False: " + COPY.format(n=3)),
    ("renamed", 'lo.QueryTable.CommandText = "SELECT Name FROM T WHERE Id <= 2 ORDER BY Id": lo.QueryTable.Refresh False: ' + COPY.format(n=4)),
    ("reordered", 'lo.QueryTable.CommandText = "SELECT Name, Id FROM T WHERE Id <= 2 ORDER BY Id": lo.QueryTable.Refresh False: '
     + COPY.format(n=5)),
    ("three", 'lo.QueryTable.CommandText = "SELECT Id, Name, Id * 2 AS D FROM T WHERE Id <= 2 ORDER BY Id": lo.QueryTable.Refresh False: '
     + COPY.format(n=6)))
CASES["first_one"] = steps(("added", ADD), ("one", 'lo.QueryTable.CommandText = "SELECT Id FROM T WHERE Id <= 2 ORDER BY Id": '
                                                   'lo.QueryTable.Refresh False: ' + COPY.format(n=1)))
CASES["first_three"] = steps(("added", ADD), ("three", 'lo.QueryTable.CommandText = "SELECT Id, Name, Id * 2 AS D FROM T WHERE Id <= 2 ORDER BY Id": '
                                                       'lo.QueryTable.Refresh False: ' + COPY.format(n=1)))


def parts(data: bytes, folder: str) -> dict[str, str]:
    """The parts a save wrote that say what the table is, with the run's folder written C:\\data."""
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        found = {name: package.read(name).decode("utf-8") for name in package.namelist()
                 if "connections" in name or "queryTable" in name or "tables/" in name or name in ("xl/workbook.xml", "[Content_Types].xml")
                 or name.startswith("xl/worksheets/")}
    return {name: body.replace(folder, FOLDER) for name, body in found.items()}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    rows: dict[str, list[str]] = {}
    saves: dict[str, dict[str, dict[str, str]]] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_la_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for name, case_steps in CASES.items():
            out = Path(folder) / f"{name}~"
            code = CASE.replace("{provider}", PROVIDER).replace("{steps}", "\n".join(f"    {step}" for step in case_steps))
            answer = excel.run_vba(code, "Go", (file, str(out)), timeout=120.0, line_numbers=False)
            rows[name] = (str(answer.value).replace(folder, FOLDER).split("\t") if answer.ok else [f"[{answer.outcome}] {answer.message}"])
            for saved in sorted(Path(folder).glob(f"{name}~*.xlsx")):
                saves.setdefault(name, {})[saved.stem.partition("~")[2] or "book"] = parts(saved.read_bytes(), folder)
            print(name, " / ".join(part[:120] for part in rows[name] if not part.startswith("|")), flush=True)
    OUT.write_text(json.dumps({"provider": PROVIDER, "cases": CASES, "rows": rows, "saves": saves,
                               "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
