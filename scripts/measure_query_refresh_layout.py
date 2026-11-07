"""Where a database query table's rows go on refresh, and what happens to the cells around them, from an Access file made for the run through
Excel's own ACE OLEDB provider. No network source is ever named.

Each case adds a query table at B2 over a table of six numbered rows, refreshes it for some of them, and may refresh again for another number,
with marker cells beside and below. It records the result range, the sheet's names and the grid B1 to I14: each cell's value with its
type, and where it has one, its formula and whether it is bold.

Predictions, written before the run:
- xlInsertDeleteCells (the default): a growing result shifts the cells below it down within its columns only, and a shrinking one pulls
  them back up; cells beside it stay.
- xlOverwriteCells: the rows overwrite what is below, and a shrinking result clears its old rows.
- xlInsertEntireRows: whole rows are inserted below as it grows, and not deleted as it shrinks; the old rows are cleared.
- FieldNames False drops the header row; RowNumbers True adds a first column numbered from 1.
- FillAdjacentFormulas fills a formula beside the result down to its last row.
- PreserveFormatting keeps a bold header and first data row bold on the next refresh, and gives new rows the row above's format.

    python scripts/measure_query_refresh_layout.py

writes tests/fixtures/query_refresh_layout.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_query_refresh import PROVIDER  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/query_refresh_layout.json"
MAKE = r'''Function MakeDb(ByVal file As String) As String
    Dim cat As Object, cn As Object, i As Long
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider={provider};Data Source=" & file
    Set cn = cat.ActiveConnection
    cn.Execute "CREATE TABLE T (Id LONG, Name TEXT(10))"
    For i = 1 To 6
        cn.Execute "INSERT INTO T VALUES (" & i & ", 'n" & i & "')"
    Next
    cn.Close
    MakeDb = "made"
End Function
'''
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
Function Go(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, q As QueryTable, lo As ListObject, s As String, r As Long, c As Long, nm As Name
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
{steps}
    If Err.Number <> 0 Then s = "E" & Err.Number & ":" & Err.Description
    Err.Clear
    s = s & "|" & q.ResultRange.Address
    If Err.Number <> 0 Then s = s & "noresult"
    Err.Clear
    On Error GoTo 0
    For Each nm In wb.Names
        s = s & "|" & nm.Name & "=" & nm.RefersTo
    Next
    s = s & "|tables " & ws.ListObjects.Count
    If ws.ListObjects.Count > 0 Then s = s & " " & ws.ListObjects(1).Range.Address & " " & ws.ListObjects(1).Name
    For r = 1 To 14
        For c = 2 To 9
            s = s & vbTab & Cell(ws.Cells(r, c))
        Next
    Next
    wb.Close False
    Go = s
End Function
'''
ADD = 'Set q = ws.QueryTables.Add("OLEDB;Provider={provider};Data Source=" & file, ws.Range("B2"), Rows_({n}))'
MARKERS = 'ws.Range("B12").Value = "below": ws.Range("C13").Value = "below2": ws.Range("F4").Value = "right": ws.Range("A9").Value = "left"'
CASES: dict[str, list[str]] = {}
for style, label in ((1, "insert_cells"), (0, "overwrite"), (2, "entire_rows")):
    for first, second in ((2, 5), (5, 2), (3, 3)):
        CASES[f"{label}_{first}_to_{second}"] = [ADD.format(provider=PROVIDER, n=first), f"q.RefreshStyle = {style}", MARKERS, "q.Refresh False",
                                                 f"q.CommandText = Rows_({second})", "q.Refresh False"]
CASES["first_refresh_over_markers"] = [MARKERS, ADD.format(provider=PROVIDER, n=4), "q.Refresh False"]
CASES["no_field_names"] = [ADD.format(provider=PROVIDER, n=3), "q.FieldNames = False", "q.Refresh False"]
CASES["row_numbers"] = [ADD.format(provider=PROVIDER, n=3), "q.RowNumbers = True", "q.Refresh False"]
CASES["fill_formulas_grow"] = [ADD.format(provider=PROVIDER, n=2), "q.FillAdjacentFormulas = True", "q.Refresh False",
                               'ws.Range("D3").Formula = "=B3*10"', 'ws.Range("D4").Formula = "=B4*10"', "q.CommandText = Rows_(5)", "q.Refresh False"]
CASES["no_fill_formulas_grow"] = [ADD.format(provider=PROVIDER, n=2), "q.FillAdjacentFormulas = False", "q.Refresh False",
                                  'ws.Range("D3").Formula = "=B3*10"', 'ws.Range("D4").Formula = "=B4*10"', "q.CommandText = Rows_(5)", "q.Refresh False"]
CASES["preserve_format"] = [ADD.format(provider=PROVIDER, n=2), "q.PreserveFormatting = True", "q.Refresh False", 'ws.Range("B2:C3").Font.Bold = True',
                            "q.CommandText = Rows_(4)", "q.Refresh False"]
CASES["no_preserve_format"] = [ADD.format(provider=PROVIDER, n=2), "q.PreserveFormatting = False", "q.Refresh False", 'ws.Range("B2:C3").Font.Bold = True',
                               "q.CommandText = Rows_(4)", "q.Refresh False"]
NEAR = 'ws.Range("B6").Value = "under": ws.Range("D3").Value = "beside": ws.Range("C7").Value = "under2"'
for style, label in ((1, "insert_cells"), (0, "overwrite"), (2, "entire_rows")):
    for first, second in ((2, 5), (5, 2)):
        CASES[f"near_{label}_{first}_to_{second}"] = [ADD.format(provider=PROVIDER, n=first), f"q.RefreshStyle = {style}", "q.Refresh False", NEAR,
                                                      f"q.CommandText = Rows_({second})", "q.Refresh False"]
    CASES[f"near_first_{label}"] = [NEAR, ADD.format(provider=PROVIDER, n=5), f"q.RefreshStyle = {style}", "q.Refresh False"]
CASES["defaults"] = [ADD.format(provider=PROVIDER, n=1), 's = "fieldnames=" & q.FieldNames & " rownumbers=" & q.RowNumbers & " fill=" & q.FillAdjacentFormulas & '
                     '" preserve=" & q.PreserveFormatting & " style=" & q.RefreshStyle & " adjust=" & q.AdjustColumnWidth & " bq=" & q.BackgroundQuery & '
                     '" type=" & q.QueryType & " ctype=" & q.CommandType & " savedata=" & q.SaveData']
CASES["zero_rows"] = [ADD.format(provider=PROVIDER, n=3), MARKERS, "q.Refresh False", "q.CommandText = Rows_(0)", "q.Refresh False"]
CASES["list_object"] = ['Set lo = ws.ListObjects.Add(0, Array("OLEDB;Provider={provider};Data Source=" & file), , 1, ws.Range("B2"))'.format(provider=PROVIDER),
                        "Set q = lo.QueryTable", "q.CommandText = Rows_(3)", "q.Refresh False", MARKERS, "q.CommandText = Rows_(5)", "q.Refresh False"]
CASES["list_object_shrink"] = ['Set lo = ws.ListObjects.Add(0, Array("OLEDB;Provider={provider};Data Source=" & file), , 1, ws.Range("B2"))'.format(provider=PROVIDER),
                               "Set q = lo.QueryTable", "q.CommandText = Rows_(5)", "q.Refresh False", MARKERS, "q.CommandText = Rows_(2)", "q.Refresh False"]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    rows: dict[str, list[str]] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_ql_") as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for name, steps in CASES.items():
            code = CASE.replace("{steps}", "\n".join(f"    {step}" for step in steps))
            answer = excel.run_vba(code, "Go", (file,), timeout=120.0, line_numbers=False)
            rows[name] = (str(answer.value).replace(folder, "<folder>").split("\t") if answer.ok else [f"[{answer.outcome}] {answer.message}"])
            print(name, rows[name][0][:160], flush=True)
    OUT.write_text(json.dumps({"provider": PROVIDER, "cases": CASES, "rows": rows, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n",
                   encoding="utf-8")


if __name__ == "__main__":
    main()
