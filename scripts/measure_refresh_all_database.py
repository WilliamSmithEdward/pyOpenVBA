"""Workbook.RefreshAll over query tables on a database, a sheet's and a ListObject's: which it refreshes and when. The source is an
Access file made for the run through Excel's own ACE OLEDB provider. No network source is ever named, and nothing here touches
encryption, signing or a certificate.

A source that has gone is not measured here: Excel puts up the provider's "OLE DB Initialization Information" prompt, which a macro
cannot answer (the first run's gone_source case, its file deleted under it).

Each case starts a new workbook of two sheets, runs its steps and records, at each step, every table's result range and whether it
is refreshing, and the grids of both sheets' B2:D8.

Predictions, written before the run:
- RefreshAll refreshes a sheet's table and a ListObject's that a refresh has run on, each with its command as it now is, in the
  macro, BackgroundQuery False; one never refreshed is refreshed too, where it has a command.
- A ListObject never refreshed with no command is passed over without an error.
- A table with BackgroundQuery True is left as it was until the macro ends.
- EnableRefresh False, and a connection with RefreshWithRefreshAll False, are passed over.
- (First run only) A source that has gone raises error 1004 out of RefreshAll, and the tables before it are refreshed. It put up a
  prompt instead, and the case was taken out.

    python scripts/measure_refresh_all_database.py

writes tests/fixtures/refresh_all_database.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_query_refresh import PROVIDER  # noqa: E402
from measure_query_refresh_layout import MAKE  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/refresh_all_database.json"
FOLDER = "C:\\data"
CASE = r'''Private Function Rows_(ByVal n As Long) As String
    Rows_ = "SELECT * FROM T WHERE Id <= " & n & " ORDER BY Id"
End Function
Private Function Grid(ByVal ws As Worksheet) As String
    Dim r As Long, c As Long, s As String, v As Variant
    For r = 2 To 8
        For c = 2 To 4
            v = ws.Cells(r, c).Value
            s = s & "|" & TypeName(v) & ":" & CStr(v)
        Next
    Next
    Grid = s
End Function
Private Function State(ByVal wb As Workbook) As String
    Dim s As String, ws As Worksheet, q As QueryTable, lo As ListObject
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each q In ws.QueryTables
            s = s & " " & ws.Name & ":" & q.Name & "=" & q.ResultRange.Address & "/" & q.Refreshing
        Next
        For Each lo In ws.ListObjects
            s = s & " " & ws.Name & ":" & lo.Name & "=" & lo.Range.Address & "/" & lo.QueryTable.Refreshing
        Next
        s = s & " " & ws.Name & Grid(ws)
    Next
    Err.Clear
    State = s
End Function
Function Go(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, q As QueryTable, lo As ListObject, s As String, connection As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set ws2 = wb.Worksheets.Add(After:=ws)
    connection = "OLEDB;Provider={provider};Data Source=" & file
    On Error Resume Next
{steps}
    If Err.Number <> 0 Then s = s & vbTab & "E" & Err.Number & ":" & Err.Description
    Err.Clear
    wb.Close False
    Go = s
End Function
'''
STEP = 'Err.Clear: {step}: s = s & vbTab & "{label} E" & Err.Number & State(wb)'
SHEET = 'Set q = ws.QueryTables.Add(connection, ws.Range("B2"), Rows_({n}))'
LIST = 'Set lo = ws2.ListObjects.Add(xlSrcExternal, Array(connection), , xlYes, ws2.Range("B2"))'


def steps(*pairs: tuple[str, str]) -> list[str]:
    return [STEP.format(step=step, label=label) for label, step in pairs]


CASES: dict[str, list[str]] = {
    "sheet_refreshed": steps(("made", SHEET.format(n=2) + ": q.Refresh False"), ("command", "q.CommandText = Rows_(4)"), ("all", "wb.RefreshAll")),
    "sheet_never": steps(("made", SHEET.format(n=3)), ("all", "wb.RefreshAll")),
    "list_refreshed": steps(("made", LIST + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False"),
                            ("command", "lo.QueryTable.CommandText = Rows_(5)"), ("all", "wb.RefreshAll")),
    "list_never_command": steps(("made", LIST + ": lo.QueryTable.CommandText = Rows_(3)"), ("all", "wb.RefreshAll")),
    "list_never_bare": steps(("made", LIST), ("all", "wb.RefreshAll")),
    "both": steps(("made", SHEET.format(n=2) + ": q.Refresh False: " + LIST + ": lo.QueryTable.CommandText = Rows_(1): lo.QueryTable.Refresh False"),
                  ("command", "q.CommandText = Rows_(3): lo.QueryTable.CommandText = Rows_(4)"), ("all", "wb.RefreshAll")),
    "sheet_background": steps(("made", SHEET.format(n=2) + ": q.Refresh False: q.BackgroundQuery = True: q.CommandText = Rows_(4)"),
                              ("all", "wb.RefreshAll"), ("after", "DoEvents")),
    "list_background": steps(("made", LIST + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False: lo.QueryTable.BackgroundQuery = True"),
                             ("command", "lo.QueryTable.CommandText = Rows_(4)"), ("all", "wb.RefreshAll")),
    "sheet_disabled": steps(("made", SHEET.format(n=2) + ": q.Refresh False: q.CommandText = Rows_(4): q.EnableRefresh = False"),
                            ("all", "wb.RefreshAll")),
    "list_excluded": steps(("made", LIST + ": lo.QueryTable.CommandText = Rows_(2): lo.QueryTable.Refresh False: lo.QueryTable.CommandText = Rows_(4)"),
                           ("excluded", "wb.Connections(1).RefreshWithRefreshAll = False"), ("all", "wb.RefreshAll")),
}


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    rows: dict[str, list[str]] = {}
    with tempfile.TemporaryDirectory(prefix="pyopenvba_ra_", ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
        excel.new_document()
        file = str(Path(folder) / "rows.accdb")
        made = excel.run_vba(MAKE.replace("{provider}", PROVIDER), "MakeDb", (file,), timeout=120.0, line_numbers=False)
        assert made.ok, made.message
        for name, case_steps in CASES.items():
            code = CASE.replace("{steps}", "\n".join(f"    {step}" for step in case_steps)).replace("{provider}", PROVIDER)
            answer = excel.run_vba(code, "Go", (file,), timeout=120.0, line_numbers=False)
            text = str(answer.value).replace(folder, FOLDER)
            rows[name] = text.split("\t") if answer.ok else [f"[{answer.outcome}] {answer.message}"]
            print(name, " / ".join(part[:140] for part in rows[name]), flush=True)
    text = json.dumps({"provider": PROVIDER, "cases": CASES, "rows": rows, "measured_at": datetime.now(timezone.utc).isoformat()}, indent=1) + "\n"
    for leak in ("Users\\", "AppData", "William"):
        assert leak not in text, leak
    OUT.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
