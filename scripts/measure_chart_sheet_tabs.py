"""Where tabs land in a workbook holding chart sheets: Sheets.Add, Charts.Add, Move, Copy and Delete, each on a fresh open of
a book from tests/fixtures/chart_sheets/, read back as its tabs in order, the active sheet and what the call answered.

Predictions, written before measuring. A: Before and After count every tab, and with neither a new sheet goes before the
active tab, chart or not. B: Before and After count every tab, but with neither a new sheet goes before the active
worksheet, passing over an active chart sheet. A one-sheet book's Charts.Add(After:=) put the chart first when the books
were made, so C1: After is ignored only when it names the last tab; C2: only in a one-tab book.

    python scripts/measure_chart_sheet_tabs.py

writes tests/fixtures/chart_sheets/tabs.json, which tests/test_excel_chart_sheet_tabs.py replays.
"""
from __future__ import annotations

import json
import shutil
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/chart_sheets"
#: The books: middle is Sheet1, C, Two; first is C, Sheet1; last is Sheet1, C; two_charts is C, D, Sheet1; active is C,
#: Sheet1 with C active; names is Sheet1, C, Two with names scoped to both worksheets.
ACTIONS: dict[str, str] = {
    "sheets_add": "Set r = wb.Sheets.Add",
    "worksheets_add": "Set r = wb.Worksheets.Add",
    "add_before_1": "Set r = wb.Sheets.Add(Before:=wb.Sheets(1))",
    "add_before_2": "Set r = wb.Sheets.Add(Before:=wb.Sheets(2))",
    "add_after_1": "Set r = wb.Sheets.Add(After:=wb.Sheets(1))",
    "add_after_2": "Set r = wb.Sheets.Add(After:=wb.Sheets(2))",
    "add_after_last": "Set r = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))",
    "add_two": "Set r = wb.Sheets.Add(Count:=2)",
    "add_after_two": "Set r = wb.Sheets.Add(After:=wb.Sheets(1), Count:=2)",
    "add_type_chart": "Set r = wb.Sheets.Add(Type:=-4109)",
    "activated_chart_add": 'wb.Charts(1).Activate: Set r = wb.Sheets.Add',
    "charts_add": "Set r = wb.Charts.Add",
    "charts_before_1": "Set r = wb.Charts.Add(Before:=wb.Sheets(1))",
    "charts_before_2": "Set r = wb.Charts.Add(Before:=wb.Sheets(2))",
    "charts_after_1": "Set r = wb.Charts.Add(After:=wb.Sheets(1))",
    "charts_after_2": "Set r = wb.Charts.Add(After:=wb.Sheets(2))",
    "charts_after_last": "Set r = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count))",
    "charts_two": "Set r = wb.Charts.Add(Count:=2)",
    "activated_chart_charts_add": "wb.Charts(1).Activate: Set r = wb.Charts.Add",
    "move_sheet_first": "wb.Worksheets(wb.Worksheets.Count).Move Before:=wb.Sheets(1)",
    "move_sheet_last": "wb.Worksheets(1).Move After:=wb.Sheets(wb.Sheets.Count)",
    "move_sheet_after_chart": "wb.Worksheets(1).Move After:=wb.Charts(1)",
    "move_sheet_before_chart": "wb.Worksheets(wb.Worksheets.Count).Move Before:=wb.Charts(1)",
    "move_chart_first": "wb.Charts(wb.Charts.Count).Move Before:=wb.Sheets(1)",
    "move_chart_last": "wb.Charts(1).Move After:=wb.Sheets(wb.Sheets.Count)",
    "copy_sheet_first": "wb.Worksheets(wb.Worksheets.Count).Copy Before:=wb.Sheets(1)",
    "copy_sheet_last": "wb.Worksheets(1).Copy After:=wb.Sheets(wb.Sheets.Count)",
    "copy_chart_last": "wb.Charts(1).Copy After:=wb.Sheets(wb.Sheets.Count)",
    "copy_chart_first": "wb.Charts(1).Copy Before:=wb.Sheets(1)",
    "delete_sheet": "wb.Worksheets(wb.Worksheets.Count).Delete",
    "delete_chart": "wb.Charts(1).Delete",
    "delete_active_worksheet": "wb.Worksheets(1).Activate: wb.Worksheets(1).Delete",
    "hide_active_worksheet": "wb.Worksheets(1).Activate: wb.Worksheets(1).Visible = 0",
    "added_then_named": 'Set r = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): r.Name = "Fresh"',
    "delete_active_chart": "wb.Charts(1).Activate: wb.Charts(1).Delete",
    "rename_sheet": 'wb.Worksheets(1).Name = "Renamed"',
    "rename_chart": 'wb.Charts(1).Name = "Renamed"',
    "rename_sheet_to_chart": 'wb.Worksheets(1).Name = wb.Charts(1).Name',
    "rename_chart_to_sheet": 'wb.Charts(1).Name = wb.Worksheets(1).Name',
    "hide_chart": "wb.Charts(1).Visible = 0",
    "hide_only_visible_sheet": "wb.Worksheets(1).Visible = 0",
    "group_select": "wb.Sheets(Array(1, 2)).Select",
}
BOOKS = ("middle", "first", "last", "two_charts", "active", "names")
PROBE = '''Private Function Tabs(wb As Object) As String
    Dim s As Object, t As String
    For Each s In wb.Sheets
        t = t & "," & TypeName(s) & ":" & s.Name & ":" & s.Index
    Next
    Tabs = Mid$(t, 2) & "|active " & wb.ActiveSheet.Name & "|selected " & wb.Windows(1).SelectedSheets.Count
End Function
Function Go(ByVal file As String) As String
    Dim wb As Workbook, r As Object, e As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    On Error Resume Next
    {action}
    e = Err.Number
    Err.Clear
    Dim made As String
    If Not r Is Nothing Then made = TypeName(r) & ":" & r.Name
    Err.Clear
    Go = "E" & e & "|made " & made & "|" & Tabs(wb)
    wb.Close False
End Function
'''


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_tabs_"))
    done = OUT / "tabs.json"
    rows: list[dict[str, str]] = json.loads(done.read_text(encoding="utf-8"))["rows"] if done.exists() else []
    have = {(row["book"], row["case"]) for row in rows}
    try:
        for book in BOOKS:
            todo = [(case, action) for case, action in ACTIONS.items() if (book, case) not in have]
            if not todo:
                continue
            file = folder / f"{book}.xlsx"
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                for case, action in todo:
                    shutil.copy(OUT / file.name, file)
                    result = excel.run_vba(PROBE.format(action=action), "Go", (str(file),), timeout=45.0, line_numbers=False)
                    answer = str(result.value) if result.ok else f"failed {result.outcome} {result.message[:80]}"
                    rows.append({"book": book, "case": case, "answer": answer})
                    print(book, case, answer, flush=True)
                    if not result.ok:
                        break
            done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "actions": ACTIONS, "rows": rows}, indent=1) + "\n", encoding="utf-8")
    finally:
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    main()
