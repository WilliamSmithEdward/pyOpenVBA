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
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tests/fixtures/chart_sheets"
sys.path.insert(0, str(ROOT / "scripts"))
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
#: What the save gate makes both ways; the cases are the ones the model does, and these follow a chart's cells.
SAVED_ACTIONS: dict[str, str] = {
    "insert_row": "wb.Worksheets(1).Rows(1).Insert",
    "insert_column": "wb.Worksheets(1).Columns(1).Insert",
    "rename_then_insert": 'wb.Worksheets(1).Name = "Renamed": wb.Worksheets(1).Rows(1).Insert',
}
SAVED_REFUSED = {"add_type_chart", "charts_add", "charts_before_1", "charts_before_2", "charts_after_1", "charts_after_2", "charts_after_last", "charts_two",
                 "activated_chart_charts_add", "copy_chart_last", "copy_chart_first", "group_select", "delete_active_worksheet", "rename_sheet_to_chart",
                 "rename_chart_to_sheet"}
#: The same in a workbook of worksheets alone: Sheet1, Sheet2, Sheet3 with Sheet2 active, the active tab's place changing.
PLAIN_ACTIONS: dict[str, str] = {
    "move_active": "wb.ActiveSheet.Move After:=wb.Sheets(3)",
    "delete_before_active": "wb.Sheets(1).Delete",
    "delete_active": "wb.ActiveSheet.Delete",
    "hide_before_active": "wb.Sheets(1).Visible = 0",
    "move_active_first": "wb.ActiveSheet.Move Before:=wb.Sheets(1)",
    "add_before_then_back": "wb.Sheets.Add Before:=wb.Sheets(1): wb.Sheets(3).Activate",
    "move_other_across": "wb.Sheets(3).Move Before:=wb.Sheets(1): wb.Sheets(3).Activate",
    "activate_another": "wb.Sheets(3).Activate",
}
MAKE_PLAIN = '''Sub MakePlain(ByVal file As String)
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(-4167)
    wb.Worksheets.Add After:=wb.Worksheets(1)
    wb.Worksheets.Add After:=wb.Worksheets(2)
    wb.Worksheets(2).Activate
    wb.SaveAs file, 51
    wb.Close False
End Sub
'''
SAVE = '''Sub Act(ByVal file As String)
    Dim wb As Workbook, r As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    {action}
    wb.Save
    wb.Close False
End Sub
'''
READ = '''Private Function Tabs(wb As Object) As String
    Dim s As Object, t As String, c As Object, v As Variant
    On Error Resume Next
    For Each s In wb.Sheets
        t = t & "," & TypeName(s) & ":" & s.Name & ":" & s.Visible
    Next
    t = t & "|" & wb.ActiveSheet.Name & "|" & wb.Windows(1).SelectedSheets.Count & "|names " & wb.Names.Count
    For Each c In wb.Charts
        v = c.SeriesCollection(1).Values
        t = t & "|" & c.Name & " " & c.SeriesCollection(1).Formula & " " & v(1)
    Next
    Tabs = Mid$(t, 2)
End Function
Function ReadIt(ByVal file As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    ReadIt = Tabs(wb)
    wb.Close False
End Function
'''
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


def _view(file: Path) -> str:
    """The book's first and active tabs as the file has them."""
    with zipfile.ZipFile(file) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
    found = re.search(r"<workbookView\b[^>]*>", book)
    return " ".join(sorted(re.findall(r'\b(?:activeTab|firstSheet)="\d+"', found.group(0)))) if found else ""


def _model_acted(action: str, file: Path) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SAVE.format(action=action), name="M")
    try:
        app.run("Act", file.as_posix())
    except VBAUnsupportedError as failure:
        return f"reports: {str(failure)[:80]}"
    return ""


def gate(rows: list[dict[str, str]], folder: Path) -> list[dict[str, str]]:
    """Each action made by Excel and by the model, each saved, each file read by Excel."""
    from pyvbaharness import ExcelSession, HarnessConfig

    from fixture_workbook import strip_save_path

    saved: list[dict[str, str]] = []
    actions = {case: action for case, action in ACTIONS.items() if case not in SAVED_REFUSED} | SAVED_ACTIONS
    for book in (*BOOKS, "plain"):
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            if book == "plain" and not (OUT / "plain.xlsx").exists():
                made = folder / "plain.xlsx"
                excel.run_vba(MAKE_PLAIN, "MakePlain", (str(made),), timeout=45.0)
                shutil.copy(made, OUT / "plain.xlsx")
                strip_save_path(OUT / "plain.xlsx")
            for case, action in (PLAIN_ACTIONS if book == "plain" else actions).items():
                ours, theirs = folder / f"{book}_{case}_model.xlsx", folder / f"{book}_{case}_excel.xlsx"
                shutil.copy(OUT / f"{book}.xlsx", ours)
                shutil.copy(OUT / f"{book}.xlsx", theirs)
                reported = _model_acted(action, ours)
                if reported:
                    print(book, case, reported, flush=True)
                    continue
                made = excel.run_vba(SAVE.format(action=action), "Act", (str(theirs),), timeout=45.0)
                if not made.ok:
                    print(book, case, "excel failed", made.message[:80], flush=True)
                    continue
                first = excel.run_vba(READ, "ReadIt", (str(theirs),), timeout=45.0)
                second = excel.run_vba(READ, "ReadIt", (str(ours),), timeout=45.0)
                row = {"book": book, "case": case, "excel": str(first.value) if first.ok else f"failed {first.message[:80]}",
                       "model": str(second.value) if second.ok else f"failed {second.message[:80]}",
                       "excel_view": _view(theirs), "model_view": _view(ours)}
                saved.append(row)
                if row["excel"] != row["model"] or row["excel_view"] != row["model_view"]:
                    print("DIFFERS", book, case, "\n  excel", row["excel"], row["excel_view"], "\n  model", row["model"], row["model_view"], flush=True)
    return saved


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
        saved = gate(rows, folder)
        done.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "actions": ACTIONS, "rows": rows, "saved": saved}, indent=1) + "\n",
                        encoding="utf-8")
        print(len(saved), "saved", sum(row["excel"] != row["model"] or row["excel_view"] != row["model_view"] for row in saved), "differ", flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    main()
