"""The active tab and the first tab a saved workbook shows, once a macro has moved, added or deleted tabs. Excel makes each
action on a copy of a book it saved, saves it, and says what it reads there; workbookView's activeTab and firstSheet are
kept. firstSheet is where the tab bar is scrolled to, so this needs Excel shown with a sized window: it stays, activating a
tab before it brings it back, and it passes over a hidden tab. Hidden in a session that had shown nothing, Excel wrote
firstSheet equal to activeTab in every row; the tab bar there has no width, and tests/fixtures/windows/ agrees with this
rule, not that one. A firstSheet past the last tab is a file Excel refuses.

Take live.lock: this shows Excel's window, and puts its size back after each row.

    python scripts/measure_active_tab.py

writes tests/fixtures/active_tab/, which tests/test_excel_active_tab.py replays and
test_excel_reads_the_active_tab_this_wrote in tests/test_live_excel_model_gate.py checks.
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
sys.path.insert(0, str(ROOT / "scripts"))
OUT = ROOT / "tests/fixtures/active_tab"
#: Name: which of three sheets is active when Excel saves the book.
BOOKS = {"first_active": 1, "second_active": 2, "third_active": 3}
ACTIONS: dict[str, str] = {
    "edit": 'wb.Sheets(1).Range("A1").Value = 1',
    "move_active_last": "wb.ActiveSheet.Move After:=wb.Sheets(3)",
    "move_active_first": "wb.ActiveSheet.Move Before:=wb.Sheets(1)",
    "delete_first": "wb.Sheets(1).Delete",
    "delete_two_first": "wb.Sheets(1).Delete: wb.Sheets(1).Delete",
    "delete_last": "wb.Sheets(3).Delete",
    "delete_outer": "wb.Sheets(3).Delete: wb.Sheets(1).Delete",
    "delete_active": "wb.ActiveSheet.Delete",
    "hide_first": "wb.Sheets(1).Visible = 0",
    "add_last_then_back": "Set s = wb.ActiveSheet: wb.Sheets.Add After:=wb.Sheets(3): s.Activate",
    "add_first_then_back": "Set s = wb.ActiveSheet: wb.Sheets.Add Before:=wb.Sheets(1): s.Activate",
    "move_other_first_then_back": "Set s = wb.ActiveSheet: wb.Sheets(3).Move Before:=wb.Sheets(1): s.Activate",
    "activate_first": "wb.Sheets(1).Activate",
    "activate_last": "wb.Sheets(3).Activate",
}
MAKE = '''Sub Make(ByVal file As String, ByVal active As Long)
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(-4167)
    wb.Worksheets.Add After:=wb.Worksheets(1)
    wb.Worksheets.Add After:=wb.Worksheets(2)
    wb.Worksheets(active).Activate
    wb.SaveAs file, 51
    wb.Close False
End Sub
'''
ACT = '''Sub Act(ByVal file As String)
    Dim wb As Workbook, s As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    {action}
    wb.Save
    wb.Close False
End Sub
'''
SHOWN_ACT = '''Sub Act(ByVal file As String)
    Dim wb As Workbook, s As Object, was As Boolean, w As Double, h As Double, l As Double, t As Double
    Application.DisplayAlerts = False
    was = Application.Visible
    Application.Visible = True
    Application.ScreenUpdating = True
    Application.WindowState = xlNormal
    w = Application.Width: h = Application.Height: l = Application.Left: t = Application.Top
    Application.Width = 900: Application.Height = 600
    Set wb = Workbooks.Open(file)
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.Width = 700: ActiveWindow.Height = 500
    {action}
    wb.Save
    wb.Close False
    Application.Width = w: Application.Height = h: Application.Left = l: Application.Top = t
    Application.Visible = was
End Sub
'''
READ = '''Function ReadIt(ByVal file As String) As String
    Dim wb As Workbook, s As Object, t As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(file)
    For Each s In wb.Sheets
        t = t & "," & s.Name
    Next
    ReadIt = Mid$(t, 2) & "|" & wb.ActiveSheet.Name & "|" & wb.Windows(1).SelectedSheets.Count
    wb.Close False
End Function
'''


def view(file: Path) -> str:
    """The workbook's first and active tabs as the file has them."""
    with zipfile.ZipFile(file) as package:
        found = re.search(r"<workbookView\b[^>]*>", package.read("xl/workbook.xml").decode("utf-8"))
    return " ".join(sorted(re.findall(r'\b(?:activeTab|firstSheet)="\d+"', found.group(0)))) if found else ""


def main() -> None:
    from fixture_workbook import strip_save_path
    from pyvbaharness import ExcelSession, HarnessConfig

    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_active_"))
    rows: list[dict[str, str]] = []
    try:
        for book, active in BOOKS.items():
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                base = folder / f"{book}.xlsx"
                excel.run_vba(MAKE, "Make", (str(base), active), timeout=45.0)
                shutil.copy(base, OUT / base.name)
                strip_save_path(OUT / base.name)
                for case, action in ACTIONS.items():
                    file = folder / f"{book}_{case}.xlsx"
                    shutil.copy(base, file)
                    acted = excel.run_vba(SHOWN_ACT.format(action=action), "Act", (str(file),), timeout=45.0)
                    if not acted.ok:
                        print(book, case, "failed", acted.message[:80], flush=True)
                        continue
                    read = excel.run_vba(READ, "ReadIt", (str(file),), timeout=45.0)
                    rows.append({"book": book, "case": case, "reads": str(read.value), "view": view(file)})
                    print(book, case, rows[-1]["reads"], rows[-1]["view"], flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    (OUT / "active_tab.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "actions": ACTIONS, "act": ACT, "read": READ, "rows": rows},
                                                    indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
