"""What the model loses of a sheet's and a workbook's settings. For each setting live Excel saves a workbook, and says what
it reads there; the model reads the same file and says what it reads, edits a cell that has nothing to do with the
setting, and saves; Excel reads that file. A difference with Excel's first answer is lost on load, in the model's reading,
or lost on save, in Excel's second. The workbook part's and the sheet part's settings elements are kept, too, so a loss no
macro can read still shows.

Rows of panes start from files Excel saved with them shown, tests/fixtures/windows/, as showing Excel under the harness
ended its worker. Hold live.lock. Each workbook is written in tests/fixtures/silent_loss/ with its save path stripped.

    python scripts/measure_silent_loss.py

writes tests/fixtures/silent_loss/silent_loss.json, which tests/test_excel_silent_loss.py replays.
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
sys.path.insert(0, str(ROOT / "src"))

from fixture_workbook import strip_save_path  # noqa: E402
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402
from pyvbaharness import ExcelSession, HarnessConfig  # noqa: E402

OUT = ROOT / "tests/fixtures/silent_loss"

W = "wb.Windows(1)"
T = "ws.Tab"
P = "ws.Protection"
#: Name: what is done before saving; what is read; whether panes need Excel shown.
ROWS: dict[str, tuple[str, list[str], bool]] = {
    "tab_color": (f"{T}.Color = 255", [f"{T}.Color", f"{T}.ColorIndex"], False),
    "tab_color_index": (f"{T}.ColorIndex = 6", [f"{T}.Color", f"{T}.ColorIndex"], False),
    "tab_theme": (f"{T}.ThemeColor = 5\n{T}.TintAndShade = 0.4", [f"{T}.ThemeColor", f"{T}.TintAndShade", f"{T}.ColorIndex"], False),
    "protect_plain": ("ws.Protect", ["ws.ProtectContents", "ws.ProtectDrawingObjects", "ws.ProtectScenarios"], False),
    "protect_contents_only": ("ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False",
                              ["ws.ProtectContents", "ws.ProtectDrawingObjects", "ws.ProtectScenarios"], False),
    "protect_allows": ("ws.Protect AllowFormattingCells:=True, AllowSorting:=True, AllowFiltering:=True, AllowInsertingRows:=True",
                       [f"{P}.AllowFormattingCells", f"{P}.AllowSorting", f"{P}.AllowFiltering", f"{P}.AllowInsertingRows", f"{P}.AllowDeletingRows"], False),
    "protect_password": ('ws.Protect Password:="pw"', ["ws.ProtectContents", "WrongPassword(ws)", "RightPassword(ws)"], False),
    "protect_select": ("ws.EnableSelection = xlUnlockedCells\nws.Protect", ["ws.ProtectContents", "ws.EnableSelection"], False),
    "book_protect": ("wb.Protect Structure:=True", ["wb.ProtectStructure", "wb.ProtectWindows"], False),
    "book_protect_password": ('wb.Protect Password:="pw", Structure:=True', ["wb.ProtectStructure", "BookPassword(wb)"], False),
    "zoom": (f"{W}.Zoom = 150", [f"{W}.Zoom"], False),
    "gridlines": (f"{W}.DisplayGridlines = False", [f"{W}.DisplayGridlines"], False),
    "headings": (f"{W}.DisplayHeadings = False", [f"{W}.DisplayHeadings"], False),
    "zeros": (f"{W}.DisplayZeros = False", [f"{W}.DisplayZeros"], False),
    "formulas": (f"{W}.DisplayFormulas = True", [f"{W}.DisplayFormulas"], False),
    "outline_symbols": (f"{W}.DisplayOutline = False", [f"{W}.DisplayOutline"], False),
    "ruler_tabs": (f"{W}.DisplayWorkbookTabs = False\n{W}.TabRatio = 0.3",
                   [f"{W}.DisplayWorkbookTabs", f"{W}.TabRatio"], False),
    "scroll_bars": (f"{W}.DisplayHorizontalScrollBar = False\n{W}.DisplayVerticalScrollBar = False",
                    [f"{W}.DisplayHorizontalScrollBar", f"{W}.DisplayVerticalScrollBar"], False),
    "view": (f"{W}.View = xlPageBreakPreview", [f"{W}.View", f"{W}.Zoom"], False),
    "right_to_left": (f"ws.DisplayRightToLeft = True", ["ws.DisplayRightToLeft"], False),
    "scroll": ("", [f"{W}.ScrollRow", f"{W}.ScrollColumn"], False),
    "selection": ('ws.Range("C3:D4").Select', [f"{W}.RangeSelection.Address", f"{W}.ActiveCell.Address"], False),
    "selection_areas": ('ws.Range("A1,C3:D4").Select', [f"{W}.RangeSelection.Address", f"{W}.ActiveCell.Address"], False),
    "freeze_rows": ("", [f"{W}.FreezePanes", f"{W}.SplitRow", f"{W}.SplitColumn"], False),
    "freeze_columns": ("", [f"{W}.FreezePanes", f"{W}.SplitRow", f"{W}.SplitColumn"], False),
    "freeze_both": ("", [f"{W}.FreezePanes", f"{W}.SplitRow", f"{W}.SplitColumn"], False),
    "split": ("", [f"{W}.Split", f"{W}.SplitRow", f"{W}.FreezePanes"], False),
    "split_columns": ("", [f"{W}.Split", f"{W}.SplitColumn", f"{W}.FreezePanes"], False),
    "active_sheet": ("wb.Worksheets(2).Activate", ["wb.ActiveSheet.Name"], False),
    "group": ('wb.Worksheets(Array(1, 2)).Select', ["wb.Windows(1).SelectedSheets.Count", "wb.ActiveSheet.Name"], False),
    "scroll_area": ('ws.ScrollArea = "A1:B2"', ["ws.ScrollArea"], False),
    "row_height": ("ws.Rows(3).RowHeight = 30", ["ws.Rows(3).RowHeight", "ws.Rows(4).RowHeight", "ws.StandardHeight"], False),
    "default_height": ("ws.Rows.RowHeight = 25", ["ws.Rows(3).RowHeight", "ws.StandardHeight"], False),
    "column_width": ("ws.Columns(2).ColumnWidth = 20", ["ws.Columns(2).ColumnWidth", "ws.Columns(3).ColumnWidth"], False),
    "standard_width": ("ws.StandardWidth = 12", ["ws.StandardWidth", "ws.Columns(3).ColumnWidth"], False),
    "hidden_row": ("ws.Rows(3).Hidden = True", ["ws.Rows(3).Hidden", "ws.Rows(3).RowHeight"], False),
    "hidden_column": ("ws.Columns(2).Hidden = True", ["ws.Columns(2).Hidden", "ws.Columns(2).ColumnWidth"], False),
    "outline_rows": ('ws.Rows("2:3").Group', ["ws.Rows(2).OutlineLevel", "ws.Rows(4).OutlineLevel"], False),
    "outline_collapsed": ('ws.Rows("2:3").Group\nws.Outline.ShowLevels RowLevels:=1', ["ws.Rows(2).Hidden", "ws.Rows(2).OutlineLevel"], False),
    "outline_summary": ("ws.Outline.SummaryRow = xlAbove\nws.Outline.SummaryColumn = xlLeft", ["ws.Outline.SummaryRow", "ws.Outline.SummaryColumn"], False),
    "outline_columns": ('ws.Columns("B:C").Group', ["ws.Columns(2).OutlineLevel", "ws.Columns(4).OutlineLevel"], False),
    "transition": ("ws.TransitionExpEval = True\nws.TransitionFormEntry = True", ["ws.TransitionExpEval", "ws.TransitionFormEntry"], False),
    "code_name": ("", ["ws.CodeName"], False),
    "date_1904": ("wb.Date1904 = True", ["wb.Date1904"], False),
    "precision": ('ws.Range("A1").NumberFormat = "0.0"\nws.Range("A1").Value = 1.26\nwb.PrecisionAsDisplayed = True', ["wb.PrecisionAsDisplayed", 'ws.Range("A1").Value2'], False),
    "force_full": ("wb.ForceFullCalculation = True", ["wb.ForceFullCalculation"], False),
    "save_link_values": ("wb.SaveLinkValues = False", ["wb.SaveLinkValues"], False),
    "update_links": ("wb.UpdateLinks = xlUpdateLinksNever", ["wb.UpdateLinks"], False),
    "check_compatibility": ("wb.CheckCompatibility = True", ["wb.CheckCompatibility"], False),
    "remove_personal": ("wb.RemovePersonalInformation = True", ["wb.RemovePersonalInformation"], False),
    "drawing_objects": ("wb.DisplayDrawingObjects = xlHide", ["wb.DisplayDrawingObjects"], False),
    "read_only_recommended": ("wb.ReadOnlyRecommended = True", ["wb.ReadOnlyRecommended"], False),
    "default_table_style": ('wb.DefaultTableStyle = "TableStyleLight1"', ["wb.DefaultTableStyle"], False),
    "calc_manual": ("Application.Calculation = xlCalculationManual", [], False),
    "calc_iteration": ("Application.Iteration = True\nApplication.MaxIterations = 50\nApplication.MaxChange = 0.5", [], False),
    "calc_not_before_save": ("Application.Calculation = xlCalculationManual\nApplication.CalculateBeforeSave = False", [], False),
    "calc_semi": ("Application.Calculation = 2", [], False),
}

#: Rows whose file Excel saved already, with its window shown (scripts/measure_windows.py).
FROM_FILE = {"scroll": "scrolled.xlsx", "freeze_rows": "frozen_rows.xlsx", "freeze_columns": "frozen_columns.xlsx",
             "freeze_both": "frozen_both.xlsx", "split": "split_rows.xlsx", "split_columns": "split_columns.xlsx"}
WINDOWS = ROOT / "tests/fixtures/windows"
#: The elements the comparison keeps, by part.
BOOK_ELEMENTS = ("workbookPr", "workbookProtection", "bookViews", "calcPr", "fileSharing")
SHEET_ELEMENTS = ("sheetPr", "sheetViews", "sheetFormatPr", "cols", "sheetProtection", "dimension")

HELP = '''Private Function WrongPassword(ws As Object) As String
    On Error Resume Next
    ws.Unprotect "no"
    WrongPassword = "E" & Err.Number & "," & ws.ProtectContents
End Function
Private Function RightPassword(ws As Object) As String
    On Error Resume Next
    ws.Unprotect "pw"
    RightPassword = "E" & Err.Number & "," & ws.ProtectContents
End Function
Private Function BookPassword(wb As Object) As String
    On Error Resume Next
    wb.Unprotect "no"
    BookPassword = "E" & Err.Number & "," & wb.ProtectStructure
End Function
Private Function Tried(ByVal n As Long, wb As Object, ws As Object) As String
    Dim v As Variant
    On Error GoTo Failed
    Select Case n
{cases}
    End Select
    Tried = CStr(v)
    Exit Function
Failed:
    Tried = "E" & Err.Number
End Function
'''

SAVE = HELP + '''
Function Make(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Dim calc As Long, iter As Boolean, maxI As Long, maxC As Double, before As Boolean, shown As Boolean, w As Double, h As Double, l As Double, t As Double
    Application.DisplayAlerts = False
    calc = Application.Calculation: iter = Application.Iteration: maxI = Application.MaxIterations: maxC = Application.MaxChange
    before = Application.CalculateBeforeSave
    {show}
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    wb.Worksheets.Add After:=ws
    ws.Activate
    ws.Range("A1").Value = 1
    ws.Range("B2").Value = 2
    ws.Range("H20").Locked = False
    {action}
    wb.SaveAs file, 51
    wb.Close False
    Application.Calculation = calc: Application.Iteration = iter: Application.MaxIterations = maxI: Application.MaxChange = maxC
    Application.CalculateBeforeSave = before
    {hide}
    Make = "ok"
End Function
'''

READ = HELP + '''
Function Read1(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Application.DisplayAlerts = False
    {show}
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)
    wb.Activate
    For i = 0 To {last}
        s = s & "|" & Tried(i, wb, ws)
    Next
    wb.Close False
    {hide}
    Read1 = Mid$(s, 2)
End Function
'''

SHOW = ("Application.Visible = True\nApplication.WindowState = xlNormal\nw = Application.Width: h = Application.Height: "
        "l = Application.Left: t = Application.Top\nApplication.ScreenUpdating = True\nApplication.Width = 900: Application.Height = 600")
SHOW_READ = "Application.Visible = True\nApplication.WindowState = xlNormal\nApplication.ScreenUpdating = True"
HIDE = "Application.Width = w: Application.Height = h: Application.Left = l: Application.Top = t\nApplication.Visible = False"
HIDE_READ = "Application.Visible = False"


def cases(reads: list[str]) -> str:
    return "\n".join(f"    Case {index}\n        v = {read}" for index, read in enumerate(reads)) or "    Case -1\n        v = 0"


def make_source(action: str, reads: list[str], shown: bool) -> str:
    return SAVE.format(cases=cases(reads), action=action, show=SHOW if shown else "", hide=HIDE if shown else "")


def read_source(reads: list[str], shown: bool) -> str:
    return READ.format(cases=cases(reads), last=len(reads) - 1, show=SHOW_READ if shown else "", hide=HIDE_READ if shown else "")


def elements(file: Path) -> dict[str, str]:
    """The settings elements of the workbook part and of the first sheet's."""
    with zipfile.ZipFile(file) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        relations = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        first = re.search(r'<sheet\b[^>]*r:id="([^"]+)"', book)
        assert first is not None
        target = re.search(rf'<Relationship\b[^>]*Id="{first.group(1)}"[^>]*Target="([^"]+)"|<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="{first.group(1)}"', relations)
        assert target is not None
        sheet = package.read("xl/" + (target.group(1) or target.group(2)).lstrip("/").removeprefix("xl/")).decode("utf-8")
    found: dict[str, str] = {}
    for name, text, names in (("book", book, BOOK_ELEMENTS), ("sheet", sheet, SHEET_ELEMENTS)):
        for element in names:
            match = re.search(rf"<{element}\b(?:[^>]*/>|.*?</{element}>)", text, re.DOTALL)
            found[f"{name}.{element}"] = match.group(0) if match else ""
    return found


MODEL_READ = '''Function Read1(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)
    wb.Activate
    For i = 0 To {last}
        s = s & "|" & Tried(i, wb, ws)
    Next
    wb.Close False
    Read1 = Mid$(s, 2)
End Function
'''
MODEL_EDIT = '''Sub Edit1(ByVal file As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    wb.Worksheets(1).Range("H20").Value = 7
    wb.Save
    wb.Close False
End Sub
'''


def model_answer(source: str, entry: str, file: Path) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    try:
        return str(app.run(entry, file.as_posix()))
    except VBAUnsupportedError as failure:
        return f"reports: {str(failure)[:80]}"
    except Exception as failure:  # noqa: BLE001 - what the model does with a real file is what is being measured
        return f"crash: {type(failure).__name__}: {str(failure)[:80]}"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_silent_"))
    done = OUT / "silent_loss.json"
    rows: list[dict[str, object]] = json.loads(done.read_text(encoding="utf-8"))["rows"] if done.exists() else []
    only = set(sys.argv[1:])
    have = {row["name"] for row in rows}
    try:
        for name, (action, reads, shown) in ROWS.items():
            if (only and name not in only) or name in have:
                continue
            original = folder / f"{name}.xlsx"
            through = folder / f"{name}_through.xlsx"
            with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                excel.new_document()
                if name in FROM_FILE:
                    shutil.copy(WINDOWS / FROM_FILE[name], original)
                    made = None
                else:
                    made = excel.run_vba(make_source(action, reads, shown), "Make", (str(original),), timeout=60.0)
                if made is not None and not made.ok:
                    print(name, "made failed", made.outcome, made.message[:80], flush=True)
                    rows.append({"name": name, "action": action, "reads": reads, "failed": f"{made.outcome} {made.message[:80]}"})
                    continue
                first = excel.run_vba(read_source(reads, shown), "Read1", (str(original),), timeout=60.0)
            excel_first = str(first.value) if first.ok else f"failed {first.outcome}"
            helpers = HELP.format(cases=cases(reads))
            model_first = model_answer(helpers + MODEL_READ.format(last=len(reads) - 1), "Read1", original)
            shutil.copy(original, through)
            edited = model_answer(MODEL_EDIT, "Edit1", through)
            excel_second = "unsaved"
            kept = elements(through) if not edited.startswith(("reports", "crash")) else {}
            if not edited.startswith(("reports", "crash")):
                with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                    excel.new_document()
                    second = excel.run_vba(read_source(reads, shown), "Read1", (str(through),), timeout=60.0)
                excel_second = str(second.value) if second.ok else f"failed {second.outcome}"
            given = elements(original)
            lost = {key: [given[key], kept.get(key, "")] for key in given if kept and kept.get(key, "") != given[key]}
            verdict = ("load " if model_first != excel_first else "") + ("save " if excel_second != excel_first else "") + ("xml" if lost else "")
            print(f"{name:24} {verdict or 'ok':12} excel={excel_first[:44]} model={model_first[:44]} after={excel_second[:44]}", flush=True)
            fixture = OUT / f"{name}.xlsx"
            shutil.copy(original, fixture)
            strip_save_path(fixture)
            rows.append({"name": name, "action": action, "reads": reads, "excel": excel_first, "model": model_first,
                         "after": excel_second, "edited": edited, "lost": lost})
            (OUT / "silent_loss.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "help": HELP,
                                                              "rows": rows}, indent=1) + "\n", encoding="utf-8")
    finally:
        shutil.rmtree(folder, ignore_errors=True)


if __name__ == "__main__":
    main()
