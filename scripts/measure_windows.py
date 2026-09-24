"""Windows: what ActiveWindow answers, and what a file keeps of a sheet's view.

Each workbook below is made in live Excel from VBA -- cells selected and
activated, panes frozen or split, the window scrolled and zoomed,
gridlines and headings turned off, rows and columns inserted and deleted
round the panes, sheets added, copied, moved and deleted -- and the
window is read back through the object model; then the workbook is saved
as xlsx and the sheetViews of each sheet and the workbook's bookViews
are read out of the package. Excel is shown and each window is 700 by
500 points: hidden, Excel lays out no panes. Last, three workbooks are
added, activated and closed, with the screen updating and without, and
the order Windows lists their windows in is read.

    python scripts/measure_windows.py

writes tests/fixtures/windows/ -- one .xlsx per workbook and
windows.json -- which tests/test_excel_windows.py replays.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "windows"

#: What is read of the active window after each workbook is made: each expression, or the error it raises.
READS: list[str] = [
    "w.Caption", "w.FreezePanes", "w.SplitRow", "w.SplitColumn", "w.ScrollRow", "w.ScrollColumn", "w.Zoom",
    "w.DisplayGridlines", "w.DisplayHeadings", "w.DisplayZeros", "w.DisplayFormulas", "w.View", "w.ActiveSheet.Name",
    "w.ActiveCell.Address", "w.Selection.Address", "w.RangeSelection.Address", "w.Split", "w.WindowState",
    "w.Visible", "w.Index", "TypeName(w)", "TypeName(w.Parent)", "w.DisplayWorkbookTabs",
    "w.DisplayHorizontalScrollBar", "w.DisplayVerticalScrollBar", "w.TabRatio", "w.SplitVertical",
    "w.SplitHorizontal", "w.Panes.Count",
]

READER = "\n".join(
    f"""Private Function R{index}(w As Object) As String
    On Error GoTo Bad
    R{index} = CStr({expression})
    Exit Function
Bad:
    R{index} = "!" & Err.Number
End Function
""" for index, expression in enumerate(READS)) + """
Private Function Seen(w As Object) As String
    Seen = """ + ' & "|" & '.join(f"R{index}(w)" for index in range(len(READS))) + """
End Function
"""

#: Each workbook: its name, and the VBA that sets up its window, with wb the workbook and ws its first sheet.
BOOKS: list[tuple[str, str]] = [
    ("fresh", ""),
    ("selected", 'ws.Range("C5").Select'),
    ("selected_block", 'ws.Range("B2:D4").Select'),
    ("frozen_both", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True'),
    ("frozen_rows", 'ws.Range("A3").Select\nActiveWindow.FreezePanes = True'),
    ("frozen_columns", 'ws.Range("C1").Select\nActiveWindow.FreezePanes = True'),
    ("frozen_at_block", 'ws.Range("C4:E6").Select\nActiveWindow.FreezePanes = True'),
    ("frozen_then_selected", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nws.Range("D9").Select'),
    ("frozen_scrolled", 'ActiveWindow.ScrollRow = 10\nActiveWindow.ScrollColumn = 3\nws.Range("D12").Select\n'
                        'ActiveWindow.FreezePanes = True'),
    ("unfrozen", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.FreezePanes = False'),
    ("split_rows", "ActiveWindow.SplitRow = 3"),
    ("split_columns", "ActiveWindow.SplitColumn = 2"),
    ("split_then_frozen", "ActiveWindow.SplitRow = 3\nActiveWindow.SplitColumn = 2\nActiveWindow.FreezePanes = True"),
    ("scrolled", "ActiveWindow.ScrollRow = 20\nActiveWindow.ScrollColumn = 5"),
    ("zoomed", "ActiveWindow.Zoom = 85"),
    ("zoomed_far", "ActiveWindow.Zoom = 250"),
    ("no_gridlines", "ActiveWindow.DisplayGridlines = False"),
    ("no_headings", "ActiveWindow.DisplayHeadings = False"),
    ("no_zeros", "ActiveWindow.DisplayZeros = False"),
    ("formulas", "ActiveWindow.DisplayFormulas = True"),
    ("page_break_view", "ActiveWindow.View = xlPageBreakPreview"),
    ("page_layout_view", "ActiveWindow.View = xlPageLayoutView"),
    ("no_tabs", "ActiveWindow.DisplayWorkbookTabs = False\nActiveWindow.TabRatio = 0.4"),
    ("no_scroll_bars", "ActiveWindow.DisplayHorizontalScrollBar = False\n"
                       "ActiveWindow.DisplayVerticalScrollBar = False"),
    ("second_sheet", 'wb.Worksheets.Add After:=ws\nwb.Worksheets(2).Range("E7").Select\n'
                     "ActiveWindow.DisplayGridlines = False\nActiveWindow.Zoom = 120"),
    ("second_sheet_back", 'wb.Worksheets.Add After:=ws\nActiveWindow.Zoom = 120\nActiveWindow.FreezePanes = False\n'
                          'wb.Worksheets(2).Range("B2").Select\nActiveWindow.FreezePanes = True\nws.Activate\n'
                          'ws.Range("C3").Select'),
    ("caption", 'ActiveWindow.Caption = "Report"'),
    # What the macro recorder writes for Freeze Top Row and Freeze First Column.
    ("recorded_top_row", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\nActiveWindow.FreezePanes = True"),
    ("recorded_first_column", "With ActiveWindow\n.SplitColumn = 1\n.SplitRow = 0\nEnd With\n"
                              "ActiveWindow.FreezePanes = True"),
    ("recorded_top_row_selected", 'ws.Range("C5").Select\nWith ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\n'
                                  "End With\nActiveWindow.FreezePanes = True"),
    ("recorded_both", "With ActiveWindow\n.SplitColumn = 2\n.SplitRow = 3\nEnd With\nActiveWindow.FreezePanes = True"),
    ("recorded_then_selected", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                               'ActiveWindow.FreezePanes = True\nws.Range("D9").Select'),
    ("recorded_unfrozen", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                          "ActiveWindow.FreezePanes = True\nActiveWindow.FreezePanes = False"),
    ("split_zero", "ActiveWindow.SplitRow = 0\nActiveWindow.SplitColumn = 0"),
    ("frozen_split_row_moved", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.SplitRow = 4'),
    ("split_removed", "ActiveWindow.SplitRow = 3\nActiveWindow.Split = False"),
    ("frozen_split_off", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.Split = False'),
    ("frozen_then_scrolled", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.ScrollRow = 10\n'
                             "ActiveWindow.ScrollColumn = 4"),
    ("unfrozen_scrolled", 'ActiveWindow.ScrollRow = 10\nActiveWindow.ScrollColumn = 3\nws.Range("D12").Select\n'
                          "ActiveWindow.FreezePanes = True\nActiveWindow.FreezePanes = False"),
    ("freeze_twice", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nws.Range("D4").Select\n'
                     "ActiveWindow.FreezePanes = True"),
    ("break_zoomed", "ActiveWindow.View = xlPageBreakPreview\nActiveWindow.Zoom = 70"),
    ("break_zoomed_back", "ActiveWindow.View = xlPageBreakPreview\nActiveWindow.Zoom = 70\n"
                          "ActiveWindow.View = xlNormalView"),
    ("layout_zoomed", "ActiveWindow.View = xlPageLayoutView\nActiveWindow.Zoom = 80"),
    ("zoomed_then_break", "ActiveWindow.Zoom = 85\nActiveWindow.View = xlPageBreakPreview"),
    ("zoom_too_small", "ActiveWindow.Zoom = 5"),
    ("zoom_fraction", "ActiveWindow.Zoom = 85.6"),
    ("scroll_zero", "ActiveWindow.ScrollRow = 0"),
    ("view_unknown", "ActiveWindow.View = 7"),
    ("multi_area", 'ws.Range("B2:C3,E5").Select'),
    ("activate_inside", 'ws.Range("B2:D4").Select\nws.Range("C3").Activate'),
    ("activate_outside", 'ws.Range("B2:D4").Select\nws.Range("F8").Activate'),
    ("select_far", 'ws.Range("Z100").Select'),
    ("goto_far", 'Application.Goto ws.Range("Z100"), True'),
    ("select_other_sheet", 'wb.Worksheets.Add After:=ws\nws.Range("C3").Select'),
    ("copied", 'ws.Range("C5").Select\nActiveWindow.Zoom = 85\nws.Range("B2").Select\nActiveWindow.FreezePanes = True\n'
               "ws.Copy After:=ws"),
    *[(f"frozen_zoomed_{zoom}", f'ActiveWindow.Zoom = {zoom}\nws.Range("D5").Select\nActiveWindow.FreezePanes = True')
      for zoom in (33, 47, 75, 90, 110, 115, 150, 175, 200, 399)],
    ("frozen_select_a1", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nws.Range("A1").Select'),
    ("recorded_split_off", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                           "ActiveWindow.FreezePanes = True\nActiveWindow.Split = False"),
    ("recorded_scrolled", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                          "ActiveWindow.FreezePanes = True\nActiveWindow.ScrollRow = 10"),
    ("scrolled_recorded", "ActiveWindow.ScrollRow = 10\nWith ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                          "ActiveWindow.FreezePanes = True"),
    ("goto_near", 'Application.Goto ws.Range("C5")'),
    ("goto_near_scroll", 'Application.Goto ws.Range("C5"), True'),
    ("goto_far_no_scroll", 'Application.Goto ws.Range("Z100")'),
    ("goto_other_sheet", 'wb.Worksheets.Add After:=ws\nApplication.Goto ws.Range("C5"), True'),
    ("goto_text", 'Application.Goto "R3C2"'),
    ("goto_frozen_scroll", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\n'
                           'Application.Goto ws.Range("Z100"), True'),
    ("activate_block_inside", 'ws.Range("B2:D4").Select\nws.Range("C3:D4").Activate'),
    ("activate_block_outside", 'ws.Range("B2:D4").Select\nws.Range("F8:G9").Activate'),
    ("activate_other_sheet", 'wb.Worksheets.Add After:=ws\nws.Range("C3").Activate'),
    ("select_other_book", 'Set other = Workbooks.Add(xlWBATWorksheet)\nws.Range("C3").Select\nother.Close False\n'
                          "wb.Activate"),
    ("zoom_400", "ActiveWindow.Zoom = 400"),
    ("zoom_401", "ActiveWindow.Zoom = 401"),
    ("zoom_10", "ActiveWindow.Zoom = 10"),
    ("scroll_last", "ActiveWindow.ScrollRow = 1048576"),
    ("select_whole_column", 'ws.Columns("C").Select'),
    ("select_whole_row", "ws.Rows(3).Select"),
    ("select_all", "ws.Cells.Select"),
    ("delete_active_sheet", "wb.Worksheets.Add After:=ws\nApplication.DisplayAlerts = False\nwb.Worksheets(2).Delete"),
    ("move_active_sheet", "wb.Worksheets.Add After:=ws\nwb.Worksheets(2).Move Before:=ws"),
    ("activate_block_overlapping", 'ws.Range("B2:D4").Select\nws.Range("D4:E5").Activate'),
    ("activate_block_first_outside", 'ws.Range("B2:D4").Select\nws.Range("A1:C3").Activate'),
    ("recorded_top_row_scrolled_across", "With ActiveWindow\n.SplitColumn = 0\n.SplitRow = 1\nEnd With\n"
                                         "ActiveWindow.FreezePanes = True\nActiveWindow.ScrollColumn = 5"),
    ("recorded_first_column_scrolled_down", "With ActiveWindow\n.SplitColumn = 1\n.SplitRow = 0\nEnd With\n"
                                            "ActiveWindow.FreezePanes = True\nActiveWindow.ScrollRow = 10"),
    ("frozen_scrolled_above", 'ws.Range("B3").Select\nActiveWindow.FreezePanes = True\nActiveWindow.ScrollRow = 2'),
    ("frozen_scrolled_unfrozen", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.ScrollRow = 10\n'
                                 "ActiveWindow.ScrollColumn = 4\nActiveWindow.FreezePanes = False"),
    ("freeze_out_of_view", 'ActiveWindow.ScrollRow = 10\nws.Range("B2").Select\nActiveWindow.FreezePanes = True'),
    ("caption_empty", 'ActiveWindow.Caption = ""'),
    ("tab_ratio_big", "ActiveWindow.TabRatio = 1.5"),
    ("split_negative", "ActiveWindow.SplitRow = -1"),
    ("view_same", "ActiveWindow.View = xlNormalView"),
    ("zoom_100", "ActiveWindow.Zoom = 100"),
    ("frozen_split_true", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\nActiveWindow.Split = True'),
    ("split_true", 'ws.Range("C5").Select\nActiveWindow.Split = True'),
    # Rows and columns inserted and deleted round the selection and the panes.
    ("selected_row_deleted_above", 'ws.Range("C5").Select\nws.Rows(2).Delete'),
    ("selected_row_inserted_above", 'ws.Range("C5").Select\nws.Rows(2).Insert'),
    ("selected_column_deleted_inside", 'ws.Range("B2:D4").Select\nws.Columns("C").Delete'),
    ("selected_row_deleted_itself", 'ws.Range("C5").Select\nws.Rows(5).Delete'),
    ("frozen_row_inserted_above", 'ws.Range("B3").Select\nActiveWindow.FreezePanes = True\nws.Rows(1).Insert'),
    ("frozen_column_deleted_left", 'ws.Range("C3").Select\nActiveWindow.FreezePanes = True\nws.Columns("A").Delete'),
    ("scrolled_row_deleted_above", 'ActiveWindow.ScrollRow = 10\nws.Rows(2).Delete'),
    ("frozen_cells_shifted_down", 'ws.Range("B3").Select\nActiveWindow.FreezePanes = True\n'
                                  'ws.Range("A1:C1").Insert Shift:=xlDown'),
    ("frozen_scrolled_row_inserted_between", 'ws.Range("B2").Select\nActiveWindow.FreezePanes = True\n'
                                             "ActiveWindow.ScrollRow = 10\nws.Rows(5).Insert"),
    ("frozen_row_deleted_at_split", 'ws.Range("B3").Select\nActiveWindow.FreezePanes = True\nws.Rows(3).Delete'),
    ("frozen_rows_deleted_all", 'ws.Range("B3").Select\nActiveWindow.FreezePanes = True\nws.Rows("1:2").Delete'),
]

#: Which window Windows(1), (2), ... is as books are added, activated and closed: b1, b2 and b3 are three books
#: this adds, "other" a window that was open before.
ORDER_NAMES = "Private n1 As String, n2 As String, n3 As String\n"
ORDER = """
Private Function Named(text As String) As String
    Named = "other"
    If text = n1 Then Named = "b1"
    If text = n2 Then Named = "b2"
    If text = n3 Then Named = "b3"
End Function

Private Function Listed() As String
    Dim i As Long, out As String
    For i = 1 To Windows.Count
        out = out & Named(Windows(i).Caption) & ":" & Windows(i).Index & " "
    Next
    Listed = Trim(out)
End Function

Private Function Order(updating As Boolean) As String
    Dim b1 As Object, b2 As Object, b3 As Object, out As String, was As Boolean
    was = Application.ScreenUpdating
    Application.ScreenUpdating = updating
    Set b1 = Workbooks.Add(xlWBATWorksheet)
    Set b2 = Workbooks.Add(xlWBATWorksheet)
    Set b3 = Workbooks.Add(xlWBATWorksheet)
    n1 = b1.Name: n2 = b2.Name: n3 = b3.Name
    out = "added " & Listed()
    b1.Activate
    out = out & "~;~b1 activated " & Listed()
    Windows(n2).Activate
    out = out & "~;~b2 activated by its window " & Listed() & " active " & Named(ActiveWorkbook.Name)
    out = out & "~;~b2.Windows " & b2.Windows.Count & " " & Named(b2.Windows(1).Caption) & ":" & _
        b2.Windows(1).Index & " ActiveWindow.Index " & ActiveWindow.Index
    b2.Close False
    out = out & "~;~b2 closed active " & Named(ActiveWorkbook.Name) & " " & Listed()
    b1.Close False
    b3.Close False
    Application.ScreenUpdating = was
    Order = out
End Function
"""


def module() -> str:
    # Panes are laid out only in a window Excel shows: hidden, FreezePanes saves a pane with no split in it.
    lines = ["Public Function Probe() As String", "Dim out As String, was As Boolean", "was = Application.Visible",
             "Application.DisplayAlerts = False", "Application.Visible = True"]
    lines += [f'out = out & B{index}("{FOLDER / (name + ".xlsx")}")' for index, (name, _) in enumerate(BOOKS)]
    lines += ['out = out & "order~:~" & Order(False) & "~|~order_updating~:~" & Order(True)',
              "Application.Visible = was", "Probe = out", "End Function"]
    for index, (name, making) in enumerate(BOOKS):
        lines += [f"Private Function B{index}(path As String) As String",
                  "Dim wb As Object, ws As Object, other As Object, out As String, failed As Long",
                  "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", "wb.Activate",
                  "ActiveWindow.WindowState = xlNormal", "ActiveWindow.Width = 700", "ActiveWindow.Height = 500",
                  "On Error Resume Next", making, "failed = Err.Number", "On Error GoTo 0",
                  f'out = "{name}~:~" & failed & "|" & Seen(ActiveWindow) & "~|~"',
                  "wb.SaveAs Filename:=path, FileFormat:=51", "wb.Close False", f"B{index} = out", "End Function"]
    return ORDER_NAMES + READER + ORDER + "\n".join(lines) + "\n"


def views(path: Path) -> dict[str, str]:
    """Each sheet's sheetViews element and the workbook's bookViews, as Excel wrote them."""
    out: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if name.startswith("xl/worksheets/sheet"):
                found = re.search(r"<sheetViews>.*?</sheetViews>", package.read(name).decode("utf-8"), re.DOTALL)
                out[name] = found.group(0) if found else ""
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        found = re.search(r"<bookViews>.*?</bookViews>", workbook, re.DOTALL)
        out["bookViews"] = found.group(0) if found else ""
    return out


def main() -> None:
    FOLDER.mkdir(parents=True, exist_ok=True)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    for name, _ in BOOKS:
        strip_save_path(FOLDER / f"{name}.xlsx")
    answers = dict(part.split("~:~", 1) for part in str(result.value).split("~|~") if "~:~" in part)
    record = {"reader": READER, "reads": READS, "books": [{"name": name, "making": making, "seen": answers[name],
                                           "views": views(FOLDER / f"{name}.xlsx")} for name, making in BOOKS],
              "order_probe": ORDER_NAMES + ORDER,
              "order": {"screen_updating_off": answers["order"].split("~;~"),
                        "screen_updating_on": answers["order_updating"].split("~;~")}}
    (FOLDER / "windows.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for book in record["books"]:
        print(f"--- {book['name']}: {book['seen']}")
        for part, xml in book["views"].items():
            print(f"    {part}: {xml}")
    for updating, seen in record["order"].items():
        print(f"--- order, {updating}:", *seen, sep="\n    ")


if __name__ == "__main__":
    main()
