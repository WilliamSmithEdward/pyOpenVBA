"""Manual page breaks in live Excel: HPageBreaks, VPageBreaks, Range.PageBreak, ResetAllPageBreaks and the breaks
Subtotal puts between groups.

A macro gives each case a sheet of its own in one new workbook, fills
A1:C10 (or types a list for the Subtotal cases), adds, sets, deletes or
moves breaks, edits rows or columns and calls Subtotal, an error's number
noted where a step raises one. It reads back each collection's Count and
every break's Location, Type and Extent, and the PageBreak of rows 1 to
15 and columns A to F. The lists are short enough to fit a page, so
Excel works out no automatic break. The workbook is saved, each sheet's
rowBreaks and colBreaks kept, opened again and every sheet read once
more.

    python scripts/measure_page_breaks.py

writes tests/fixtures/page_breaks/: the workbook as Excel saved it and
page_breaks.json, which tests/test_excel_page_breaks.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "page_breaks"

#: A list for the Subtotal cases, rows split by "!" and cells by ",".
LIST = "Region,Product,Sales!East,A,10!East,B,20!West,A,30!West,C,40!North,A,5"
SUB = 'ws.Range("A1").Subtotal 1, xlSum, Array(3)'

#: name -> (the list typed, "" for A1:C10 filled with 1, the steps).
CASES: dict[str, tuple[str, list[str]]] = {
    "none": ("", []),
    "add_row": ("", ['Set b = ws.HPageBreaks.Add(Before:=ws.Range("A5"))',
                     'notes = notes & TypeName(b) & ":" & b.Location.Address & ":" & b.Type & ":" & b.Extent & ":" & '
                     'TypeName(b.Parent) & " "']),
    "add_col": ("", ['Set b = ws.VPageBreaks.Add(Before:=ws.Range("C1"))',
                     'notes = notes & TypeName(b) & ":" & b.Location.Address & ":" & TypeName(b.Parent) & " "']),
    "add_cell_both": ("", ['ws.HPageBreaks.Add Before:=ws.Range("C7")', 'ws.VPageBreaks.Add Before:=ws.Range("C7")']),
    "add_block": ("", ['ws.HPageBreaks.Add Before:=ws.Range("B5:C7")', 'ws.VPageBreaks.Add Before:=ws.Range("B5:C7")']),
    "add_twice": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.HPageBreaks.Add Before:=ws.Range("A5")']),
    "add_order": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A9")', 'ws.HPageBreaks.Add Before:=ws.Range("A3")',
                       'ws.HPageBreaks.Add Before:=ws.Range("A6")', 'ws.VPageBreaks.Add Before:=ws.Range("D1")',
                       'ws.VPageBreaks.Add Before:=ws.Range("B1")']),
    "add_first_row": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A1")']),
    "add_first_col": ("", ['ws.VPageBreaks.Add Before:=ws.Range("A3")']),
    "add_other_sheet": ("", ['ws.HPageBreaks.Add Before:=scratch.Range("A5")']),
    "add_beyond": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A50")', 'ws.VPageBreaks.Add Before:=ws.Range("K1")']),
    "add_edge": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A10")', 'ws.HPageBreaks.Add Before:=ws.Range("A11")',
                      'ws.HPageBreaks.Add Before:=ws.Range("A12")', 'ws.VPageBreaks.Add Before:=ws.Range("C1")',
                      'ws.VPageBreaks.Add Before:=ws.Range("D1")', 'ws.VPageBreaks.Add Before:=ws.Range("E1")']),
    "add_formatted_beyond": ("", ['ws.Range("A20").Interior.Color = RGB(1, 2, 3)',
                                  'ws.HPageBreaks.Add Before:=ws.Range("A15")']),
    "add_print_area": ("", ['ws.PageSetup.PrintArea = "A1:B6"', 'ws.HPageBreaks.Add Before:=ws.Range("A5")',
                            'ws.HPageBreaks.Add Before:=ws.Range("A8")', 'ws.VPageBreaks.Add Before:=ws.Range("B1")',
                            'ws.VPageBreaks.Add Before:=ws.Range("C1")']),
    "add_print_area_edge": ("", ['ws.PageSetup.PrintArea = "B3:C8"', 'ws.HPageBreaks.Add Before:=ws.Range("A3")',
                                 'ws.HPageBreaks.Add Before:=ws.Range("A4")', 'ws.HPageBreaks.Add Before:=ws.Range("A9")',
                                 'ws.VPageBreaks.Add Before:=ws.Range("B1")',
                                 'ws.VPageBreaks.Add Before:=ws.Range("C1")']),
    "delete_first": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A3")', 'ws.HPageBreaks.Add Before:=ws.Range("A6")',
                          "ws.HPageBreaks(1).Delete"]),
    "delete_vertical": ("", ['ws.VPageBreaks.Add Before:=ws.Range("B1")', 'ws.VPageBreaks.Add Before:=ws.Range("C1")',
                             "ws.VPageBreaks(2).Delete"]),
    "delete_beyond": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A50")', "ws.HPageBreaks(1).Delete"]),
    "item_bad": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A3")', "v = ws.HPageBreaks(2).Type"]),
    "item_zero": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A3")', "v = ws.HPageBreaks(0).Type"]),
    "location_set": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")',
                          'Set ws.HPageBreaks(1).Location = ws.Range("A8")']),
    "foreach": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A3")', 'ws.HPageBreaks.Add Before:=ws.Range("A7")',
                     'For Each b In ws.HPageBreaks: notes = notes & b.Location.Row & ",": Next',
                     'notes = notes & TypeName(ws.HPageBreaks) & ":" & TypeName(ws.HPageBreaks.Parent) & ":" & '
                     'TypeName(ws.VPageBreaks) & " "']),
    "pagebreak_cell": ("", ['ws.Range("B5").PageBreak = xlPageBreakManual']),
    "pagebreak_lines": ("", ["ws.Rows(5).PageBreak = xlPageBreakManual", "ws.Columns(3).PageBreak = xlPageBreakManual"]),
    "pagebreak_block": ("", ['ws.Range("B5:C7").PageBreak = xlPageBreakManual']),
    "pagebreak_entire_row": ("", ['ws.Range("B5").EntireRow.PageBreak = xlPageBreakManual']),
    "pagebreak_none_cell": ("", ['ws.Range("B5").PageBreak = xlPageBreakManual',
                                 'ws.Range("B5").PageBreak = xlPageBreakNone']),
    "pagebreak_none_row": ("", ['ws.Range("B5").PageBreak = xlPageBreakManual', "ws.Rows(5).PageBreak = xlPageBreakNone"]),
    "pagebreak_none_col": ("", ['ws.Range("B5").PageBreak = xlPageBreakManual',
                                "ws.Columns(2).PageBreak = xlPageBreakNone"]),
    "pagebreak_none_unbroken": ("", ["ws.Rows(5).PageBreak = xlPageBreakNone"]),
    "pagebreak_column_a_cell": ("", ['ws.Range("A5").PageBreak = xlPageBreakManual']),
    "pagebreak_row_1_cell": ("", ['ws.Range("C1").PageBreak = xlPageBreakManual']),
    "pagebreak_a1": ("", ['ws.Range("A1").PageBreak = xlPageBreakManual']),
    "pagebreak_row_1": ("", ["ws.Rows(1).PageBreak = xlPageBreakManual"]),
    "pagebreak_column_a": ("", ["ws.Columns(1).PageBreak = xlPageBreakManual"]),
    "pagebreak_automatic": ("", ["ws.Rows(5).PageBreak = xlPageBreakAutomatic"]),
    "pagebreak_other": ("", ["ws.Rows(5).PageBreak = 5"]),
    "pagebreak_reads": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.VPageBreaks.Add Before:=ws.Range("C1")',
                             'notes = notes & Try(ws.Range("A5")) & Try(ws.Range("B5:C6")) & Try(ws.Rows("5:6")) & '
                             'Try(ws.Rows("4:5")) & Try(ws.Columns("C:D")) & Try(ws.Range("A:A")) & Try(ws.Range("5:5"))']),
    "reset": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A3")', 'ws.VPageBreaks.Add Before:=ws.Range("C1")',
                   "ws.ResetAllPageBreaks"]),
    "reset_none": ("", ["ws.ResetAllPageBreaks"]),
    # Breaks as rows and columns go in and out.
    "insert_above": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(3).Insert"]),
    "insert_at": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(5).Insert"]),
    "insert_below": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(7).Insert"]),
    "insert_before_at": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(4).Insert"]),
    "insert_many": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.Rows("2:4").Insert']),
    "delete_at": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(5).Delete"]),
    "delete_before_at": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(4).Delete"]),
    "delete_below": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', "ws.Rows(7).Delete"]),
    "delete_span": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.Rows("3:6").Delete']),
    "delete_many_above": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.Rows("2:3").Delete']),
    "delete_two_breaks": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A4")', 'ws.HPageBreaks.Add Before:=ws.Range("A5")',
                               "ws.Rows(4).Delete"]),
    "columns_insert_at": ("", ['ws.VPageBreaks.Add Before:=ws.Range("C1")', "ws.Columns(3).Insert"]),
    "columns_insert_left": ("", ['ws.VPageBreaks.Add Before:=ws.Range("C1")', "ws.Columns(2).Insert"]),
    "columns_delete_at": ("", ['ws.VPageBreaks.Add Before:=ws.Range("C1")', "ws.Columns(3).Delete"]),
    "columns_delete_left": ("", ['ws.VPageBreaks.Add Before:=ws.Range("C1")', "ws.Columns(2).Delete"]),
    "rows_leave_columns": ("", ['ws.VPageBreaks.Add Before:=ws.Range("C1")', "ws.Rows(3).Insert", "ws.Rows(6).Delete"]),
    "cells_insert_down": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")',
                               'ws.Range("A3:C3").Insert Shift:=xlDown']),
    "cells_delete_up": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A5")', 'ws.Range("A3:C3").Delete Shift:=xlUp']),
    "copied": ("", ['ws.HPageBreaks.Add Before:=ws.Range("A4")', 'ws.VPageBreaks.Add Before:=ws.Range("B1")',
                    "ws.Copy After:=ws",
                    'notes = notes & "copy=" & List(wb.Worksheets(ws.Index + 1)) & " "']),
    # The breaks Subtotal puts between groups, and the setting a later call takes from the sheet.
    "sub_below": (LIST, [SUB + ", PageBreaks:=True"]),
    "sub_above": (LIST, [SUB + ", PageBreaks:=True, SummaryBelowData:=xlSummaryAbove"]),
    "sub_remove": (LIST, [SUB + ", PageBreaks:=True", 'ws.Range("A1").RemoveSubtotal']),
    "sub_remove_keeps_outside": (LIST, ['ws.HPageBreaks.Add Before:=ws.Range("A20")', SUB + ", PageBreaks:=True",
                                        'ws.Range("A1").RemoveSubtotal']),
    "sub_remove_inside_manual": (LIST, ['ws.HPageBreaks.Add Before:=ws.Range("A4")', 'ws.Range("A1").RemoveSubtotal']),
    "sub_replace_default": (LIST, [SUB + ", PageBreaks:=True", 'ws.Range("A1").Subtotal 2, xlSum, Array(3)']),
    "sub_replace_false": (LIST, [SUB + ", PageBreaks:=True",
                                 'ws.Range("A1").Subtotal 2, xlSum, Array(3), PageBreaks:=False']),
    "sub_false_then_default": (LIST, [SUB + ", PageBreaks:=True",
                                      'ws.Range("A1").Subtotal 2, xlSum, Array(3), PageBreaks:=False',
                                      'ws.Range("A1").Subtotal 1, xlSum, Array(3)']),
    "sub_remove_then_default": (LIST, [SUB + ", PageBreaks:=True", 'ws.Range("A1").RemoveSubtotal', SUB]),
    "sub_nested_default": (LIST, [SUB + ", PageBreaks:=True", 'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False']),
    "sub_nested_false": (LIST, [SUB + ", PageBreaks:=True",
                                'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, PageBreaks:=False']),
    "sub_reset_then_nested": (LIST, [SUB + ", PageBreaks:=True", "ws.ResetAllPageBreaks",
                                     'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False']),
    "sub_other_sheet": (LIST, [SUB + ", PageBreaks:=True",
                               "Set other = wb.Worksheets.Add(After:=ws)", 'Fill other, "' + LIST + '"',
                               'other.Range("A1").Subtotal 1, xlSum, Array(3)',
                               'notes = notes & "other=" & List(other) & " "', "other.Delete"]),
    "sub_manual_inside": (LIST, ['ws.HPageBreaks.Add Before:=ws.Range("A4")', SUB + ", PageBreaks:=True"]),
    "sub_manual_inside_replace": (LIST, ['ws.HPageBreaks.Add Before:=ws.Range("A4")', SUB]),
    "sub_blank_group": ("Key,Sales!a,1!a,2!,3!b,4", ['ws.Range("A1").Subtotal 1, xlSum, Array(2), PageBreaks:=True']),
    "sub_single_group": ("Key,Sales!a,1!a,2", ['ws.Range("A1").Subtotal 1, xlSum, Array(2), PageBreaks:=True']),
    "sub_above_nested": (LIST, [SUB + ", PageBreaks:=True, SummaryBelowData:=xlSummaryAbove",
                                'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False']),
    "sub_remove_manual_in_data": (LIST, [SUB, 'ws.HPageBreaks.Add Before:=ws.Range("A3")',
                                         'ws.Range("A1").RemoveSubtotal']),
    "sub_remove_manual_after_total": (LIST, [SUB, 'ws.HPageBreaks.Add Before:=ws.Range("A5")',
                                             'ws.Range("A1").RemoveSubtotal']),
    "sub_nested_true_manual": (LIST, [SUB, 'ws.HPageBreaks.Add Before:=ws.Range("A3")',
                                      'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, PageBreaks:=True']),
}

HELPER = '''Private Sub Fill(ws As Object, rows As String)
    Dim r As Long, c As Long, line As Variant, part As Variant
    If rows = "" Then
        ws.Range("A1:C10").Value = 1
        Exit Sub
    End If
    For Each line In Split(rows, "!")
        r = r + 1
        c = 0
        For Each part In Split(line, ",")
            c = c + 1
            If part <> "" Then ws.Cells(r, c).Formula = part
        Next
    Next
End Sub

Private Function Try(target As Object) As String
    Dim v As Variant
    On Error GoTo Failed
    v = target.PageBreak
    Try = target.Address(False, False) & "=" & v & " "
    Exit Function
Failed:
    Try = target.Address(False, False) & "=E" & Err.Number & " "
End Function

Private Function List(ws As Object) As String
    Dim out As String, i As Long, b As Object
    On Error GoTo Failed
    out = "H" & ws.HPageBreaks.Count & "["
    For i = 1 To ws.HPageBreaks.Count
        Set b = ws.HPageBreaks(i)
        out = out & b.Location.Address(False, False) & "/" & b.Type & "/" & b.Extent & " "
    Next
    out = out & "]V" & ws.VPageBreaks.Count & "["
    For i = 1 To ws.VPageBreaks.Count
        Set b = ws.VPageBreaks(i)
        out = out & b.Location.Address(False, False) & "/" & b.Type & "/" & b.Extent & " "
    Next
    List = out & "]"
    Exit Function
Failed:
    List = out & "E" & Err.Number
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, i As Long
    out = List(ws) & "#"
    For i = 1 To 15
        out = out & ws.Rows(i).PageBreak & ","
    Next
    out = out & "#"
    For i = 1 To 6
        out = out & ws.Columns(i).PageBreak & ","
    Next
    Dump = out
End Function
'''

BUILD = HELPER + '''{functions}
Public Function Build(folder As String) As String
    Dim wb As Object, ws As Object, scratch As Object, out As String, name As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set scratch = wb.Worksheets(1)
    scratch.Name = "scratch"
{cases}
    wb.Worksheets(1).Activate
    name = "page_breaks.xlsx"
    wb.SaveAs folder & "\\" & name, 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\" & name)
    out = out & "@@"
    For Each ws In wb.Worksheets
        out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    Build = out
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def case_code(index: int, name: str, rows: str, steps: list[str]) -> str:
    """A function of its own for one case, which a procedure's size limit keeps out of Build."""
    lines = [f"Private Function Make{index}(wb As Object, scratch As Object) As String",
             "    Dim ws As Object, other As Object, b As Object, notes As String, v As Variant",
             "    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
             f'    ws.Name = "{name}"', f"    Fill ws, {vba_text(rows)}", "    On Error Resume Next"]
    for step in steps:
        lines += ["    Err.Clear", f"    {step}",
                  '    If Err.Number <> 0 Then notes = notes & "E" & Err.Number & ":" & Err.Description & " "']
    lines.append("    On Error GoTo 0")
    lines.append(f'    Make{index} = "{name}^" & notes & "@" & Dump(ws) & "~"')
    return "\n".join([*lines, "End Function"]) + "\n"


def sheet_breaks(path: Path) -> dict[str, list[str]]:
    """Each sheet's rowBreaks and colBreaks as the file keeps them, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            out[name] = re.findall(r"<(?:rowBreaks|colBreaks)\b.*?</(?:rowBreaks|colBreaks)>|"
                                   r"<(?:rowBreaks|colBreaks)\b[^>]*/>", text, re.DOTALL)
        return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    functions = "".join(case_code(index, name, *case) for index, (name, case) in enumerate(CASES.items()))
    calls = "".join(f"    out = out & Make{index}(wb, scratch)\n" for index in range(len(CASES)))
    code = BUILD.replace("{functions}", functions).replace("{cases}", calls)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", (str(OUT),), timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("@@", 1)
    path = OUT / "page_breaks.xlsx"
    strip_save_path(path)
    record = {"build": code, "cases": [{"name": name, "rows": rows, "steps": steps}
                                       for name, (rows, steps) in CASES.items()],
              "session": dict(one.split("^", 1) for one in session.split("~") if "^" in one),
              "reopened": dict(one.split("^", 1) for one in reopened.split("~") if "^" in one),
              "breaks": sheet_breaks(path)}
    (OUT / "page_breaks.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, dump in record["session"].items():
        print(f"{name}: {dump}   {record['breaks'].get(name, '')}")


if __name__ == "__main__":
    main()
