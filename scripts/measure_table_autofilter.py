"""A table's own AutoFilter, as Excel's object model answers it and as Excel saves it.

Every layout sits on a sheet of its own in one new workbook: the table
scripts/measure_autofilter.py filters, made a ListObject, then a setup
that filters it one way or another. The dump reads which of rows 1 to 15
are hidden; the sheet's AutoFilterMode, FilterMode and the filter it has
of its own, if any; then, for each table, its range, ShowAutoFilter, its
AutoFilter's range and FilterMode and each column's criteria; the hidden
_FilterDatabase name if there is one; and what the last call returned or
the error it raised. The layouts ask what filters a table rather than the
sheet -- the table's range, a cell or part of it, whole columns -- how its
filter turns off and on, clears, and lives beside the sheet's own, and
what SUBTOTAL, edits and the totals row make of the rows it hides.

Then the workbook is saved and opened again, each sheet read the same
way, and what the file holds -- each sheet's sheetPr, its hidden rows and
autoFilter element, each table part's table and autoFilter elements, and
the workbook's defined names -- goes beside what Excel read.

    python scripts/measure_table_autofilter.py

writes tests/fixtures/table_autofilter/: table_autofilter.xlsx as Excel
saved it, and table_autofilter.json, which
tests/test_excel_table_autofilter.py replays.
"""

from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_autofilter import HELPER, NUMBERS, TABLE  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "table_autofilter"

#: Rows 1 to 15 hidden, the sheet's filter as measure_autofilter's Filters reads it, then each table's.
TABLE_DUMP = '''Private Function TableDump(ws As Object) As String
    Dim out As String, r As Long, o As Object, a As Object, f As Object, i As Long, part As String, n As Object
    For r = 1 To 15
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    out = "hidden=" & out & " " & Filters(ws) & " sheet=" & TypeName(ws.AutoFilter)
    On Error Resume Next
    If Not ws.AutoFilter Is Nothing Then out = out & ":" & ws.AutoFilter.Range.Address
    For Each o In ws.ListObjects
        out = out & " [" & o.Name & " " & o.Range.Address & " show=" & o.ShowAutoFilter
        Set a = Nothing
        Set a = o.AutoFilter
        If a Is Nothing Then
            out = out & " af=Nothing"
        Else
            Err.Clear
            out = out & " range=" & a.Range.Address & " fm=" & a.FilterMode & " count=" & a.Filters.Count
            If Err.Number <> 0 Then out = out & " E" & Err.Number
            For i = 1 To a.Filters.Count
                Set f = a.Filters(i)
                part = " f" & i & ":" & f.On
                If f.On Then
                    Err.Clear
                    part = part & " op=" & f.Operator
                    If Err.Number <> 0 Then part = part & " op=E" & Err.Number
                    Err.Clear
                    part = part & " c1=" & Show(f.Criteria1)
                    If Err.Number <> 0 Then part = part & " c1=E" & Err.Number
                    Err.Clear
                    part = part & " c2=" & Show(f.Criteria2)
                    If Err.Number <> 0 Then part = part & " c2=E" & Err.Number
                End If
                out = out & part
            Next
        End If
        out = out & "]"
    Next
    TableDump = out
End Function
'''
#: The layouts' table: measure_autofilter's A1:D13, made a table with its header row.
MADE = 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D13"), , xlYes)\n'
APPLE = 'Field:=1, Criteria1:="apple"'
#: measure_autofilter's numbers in F1:F9 filtered by the sheet, beside the table.
SHEET_FILTER = NUMBERS + 'ws.Range("F1:F9").AutoFilter Field:=1, Criteria1:=">3"\n'


def tf(arguments: str = "", target: str = "o.Range") -> str:
    return f"v = {target}.AutoFilter({arguments})\n" if arguments else f"v = {target}.AutoFilter\n"


#: name -> setup; each works on ws, the layout's own sheet, and o, its table, keeping the last call's result in v.
LAYOUTS: dict[str, str] = {
    "table": "",
    "equals": tf(APPLE),
    "not_equal": tf('Field:=1, Criteria1:="<>apple"'),
    "wildcard": tf('Field:=1, Criteria1:="a*"'),
    "number_greater": tf('Field:=2, Criteria1:=">2"'),
    "between_and": tf('Field:=2, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=10"'),
    "either_or": tf('Field:=1, Criteria1:="apple", Operator:=xlOr, Criteria2:="cherry"'),
    "or_blank": tf('Field:=4, Criteria1:="=", Operator:=xlOr, Criteria2:="x"'),
    "values_list": tf('Field:=1, Criteria1:=Array("apple", "cherry"), Operator:=xlFilterValues'),
    "values_three": tf('Field:=1, Criteria1:=Array("apple", "cherry", "banana"), Operator:=xlFilterValues'),
    "top_items": tf('Field:=2, Criteria1:="3", Operator:=xlTop10Items'),
    "above_average": tf("Field:=2, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic"),
    "date_after": tf('Field:=3, Criteria1:=">1/15/2020"'),
    "blanks": tf('Field:=4, Criteria1:="="'),
    "two_fields": tf('Field:=1, Criteria1:="cherry"') + tf('Field:=4, Criteria1:="y"'),
    "clear_field": tf(APPLE) + tf("Field:=1"),
    # What a filter is called on: the table's own, whatever part of it, or the sheet's.
    "cell_in_table": tf(APPLE, 'ws.Range("B3")'),
    "part_of_table": tf('Field:=2, Criteria1:=">2"', 'ws.Range("B1:C5")'),
    "data_body": tf(APPLE, "o.DataBodyRange"),
    "column_range": tf('Field:=1, Criteria1:=">2"', "o.ListColumns(2).Range"),
    "whole_columns": tf(APPLE, 'ws.Range("A:D")'),
    "overlapping": tf(APPLE, 'ws.Range("C1:F5")'),
    "field_outside": tf('Field:=5, Criteria1:="x"'),
    "hidden_dropdown": tf('Field:=1, Criteria1:="apple", VisibleDropDown:=False'),
    # Turning it off and on, and clearing it.
    "toggle": tf(),
    "toggle_twice": tf() + tf(),
    "toggle_filtered": tf(APPLE) + tf(),
    "cell_toggle": tf(target='ws.Range("B3")'),
    "show_all_object": tf(APPLE) + "o.AutoFilter.ShowAllData\n",
    "show_all_unfiltered": "o.AutoFilter.ShowAllData\n",
    "show_all_sheet": tf(APPLE) + "ws.ShowAllData\n",
    "show_all_sheet_unfiltered": "ws.ShowAllData\n",
    "show_off": tf(APPLE) + "o.ShowAutoFilter = False\n",
    "show_off_on": tf(APPLE) + "o.ShowAutoFilter = False\no.ShowAutoFilter = True\n",
    "show_off_then_filter": "o.ShowAutoFilter = False\n" + tf(APPLE),
    "mode_off": tf(APPLE) + "ws.AutoFilterMode = False\n",
    "apply_filter": tf(APPLE) + 'ws.Range("A2").Value = "cherry"\no.AutoFilter.ApplyFilter\n',
    "edit_after": tf(APPLE) + 'ws.Range("A4").Value = "apple"\n',
    "parent": "v = TypeName(o.AutoFilter.Parent)\n",
    # Beside the sheet's own filter and another table's.
    "beside_sheet_filter": SHEET_FILTER + tf(APPLE),
    "sheet_filter_after": tf(APPLE) + SHEET_FILTER,
    "sheet_show_all_both": SHEET_FILTER + tf(APPLE) + "ws.ShowAllData\n",
    "sheet_mode_off_both": SHEET_FILTER + tf(APPLE) + "ws.AutoFilterMode = False\n",
    "two_tables": NUMBERS + 'Set p = ws.ListObjects.Add(xlSrcRange, ws.Range("F1:F9"), , xlYes)\n' + tf(APPLE) +
                  tf('Field:=1, Criteria1:=">3"', "p.Range"),
    # The rows it hides, as SUBTOTAL, edits and the totals row take them.
    "subtotal": tf(APPLE) + 'ws.Range("H1").Formula = "=SUBTOTAL(9,B2:B13)"\n'
                            'ws.Range("H2").Formula = "=SUBTOTAL(109,B2:B13)"\n'
                            'ws.Range("H3").Formula = "=SUBTOTAL(3,A2:A13)"\n'
                            'v = ws.Range("H1").Value & "/" & ws.Range("H2").Value & "/" & ws.Range("H3").Value\n',
    "subtotal_hand_hidden": "ws.Rows(4).Hidden = True\n" + tf('Field:=2, Criteria1:="<>1"') +
                            'ws.Range("H1").Formula = "=SUBTOTAL(3,A2:A13)"\nv = ws.Range("H1").Value\n',
    "edit_outside": tf(APPLE) + 'ws.Range("F1:F13").Value = 7\nv = WorksheetFunction.Sum(ws.Range("F1:F13"))\n',
    "edit_inside": tf(APPLE) + 'o.ListColumns(4).DataBodyRange.Value = "z"\n'
                               'v = WorksheetFunction.CountIf(ws.Range("D2:D13"), "z")\n',
    "visible_cells": tf(APPLE) + "v = o.Range.SpecialCells(xlCellTypeVisible).Address\n",
    "totals": "o.ShowTotals = True\n" + tf('Field:=2, Criteria1:=">2"'),
    "totals_sum": "o.ShowTotals = True\no.ListColumns(2).TotalsCalculation = xlTotalsCalculationSum\n" +
                  tf('Field:=2, Criteria1:=">2"') + 'v = ws.Range("B14").Value\n',
    "no_headers": "o.ShowHeaders = False\nv = TypeName(o.AutoFilter)\n",
    "no_headers_filter": "o.ShowHeaders = False\n" + tf(APPLE),
    # The table changing under its filter.
    "resize_filtered": tf(APPLE) + 'o.Resize ws.Range("A1:D10")\n',
    # What the filter's range reads once the file is opened again, with data left under the table.
    "resize_unfiltered": 'o.Resize ws.Range("A1:D10")\n',
    "resize_gap": tf(APPLE) + 'ws.Range("A11:D11").ClearContents\no.Resize ws.Range("A1:D10")\n',
    "resize_one_column": tf(APPLE) + 'ws.Range("B11:D13").ClearContents\no.Resize ws.Range("A1:D10")\n',
    "row_added": tf(APPLE) + "o.ListRows.Add\n",
    "column_added": tf('Field:=2, Criteria1:=">2"') + "o.ListColumns.Add 1\n",
    "table_deleted": tf(APPLE) + "o.Delete\n",
    "deleted_show_all": tf(APPLE) + "o.Delete\nws.ShowAllData\n",
    "deleted_refilter": tf(APPLE) + "o.Delete\n" + MADE + tf(APPLE) + "o.AutoFilter.ShowAllData\n",
    # Which filter the sheet's members reach, with the active cell in or out of the table.
    "select_outside": tf(APPLE) + 'ws.Range("H20").Select\n',
    "select_outside_show_all": tf(APPLE) + 'ws.Range("H20").Select\nws.ShowAllData\n',
    "select_outside_mode_off": tf(APPLE) + 'ws.Range("H20").Select\nws.AutoFilterMode = False\n',
    "both_select_sheet": SHEET_FILTER + tf(APPLE) + 'ws.Range("F2").Select\n',
    "both_select_outside": SHEET_FILTER + tf(APPLE) + 'ws.Range("H20").Select\n',
    "both_select_sheet_show_all": SHEET_FILTER + tf(APPLE) + 'ws.Range("F2").Select\nws.ShowAllData\n',
    "both_select_sheet_mode_off": SHEET_FILTER + tf(APPLE) + 'ws.Range("F2").Select\nws.AutoFilterMode = False\n',
    "two_tables_select_second": NUMBERS + 'Set p = ws.ListObjects.Add(xlSrcRange, ws.Range("F1:F9"), , xlYes)\n' +
                                tf(APPLE) + tf('Field:=1, Criteria1:=">3"', "p.Range") + 'ws.Range("F2").Select\n',
    "two_tables_show_all": NUMBERS + 'Set p = ws.ListObjects.Add(xlSrcRange, ws.Range("F1:F9"), , xlYes)\n' +
                           tf(APPLE) + tf('Field:=1, Criteria1:=">3"', "p.Range") + 'ws.Range("H20").Select\n'
                           "ws.ShowAllData\n",
    # Whether edits and SUBTOTAL hold back from hidden rows by the filter or by FilterMode.
    "edit_selected_outside": tf(APPLE) + 'ws.Range("H20").Select\nws.Range("F1:F13").Value = 7\n'
                                         'v = WorksheetFunction.Sum(ws.Range("F1:F13"))\n',
    "subtotal_selected_outside": "ws.Rows(4).Hidden = True\n" + tf('Field:=2, Criteria1:="<>1"') +
                                 'ws.Rows(5).Hidden = True\nws.Range("H20").Select\n'
                                 'ws.Range("H1").Formula = "=SUBTOTAL(3,A2:A13)"\nv = ws.Range("H1").Value\n',
    "subtotal_selected_inside": "ws.Rows(4).Hidden = True\n" + tf('Field:=2, Criteria1:="<>1"') +
                                'ws.Rows(5).Hidden = True\nws.Range("H1").Formula = "=SUBTOTAL(3,A2:A13)"\n'
                                'v = ws.Range("H1").Value\n',
}
#: Layouts that make their table themselves, after a filter of the sheet's own.
UNMADE: dict[str, str] = {
    "made_over_filter": f'v = ws.Range("A1:D13").AutoFilter({APPLE})\n' + MADE,
    "made_over_arrows": 'v = ws.Range("A1:D13").AutoFilter\n' + MADE,
}
#: name -> a step taken on that layout's sheet once the saved file is open again, its first table as o: first
#: these, on sheets never activated since, then AFTER_REOPENING, each on its sheet activated.
UNACTIVATED: dict[str, str] = {
    "edit_after": 'ws.Range("F1:F13").Value = 7\nv = WorksheetFunction.Sum(ws.Range("F1:F13"))\n',
    "number_greater": 'ws.Range("H1").Formula = "=SUBTOTAL(3,A2:A13)"\nv = ws.Range("H1").Value\n',
}
AFTER_REOPENING: dict[str, str] = {
    "equals": "o.AutoFilter.ApplyFilter\n",
    "totals": tf('Field:=2, Criteria1:=">3"'),
    "date_after": "o.ShowTotals = True\n",
    "cell_in_table": "ws.ShowAllData\n",
    "two_tables": "o.AutoFilter.ShowAllData\n",
    "data_body": 'ws.Range("F1:F13").Value = 7\nv = WorksheetFunction.Sum(ws.Range("F1:F13"))\n',
    "hidden_dropdown": 'ws.Range("H1").Formula = "=SUBTOTAL(3,A2:A13)"\nv = ws.Range("H1").Value\n',
    "select_outside": 'ws.Range("B3").Select\n',
    "sheet_show_all_both": 'ws.Range("H20").Select\n',
}


def case_code(index: int, setup: str, name: str = "Case") -> str:
    lines = [f"Private Function {name}{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, o As Object, p As Object", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'{name}{index} = failed & Show(v) & ";" & TableDump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def _targets(package: zipfile.ZipFile, rels: str) -> dict[str, tuple[str, str]]:
    """A part's relationships: id -> (type, target)."""
    text = package.read(rels).decode("utf-8")
    found: dict[str, tuple[str, str]] = {}
    for element in re.findall(r"<Relationship\b[^>]*>", text):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
        found[attributes["Id"]] = (attributes["Type"].rsplit("/", 1)[-1], attributes["Target"])
    return found


def _element(text: str, tag: str) -> str:
    found = re.search(rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>", text, re.DOTALL)
    return found.group() if found else ""


def file_parts(package: zipfile.ZipFile) -> dict[str, dict[str, object]]:
    """What each sheet's part and its tables' parts hold of filters, by sheet name."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    sheets = _targets(package, "xl/_rels/workbook.xml.rels")
    found: dict[str, dict[str, object]] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        part = posixpath.normpath(posixpath.join("xl", sheets[rid][1].lstrip("/").removeprefix("xl/")))
        text = package.read(part).decode("utf-8")
        tables: list[dict[str, str]] = []
        rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rels in package.namelist():
            for kind, target in _targets(package, rels).values():
                if kind != "table":
                    continue
                table = package.read(posixpath.normpath(posixpath.join(posixpath.dirname(part), target))).decode()
                head = re.search(r"<table\b[^>]*>", table)
                tables.append({"table": head.group() if head else "", "autoFilter": _element(table, "autoFilter")})
        found[name] = {"sheetPr": _element(text, "sheetPr"), "autoFilter": _element(text, "autoFilter"),
                       "hidden": [int(row) for row in re.findall(r'<row r="(\d+)"[^>]*\bhidden="1"', text)],
                       "tables": tables}
    return found


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "table_autofilter.xlsx"
    if path.exists():
        path.unlink()
    layouts = {name: MADE + setup for name, setup in LAYOUTS.items()} | UNMADE
    build = [HELPER, TABLE_DUMP, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(layouts.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name[:31]}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, TABLE + setup))
    # Then the saved file, opened again, read the same way, and a step taken on some of its sheets.
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "#"', f'Set wb = Workbooks.Open("{path}")',
              "For Each ws In wb.Worksheets", '    out = out & ws.Name & "^" & TableDump(ws) & "|"', "Next",
              'out = out & "#"']
    steps = [(name, step, False) for name, step in UNACTIVATED.items()] + \
        [(name, step, True) for name, step in AFTER_REOPENING.items()]
    for index, (name, step, activated) in enumerate(steps):
        build += [f'Set ws = wb.Worksheets("{name}")', *(["ws.Activate"] if activated else []),
                  f'out = out & "{name}^" & After{index}(ws) & "|"']
        bodies.append(case_code(index, "Set o = ws.ListObjects(1)\n" + step, "After"))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    before, after, later = str(result.value).split("#")
    answers = before.split("|")[: len(layouts)]
    reopened = dict(one.split("^", 1) for one in after.split("|") if one)
    stepped = dict(one.split("^", 1) for one in later.split("|") if one)
    with zipfile.ZipFile(path) as package:
        parts = file_parts(package)
        names = re.findall(r"<definedName\b.*?</definedName>", package.read("xl/workbook.xml").decode("utf-8"))
    record = {"helper": HELPER, "dump": TABLE_DUMP, "table": TABLE, "names": names,
              "layouts": [{"name": name, "setup": setup, "answers": answer, "reopened": reopened.get(name[:31], ""),
                           "file": parts.get(name[:31], {})}
                          for (name, setup), answer in zip(layouts.items(), answers, strict=True)],
              "after_reopening": [{"name": name, "step": step, "activated": activated, "answers": stepped[name]}
                                  for name, step, activated in steps]}
    (OUT / "table_autofilter.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for (name, _), answer in zip(layouts.items(), answers, strict=True):
        print(f"{name}: {answer}")
        again = reopened.get(name[:31], "")
        if answer.split(";", 1)[1] != again:
            print(f"    reopened: {again}")
    for name, _, activated in steps:
        print(f"after reopening {name}{'' if activated else ', never activated'}: {stepped[name]}")
    print("names:", names)


if __name__ == "__main__":
    main()
