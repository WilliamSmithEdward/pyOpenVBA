"""SUBTOTAL, as Excel works it out over hidden, filtered and nested rows.

Every layout gets a new workbook: a table of names and numbers in A1:B9
of a sheet named Data, a second sheet, Other, holding the same table,
then rows hidden or filtered one way or another, and formulas in column
D -- by default SUBTOTAL(n,B2:B9) for n from 1 to 11 and 101 to 111. The
dump is each formula's value, or the error writing it raised.

    python scripts/measure_subtotal.py

writes tests/fixtures/subtotal.json, which tests/test_excel_subtotal.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "subtotal.json"

HELPER = '''Private Function Show(v As Variant) As String
    If IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Place(ws As Object, formulas As Variant) As String
    Dim i As Long, out As String
    For i = LBound(formulas) To UBound(formulas)
        On Error Resume Next
        Err.Clear
        ws.Range("D" & (i + 1)).Formula = formulas(i)
        If Err.Number <> 0 Then out = out & "S" & Err.Number & ";" Else out = out & Show(ws.Range("D" & (i + 1)).Value) & ";"
        On Error GoTo 0
    Next
    Place = out
End Function
'''

TABLE = '''For Each sheet In Array(ws, other)
    sheet.Range("A1:B1").Value = Array("Name", "Qty")
    For r = 2 To 9
        sheet.Cells(r, 1).Value = Choose((r Mod 3) + 1, "apple", "pear", "fig")
        sheet.Cells(r, 2).Value = r - 1
    Next
Next
'''

#: SUBTOTAL(n,B2:B9) for every function number.
STANDARD = [f"=SUBTOTAL({n},B2:B9)" for n in (*range(1, 12), *range(101, 112))]

#: name -> (setup after the table, formulas or None for the standard ones)
LAYOUTS: dict[str, tuple[str, list[str] | None]] = {
    "plain": ("", None),
    "hidden_by_hand": ('ws.Rows(3).Hidden = True\nws.Rows(5).Hidden = True', None),
    "zero_height": ('ws.Rows(3).RowHeight = 0', None),
    "filtered": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"', None),
    "filtered_plus_hand_inside": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\nws.Rows(2).Hidden = True',
                                  None),
    "filtered_plus_hand_outside": ('ws.Range("A1:B6").AutoFilter Field:=1, Criteria1:="apple"\nws.Rows(8).Hidden = True',
                                   None),
    "arrows_only_hand": ('ws.Range("A1:B9").AutoFilter\nws.Rows(3).Hidden = True', None),
    "shown_then_hand": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\nws.ShowAllData\n'
                        'ws.Rows(3).Hidden = True', None),
    "filtered_row_shown_again": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\n'
                                 'ws.Rows(3).Hidden = False', None),
    "filtered_row_hidden_again": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\n'
                                  'ws.Rows(3).Hidden = False\nws.Rows(3).Hidden = True', None),
    "filter_off_rows_kept": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\n'
                             'ws.Rows(3).Hidden = False\nws.Rows(3).Hidden = True\nws.AutoFilterMode = False\n'
                             'ws.Rows(5).Hidden = True', None),
    "mixed_values": ('ws.Range("B3").Value = "x"\nws.Range("B4").Value = True\nws.Range("B5").ClearContents\n'
                     'ws.Range("B6").Value = "\'5"', None),
    "error_in_range": ('ws.Range("B4").Formula = "=1/0"', None),
    "error_hidden": ('ws.Range("B4").Formula = "=1/0"\nws.Rows(4).Hidden = True', None),
    "error_filtered": ('ws.Range("B3").Formula = "=1/0"\nws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"',
                       None),
    "nested": ('ws.Range("B5").Formula = "=SUBTOTAL(9,B2:B4)"', None),
    "nested_expression": ('ws.Range("B5").Formula = "=SUBTOTAL(9,B2:B4)*2"', None),
    "nested_aggregate": ('ws.Range("B5").Formula = "=AGGREGATE(9,0,B2:B4)"', None),
    "nested_sum": ('ws.Range("B5").Formula = "=SUM(B2:B4)"', None),
    "horizontal": ('ws.Range("A12:H12").Value = Array(1, 2, 4, 8, 16, 32, 64, 128)\nws.Columns(3).Hidden = True',
                   ["=SUBTOTAL(9,A12:H12)", "=SUBTOTAL(109,A12:H12)", "=SUBTOTAL(2,A12:H12)",
                    "=SUBTOTAL(102,A12:H12)"]),
    "multiple_refs": ('ws.Rows(3).Hidden = True\nws.Rows(7).Hidden = True',
                      ["=SUBTOTAL(9,B2:B4,B6:B9)", "=SUBTOTAL(109,B2:B4,B6:B9)", "=SUBTOTAL(9,B2:B4,B2:B4)",
                       "=SUBTOTAL(103,A1:A9,B1:B9)"]),
    "function_numbers": ("", ["=SUBTOTAL(0,B2:B9)", "=SUBTOTAL(12,B2:B9)", "=SUBTOTAL(100,B2:B9)",
                              "=SUBTOTAL(112,B2:B9)", "=SUBTOTAL(9.9,B2:B9)", "=SUBTOTAL(\"9\",B2:B9)",
                              "=SUBTOTAL(-9,B2:B9)", "=SUBTOTAL(TRUE,B2:B9)", "=SUBTOTAL(109.5,B2:B9)",
                              "=SUBTOTAL(A2,B2:B9)", "=SUBTOTAL(,B2:B9)"]),
    "arguments": ("", ["=SUBTOTAL(9,{1,2,3})", "=SUBTOTAL(9,5)", "=SUBTOTAL(9,B2:B9*1)", "=SUBTOTAL(9)",
                       "=SUBTOTAL(9,B2)", "=SUBTOTAL(9,\"5\")", "=SUBTOTAL(9,B2:B9,5)"]),
    "other_sheet_hidden": ('other.Rows(3).Hidden = True\nother.Range("A1:B9").AutoFilter Field:=1, Criteria1:="pear"',
                           ["=SUBTOTAL(9,Other!B2:B9)", "=SUBTOTAL(109,Other!B2:B9)", "=SUBTOTAL(3,Other!A1:A9)",
                            "=SUBTOTAL(103,Other!A1:A9)"]),
    "whole_column": ('ws.Rows(3).Hidden = True\nws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"',
                     ["=SUBTOTAL(9,B:B)", "=SUBTOTAL(109,B:B)", "=SUBTOTAL(3,A:A)", "=SUBTOTAL(103,A:A)"]),
    "rows_over_the_filter": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"',
                             ["=SUBTOTAL(9,B1:B12)", "=SUBTOTAL(3,A1:A12)", "=SUBTOTAL(9,B8:B9)"]),
    "filtered_hand_below": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\nws.Range("B11").Value = 100\n'
                            'ws.Range("B12").Value = 200\nws.Rows(11).Hidden = True',
                            ["=SUBTOTAL(9,B2:B12)", "=SUBTOTAL(109,B2:B12)", "=SUBTOTAL(2,B2:B12)"]),
    "filtered_hand_header": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\nws.Rows(1).Hidden = True',
                             ["=SUBTOTAL(3,A1:A9)", "=SUBTOTAL(103,A1:A9)", "=SUBTOTAL(3,A1)"]),
    "filtered_row_shown_by_hand": ('ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\n'
                                   'ws.Rows(2).Hidden = False', None),
    "reference_arguments": ('ws.Parent.Names.Add Name:="Block", RefersTo:="=Data!$B$2:$B$4"\nws.Rows(3).Hidden = True',
                            ["=SUBTOTAL(109,OFFSET(B2,0,0,3))", "=SUBTOTAL(109,INDIRECT(\"B2:B4\"))",
                             "=SUBTOTAL(109,Block)", "=SUBTOTAL(109,(B2:B4))", "=SUBTOTAL(109,B2:B4 B3:B5)",
                             "=SUBTOTAL(109,IF(TRUE,B2:B4))", "=SUBTOTAL(109,INDEX(B2:B9,1):B4)",
                             "=SUBTOTAL(109,(B2:B4,B6:B7))", "=SUBTOTAL(109,B2:B4:B6)"]),
    "hidden_after_the_formula": ('ws.Range("F1").Formula = "=SUBTOTAL(109,B2:B9)"\n'
                                 'ws.Range("F2").Formula = "=SUBTOTAL(9,B2:B9)"\nws.Rows(3).Hidden = True\n'
                                 'ws.Range("G1").Value = ws.Range("F1").Value\nws.Range("G2").Value = ws.Range("F2").Value\n'
                                 'ws.Range("A1:B9").AutoFilter Field:=1, Criteria1:="apple"\n'
                                 'ws.Range("G3").Value = ws.Range("F1").Value\nws.Range("G4").Value = ws.Range("F2").Value\n'
                                 'ws.ShowAllData\nws.Range("G5").Value = ws.Range("F1").Value',
                                 ["=G1", "=G2", "=G3", "=G4", "=G5"]),
}


def case_code(index: int, setup: str, formulas: list[str]) -> str:
    items = ", ".join('"' + formula.replace('"', '""') + '"' for formula in formulas)
    lines = [f"Private Function Case{index}(ws As Object, other As Object) As String",
             "Dim failed As String, r As Long, sheet As Variant", *TABLE.splitlines(), "On Error Resume Next",
             "Err.Clear", *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f"Case{index} = failed & Place(ws, Array({items}))", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, other As Object",
             "Dim out As String", "Application.DisplayAlerts = False"]
    bodies: list[str] = []
    for index, (setup, formulas) in enumerate(LAYOUTS.values()):
        build += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Name = "Data"',
                  "Set other = wb.Worksheets.Add(After:=ws)", 'other.Name = "Other"', "ws.Activate",
                  f'out = out & Case{index}(ws, other) & "|"', "wb.Close False"]
        bodies.append(case_code(index, setup, formulas or STANDARD))
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "table": TABLE, "standard": STANDARD,
              "layouts": [{"name": name, "setup": setup, "formulas": formulas or STANDARD, "answers": answer}
                          for (name, (setup, formulas)), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
