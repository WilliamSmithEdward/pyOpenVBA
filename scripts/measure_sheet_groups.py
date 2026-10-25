"""What Sheets(Array(...)) is and does: a group of sheets, counted, walked, selected, copied, moved, deleted, hidden.

Every case makes a new workbook of four sheets, One to Four, each holding its place in A1 and a formula in B1 --
One's reading Two, Two's reading One, Three's reading itself through its own name and Four's a bare cell -- with One
active; then runs one statement, keeping what it gave in v or the error it raised. The answer is v, then the
application's state: whether the active workbook is the case's or a new one, the active sheet, the window's
selected sheets in their order, and the case's workbook's sheets and a new active workbook's, each with its
visibility and the formulas in A1 and B1. New workbooks are closed and the case's workbook too.

    python scripts/measure_sheet_groups.py

writes tests/fixtures/sheet_groups.json, which tests/test_excel_sheet_groups.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "sheet_groups.json"
HELPER = '''Private Function Show(v As Variant) As String
    If IsObject(v) Then
        Show = "Object:" & TypeName(v)
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Listed(wb As Object) As String
    Dim s As Object, out As String
    For Each s In wb.Sheets
        out = out & s.Name & "~" & s.Visible & "~" & s.Range("A1").Formula & "~" & s.Range("B1").Formula & ";"
    Next
    Listed = out
End Function

Private Function State(wb As Object) As String
    Dim out As String, s As Object
    If ActiveWorkbook Is wb Then out = "same" Else out = "new"
    out = out & "|" & ActiveSheet.Name & "|"
    For Each s In ActiveWindow.SelectedSheets
        out = out & s.Name & ","
    Next
    out = out & "|" & Listed(wb)
    If Not ActiveWorkbook Is wb Then out = out & "|" & Listed(ActiveWorkbook)
    State = out
End Function

Private Function Fresh() As Object
    Dim wb As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "One"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Three"
    wb.Worksheets.Add(After:=wb.Worksheets(3)).Name = "Four"
    wb.Worksheets("One").Range("A1").Value = 1
    wb.Worksheets("Two").Range("A1").Value = 2
    wb.Worksheets("Three").Range("A1").Value = 3
    wb.Worksheets("Four").Range("A1").Value = 4
    wb.Worksheets("One").Range("B1").Formula = "=Two!A1+1"
    wb.Worksheets("Two").Range("B1").Formula = "=One!A1*2"
    wb.Worksheets("Three").Range("B1").Formula = "=Three!A1*3"
    wb.Worksheets("Four").Range("B1").Formula = "=A1*4"
    wb.Worksheets("One").Activate
    Set Fresh = wb
End Function
'''
#: Each case: the statements it runs, with wb the case's workbook and v what it keeps.
CASES: dict[str, list[str]] = {
    "count": ['v = wb.Sheets(Array("Three", "One")).Count'],
    "type_name": ['v = TypeName(wb.Sheets(Array("Three", "One")))'],
    "walked": ['For Each s In wb.Sheets(Array("Three", "One"))', "v = v & s.Name & \",\"", "Next"],
    "item": ['v = wb.Sheets(Array("Three", "One"))(1).Name'],
    "item_two": ['v = wb.Sheets(Array("Three", "One")).Item(2).Name'],
    "indexes": ['v = wb.Sheets(Array(3, 1))(1).Name & wb.Sheets(Array(3, 1)).Count'],
    "worksheets": ['v = TypeName(wb.Worksheets(Array("One", "Two"))) & wb.Worksheets(Array("One", "Two")).Count'],
    "mixed": ['v = wb.Sheets(Array("Four", 1))(2).Name'],
    "missing": ['v = wb.Sheets(Array("One", "Nope")).Count'],
    "duplicate": ['v = wb.Sheets(Array("One", "One")).Count'],
    "empty": ['v = wb.Sheets(Array()).Count'],
    "any_case": ['v = wb.Sheets(Array("one", "THREE"))(2).Name'],
    "beyond": ['v = wb.Sheets(Array(1, 5)).Count'],
    "select": ['wb.Sheets(Array("Three", "One")).Select'],
    "select_then_one": ['wb.Sheets(Array("Three", "One")).Select', 'wb.Sheets("Two").Select'],
    "select_added": ['wb.Sheets("Two").Select', 'wb.Sheets("Four").Select Replace:=False'],
    "select_group_added": ['wb.Sheets(Array("Two", "Three")).Select Replace:=False'],
    "copy_new": ['wb.Sheets(Array("Two", "One")).Copy'],
    "copy_new_outside": ['wb.Sheets(Array("One", "Three")).Copy'],
    "copy_after": ['wb.Sheets(Array("Two", "One")).Copy After:=wb.Sheets("Four")'],
    "copy_before": ['wb.Sheets(Array("One", "Three")).Copy Before:=wb.Sheets(1)'],
    "move_new": ['wb.Sheets(Array("Four", "Three")).Move'],
    "move_after": ['wb.Sheets(Array("Four", "One")).Move After:=wb.Sheets("Two")'],
    "move_after_mid": ['wb.Sheets(Array("One", "Three")).Move After:=wb.Sheets("Two")'],
    "move_before_mid": ['wb.Sheets(Array("One", "Four")).Move Before:=wb.Sheets("Three")'],
    "move_before": ['wb.Sheets(Array("Four", "Three")).Move Before:=wb.Sheets("One")'],
    "move_before_late": ['wb.Sheets(Array("Three", "Four")).Move Before:=wb.Sheets("Two")'],
    "move_one_after": ['wb.Sheets("One").Move After:=wb.Sheets("Two")'],
    "move_group_one_after": ['wb.Sheets(Array("One")).Move After:=wb.Sheets("Two")'],
    "copy_after_mid": ['wb.Sheets(Array("Three", "One")).Copy After:=wb.Sheets("Two")'],
    "delete": ["Application.DisplayAlerts = False", 'wb.Sheets(Array("Four", "Three")).Delete',
               "Application.DisplayAlerts = True"],
    "delete_all": ["Application.DisplayAlerts = False", 'wb.Sheets(Array("One", "Two", "Three", "Four")).Delete',
                   "Application.DisplayAlerts = True"],
    "delete_read": ["Application.DisplayAlerts = False", 'wb.Sheets(Array("Two")).Delete',
                    "Application.DisplayAlerts = True"],
    "hidden": ['wb.Sheets(Array("Two", "Three")).Visible = xlSheetHidden'],
    "hidden_all": ['wb.Sheets(Array("One", "Two", "Three", "Four")).Visible = xlSheetHidden'],
    "selected_copy": ['wb.Sheets(Array("Two", "One")).Select', "ActiveWindow.SelectedSheets.Copy"],
    "selected_delete": ['wb.Sheets(Array("Three", "Four")).Select', "Application.DisplayAlerts = False",
                        "ActiveWindow.SelectedSheets.Delete", "Application.DisplayAlerts = True"],
    "activate_in_group": ['wb.Sheets(Array("One", "Three")).Select', 'wb.Sheets("Three").Activate'],
    "activate_outside": ['wb.Sheets(Array("One", "Three")).Select', 'wb.Sheets("Two").Activate'],
    "add_in_group": ['wb.Sheets(Array("One", "Three")).Select', 'v = wb.Worksheets.Add.Name'],
    "visible_read": ['v = wb.Sheets(Array("One", "Two")).Visible'],
    "item_by_name": ['v = wb.Sheets(Array("One", "Two"))("Two").Name'],
    "selected_count": ['wb.Sheets(Array("Four", "Two", "One")).Select', "v = ActiveWindow.SelectedSheets.Count"],
    "selected_item": ['wb.Sheets(Array("Four", "Two")).Select', "v = ActiveWindow.SelectedSheets(1).Name"],
    "move_shifted_three": ['wb.Sheets(Array("Four", "One", "Two")).Move Before:=wb.Sheets("Three")'],
    "copy_shifted": ['wb.Sheets(Array("Four", "One")).Copy Before:=wb.Sheets("Two")'],
    "hidden_active": ['wb.Sheets(Array("One", "Four")).Visible = xlSheetHidden'],
    "select_hidden": ['wb.Sheets("Two").Visible = xlSheetHidden', 'wb.Sheets(Array("One", "Two")).Select'],
    "delete_before_active": ['wb.Sheets("Three").Activate', "Application.DisplayAlerts = False",
                             'wb.Sheets("Two").Delete', "Application.DisplayAlerts = True"],
    "delete_active_mid": ['wb.Sheets("Two").Activate', "Application.DisplayAlerts = False",
                          'wb.Sheets("Two").Delete', "Application.DisplayAlerts = True"],
    "delete_active_last": ['wb.Sheets("Four").Activate', "Application.DisplayAlerts = False",
                           'wb.Sheets("Four").Delete', "Application.DisplayAlerts = True"],
    "delete_active_next_hidden": ['wb.Sheets("Three").Visible = xlSheetHidden', 'wb.Sheets("Two").Activate',
                                  "Application.DisplayAlerts = False", 'wb.Sheets("Two").Delete',
                                  "Application.DisplayAlerts = True"],
    "delete_active_hidden_after": ['wb.Sheets("Four").Visible = xlSheetHidden', 'wb.Sheets("Three").Activate',
                                   "Application.DisplayAlerts = False", 'wb.Sheets("Three").Delete',
                                   "Application.DisplayAlerts = True"],
    "delete_group_mid": ['wb.Sheets(Array("Two", "Three")).Select', "Application.DisplayAlerts = False",
                         "ActiveWindow.SelectedSheets.Delete", "Application.DisplayAlerts = True"],
    "delete_group_inactive": ['wb.Sheets("Four").Activate', "Application.DisplayAlerts = False",
                              'wb.Sheets(Array("Two", "One")).Delete', "Application.DisplayAlerts = True"],
    "delete_name_read": ['wb.Names.Add Name:="Far", RefersTo:="=Two!$A$1"', "Application.DisplayAlerts = False",
                         'wb.Sheets("Two").Delete', "Application.DisplayAlerts = True",
                         'v = wb.Names("Far").RefersTo'],
    "hide_active_one": ['wb.Sheets("One").Visible = xlSheetHidden'],
    "hide_active_last": ['wb.Sheets("Four").Activate', 'wb.Sheets("Four").Visible = xlSheetVeryHidden'],
    "hide_active_next_hidden": ['wb.Sheets("Three").Visible = xlSheetHidden', 'wb.Sheets("Two").Activate',
                                'wb.Sheets("Two").Visible = xlSheetHidden'],
    "hide_last_visible": ['wb.Sheets("Two").Visible = xlSheetHidden', 'wb.Sheets("Three").Visible = xlSheetHidden',
                          'wb.Sheets("Four").Visible = xlSheetHidden', 'wb.Sheets("One").Visible = xlSheetHidden'],
    "hide_inactive": ['wb.Sheets("Three").Visible = xlSheetHidden'],
    "all_copy": ["wb.Worksheets.Copy"],
    "all_select": ['wb.Sheets("Three").Activate', "wb.Worksheets.Select"],
    "all_shown": ['wb.Sheets("Two").Visible = xlSheetHidden', 'wb.Sheets("Four").Visible = xlSheetVeryHidden',
                  "wb.Worksheets.Visible = True"],
    "all_hidden": ["wb.Worksheets.Visible = False"],
    "all_delete": ["Application.DisplayAlerts = False", "wb.Worksheets.Delete", "Application.DisplayAlerts = True"],
    "all_move": ["wb.Worksheets.Move"],
    "all_move_after": ['wb.Worksheets.Move After:=wb.Sheets("Two")'],
    "all_visible_read": ["v = wb.Worksheets.Visible"],
    "delete_last_visible": ['wb.Sheets("Two").Visible = xlSheetHidden', 'wb.Sheets("Three").Visible = xlSheetHidden',
                            'wb.Sheets("Four").Visible = xlSheetHidden', "Application.DisplayAlerts = False",
                            'wb.Sheets("One").Delete', "Application.DisplayAlerts = True"],
    "move_anchor_in_after": ['wb.Sheets(Array("One", "Two")).Move After:=wb.Sheets("Two")'],
    "move_anchor_in_before": ['wb.Sheets(Array("One", "Three")).Move Before:=wb.Sheets("One")'],
    "move_anchor_in_two": ['wb.Sheets(Array("Two", "Four")).Move After:=wb.Sheets("Two")'],
    "move_anchor_in_late": ['wb.Sheets(Array("Three", "One")).Move Before:=wb.Sheets("Three")'],
}


def case_code(name: str) -> str:
    lines = [f"Public Function Case_{name}() As String",
             "    Dim wb As Object, v As Variant, s As Object, failed As String, out As String",
             "    Set wb = Fresh()",
             "    On Error Resume Next",
             *(f"    {line}" for line in CASES[name]),
             '    If Err.Number <> 0 Then failed = "E" & Err.Number & "|"',
             "    Err.Clear",
             "    On Error GoTo 0",
             '    out = failed & Show(v) & "|" & State(wb)',
             "    If Not ActiveWorkbook Is wb Then ActiveWorkbook.Close False",
             "    wb.Close False",
             f"    Case_{name} = out",
             "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    answers: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in CASES:
            result = excel.run_vba(HELPER + case_code(name), f"Case_{name}", timeout=300.0)
            assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
            answers[name] = str(result.value)
            print(name, ascii(result.value))
    record = {"helper": HELPER, "cases": [{"name": name, "code": case_code(name), "answer": answers[name]}
                                          for name in CASES]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=True) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
