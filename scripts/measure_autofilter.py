"""Range.AutoFilter and the AutoFilter object, as Excel's object model answers them.

Every layout sits on a sheet of its own in one new workbook: the same
small table, then a setup that filters it one way or another. The dump
reads which of rows 1 to 14 are hidden, then the sheet's AutoFilterMode
and FilterMode, the AutoFilter's range, each filter's On, Operator,
Criteria1 and Criteria2 where there are any, the hidden
_FilterDatabase name if there is one, and what the last call returned
or the error it raised. The layouts ask what each kind of criterion
matches -- text, wildcards, numbers, dates, blanks, lists, top and
bottom items, averages -- how two criteria and two columns combine, and
how filters are turned on, cleared and turned off.

    python scripts/measure_autofilter.py

writes tests/fixtures/autofilter.json, which tests/test_excel_autofilter.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "autofilter.json"

HELPER = '''Private Function Show(v As Variant) As String
    Dim i As Long, out As String
    If IsArray(v) Then
        out = "Array("
        For i = LBound(v) To UBound(v)
            out = out & Show(v(i)) & ","
        Next
        Show = out & ")"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        Show = "Object"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Filters(ws As Object) As String
    Dim out As String, i As Long, f As Object, part As String
    On Error Resume Next
    out = "mode=" & ws.AutoFilterMode & " filtering=" & ws.FilterMode
    If ws.AutoFilterMode Then
        out = out & " range=" & ws.AutoFilter.Range.Address & " count=" & ws.AutoFilter.Filters.Count
        For i = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(i)
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
    Err.Clear
    Dim n As Object
    Set n = Nothing
    Set n = ws.Names("_FilterDatabase")
    If Not n Is Nothing Then out = out & " db=" & n.RefersTo & ":" & n.Visible
    Filters = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, r As Long
    For r = 1 To 14
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    Dump = "hidden=" & out & " " & Filters(ws)
End Function
'''

#: The table every layout filters: a header row and twelve rows, blank at row 13 and 14 beyond it.
TABLE = '''ws.Range("A1:D1").Value = Array("Name", "Qty", "Day", "Note")
ws.Range("A2:D2").Value = Array("apple", 1, #1/1/2020#, "x")
ws.Range("A3:D3").Value = Array("Apple", 2, #1/15/2020#, "y")
ws.Range("A4:D4").Value = Array("banana", 2.5, #2/1/2020#, "x")
ws.Range("A5:D5").Value = Array("cherry", 10, #3/1/2020#, "")
ws.Range("A6").Value = "apple pie"
ws.Range("B6").Value = -1
ws.Range("C6").Value = #1/1/2021#
ws.Range("A7").Value = "b*"
ws.Range("B7").Value = 1
ws.Range("B7").NumberFormat = "0.00"
ws.Range("C7").Value = #12/31/2019#
ws.Range("A8").Value = "~"
ws.Range("B8").Value = 100
ws.Range("A9").Value = "'123"
ws.Range("B9").Value = "'5"
ws.Range("C9").Value = "text"
ws.Range("A10").Value = "Banana"
ws.Range("B10").Value = True
ws.Range("C10").Value = #1/15/2020 6:00:00 AM#
ws.Range("A11").Value = "date"
ws.Range("B11").Formula = "=1/0"
ws.Range("A12").Value = "cherry"
ws.Range("B12").Value = 2
ws.Range("C12").Value = #1/15/2020#
ws.Range("D12").Value = "y"
ws.Range("A13").Value = 12
ws.Range("B13").Value = 0.5
'''


def filt(arguments: str = "", reference: str = "A1:D13") -> str:
    return f'v = ws.Range("{reference}").AutoFilter({arguments})\n' if arguments else \
        f'v = ws.Range("{reference}").AutoFilter\n'


#: name -> setup; each works on ws, the layout's own sheet, and keeps the last call's result in v.
LAYOUTS: dict[str, str] = {
    "arrows_only": filt(),
    "arrows_twice": filt() + filt(),
    "single_cell": filt(reference="B3"),
    "equals_text": filt('Field:=1, Criteria1:="apple"'),
    "equals_sign": filt('Field:=1, Criteria1:="=apple"'),
    "not_equal": filt('Field:=1, Criteria1:="<>apple"'),
    "starts": filt('Field:=1, Criteria1:="a*"'),
    "contains": filt('Field:=1, Criteria1:="*an*"'),
    "one_char": filt('Field:=1, Criteria1:="?pple"'),
    "escaped_star": filt('Field:=1, Criteria1:="b~*"'),
    "tilde": filt('Field:=1, Criteria1:="~~"'),
    "blanks": filt('Field:=4, Criteria1:="="'),
    "non_blanks": filt('Field:=4, Criteria1:="<>"'),
    "any_text": filt('Field:=1, Criteria1:="*"'),
    "numeric_text": filt('Field:=1, Criteria1:="123"'),
    "number_equals": filt('Field:=2, Criteria1:="1"'),
    "number_equals_sign": filt('Field:=2, Criteria1:="=1"'),
    "number_shown": filt('Field:=2, Criteria1:="1.00"'),
    "number_greater": filt('Field:=2, Criteria1:=">2"'),
    "number_at_least": filt('Field:=2, Criteria1:=">=2"'),
    "number_below": filt('Field:=2, Criteria1:="<2.5"'),
    "number_not": filt('Field:=2, Criteria1:="<>2"'),
    "number_value": filt('Field:=2, Criteria1:=2'),
    "between_and": filt('Field:=2, Criteria1:=">=1", Operator:=xlAnd, Criteria2:="<=10"'),
    "either_or": filt('Field:=1, Criteria1:="apple", Operator:=xlOr, Criteria2:="cherry"'),
    "values_list": filt('Field:=1, Criteria1:=Array("apple", "cherry"), Operator:=xlFilterValues'),
    "values_list_numbers": filt('Field:=2, Criteria1:=Array("1", "2"), Operator:=xlFilterValues'),
    "values_one": filt('Field:=1, Criteria1:="banana", Operator:=xlFilterValues'),
    "top_items": filt('Field:=2, Criteria1:="3", Operator:=xlTop10Items'),
    "bottom_items": filt('Field:=2, Criteria1:="2", Operator:=xlBottom10Items'),
    "top_percent": filt('Field:=2, Criteria1:="30", Operator:=xlTop10Percent'),
    "above_average": filt('Field:=2, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic'),
    "date_after": filt('Field:=3, Criteria1:=">1/15/2020"'),
    "date_equals": filt('Field:=3, Criteria1:="1/15/2020"'),
    "date_serial": filt('Field:=3, Criteria1:=">=" & CLng(#1/15/2020#)'),
    "two_fields": filt('Field:=1, Criteria1:="cherry"') + filt('Field:=4, Criteria1:="y"'),
    "replace_field": filt('Field:=1, Criteria1:="apple"') + filt('Field:=1, Criteria1:="cherry"'),
    "clear_field": filt('Field:=1, Criteria1:="apple"') + filt("Field:=1"),
    "show_all": filt('Field:=1, Criteria1:="apple"') + "ws.ShowAllData\n",
    "show_all_unfiltered": filt() + "ws.ShowAllData\n",
    "mode_off": filt('Field:=1, Criteria1:="apple"') + "ws.AutoFilterMode = False\n",
    "toggle_off": filt('Field:=1, Criteria1:="apple"') + filt(),
    "field_outside": filt('Field:=5, Criteria1:="x"'),
    "field_zero": filt('Field:=0, Criteria1:="x"'),
    "other_range": filt('Field:=1, Criteria1:="apple"') + filt('Field:=1, Criteria1:="x"', "F1:G3"),
    "narrow": filt('Field:=2, Criteria1:=">5"', "A1:B13"),
    "case_upper": filt('Field:=1, Criteria1:="APPLE"'),
    "boolean": filt('Field:=2, Criteria1:="TRUE"'),
    "error_value": filt('Field:=2, Criteria1:="#DIV/0!"'),
    "text_number_criterion": filt('Field:=2, Criteria1:="5"'),
    "hidden_dropdown": filt('Field:=1, Criteria1:="apple", VisibleDropDown:=False'),
    "apply_filter": filt('Field:=1, Criteria1:="apple"') + 'ws.Range("A2").Value = "cherry"\n' +
                    "ws.AutoFilter.ApplyFilter\n",
    "edit_after": filt('Field:=1, Criteria1:="apple"') + 'ws.Range("A4").Value = "apple"\n',
    "whole_columns": filt('Field:=1, Criteria1:="apple"', "A:D"),
}

#: A clean column of numbers in F1:F9 under a header, for the top, bottom and average filters.
NUMBERS = '''ws.Range("F1").Value = "N"
ws.Range("F2:F9").Value = Application.Transpose(Array(5, 3, 9, 1, 7, 3, 12, 0.5))
'''
LAYOUTS.update({
    "top_number": NUMBERS + filt("Field:=1, Criteria1:=3, Operator:=xlTop10Items", "F1:F9"),
    "top_text_count": NUMBERS + filt('Field:=1, Criteria1:="3", Operator:=xlTop10Items', "F1:F9"),
    "top_ties": NUMBERS + filt("Field:=1, Criteria1:=4, Operator:=xlTop10Items", "F1:F9"),
    "bottom_number": NUMBERS + filt("Field:=1, Criteria1:=2, Operator:=xlBottom10Items", "F1:F9"),
    "top_percent_clean": NUMBERS + filt("Field:=1, Criteria1:=25, Operator:=xlTop10Percent", "F1:F9"),
    "bottom_percent": NUMBERS + filt("Field:=1, Criteria1:=40, Operator:=xlBottom10Percent", "F1:F9"),
    "above_average_clean": NUMBERS + filt("Field:=1, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic",
                                          "F1:F9"),
    "below_average": NUMBERS + filt("Field:=1, Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic", "F1:F9"),
    "top_mixed": filt("Field:=2, Criteria1:=3, Operator:=xlTop10Items"),
    "text_greater": filt('Field:=1, Criteria1:=">b"'),
    "text_below": filt('Field:=1, Criteria1:="<banana"'),
    "text_at_most": filt('Field:=1, Criteria1:="<=banana"'),
    "number_cells_text_criterion": filt('Field:=2, Criteria1:=">a"'),
    "tilde_alone": filt('Field:=1, Criteria1:="~"'),
    "tilde_question": filt('Field:=1, Criteria1:="b~?"'),
    "equals_star": filt('Field:=1, Criteria1:="=*"'),
    "not_star": filt('Field:=1, Criteria1:="<>*"'),
    "not_wildcard": filt('Field:=1, Criteria1:="<>a*"'),
    "spaces": filt('Field:=1, Criteria1:=" apple"'),
    "values_three": filt('Field:=1, Criteria1:=Array("apple", "cherry", "banana"), Operator:=xlFilterValues'),
    "values_numbers_shown": filt('Field:=2, Criteria1:=Array("1.00", "10", "TRUE"), Operator:=xlFilterValues'),
    "values_blank": filt('Field:=4, Criteria1:=Array("x", ""), Operator:=xlFilterValues'),
    "values_dates": filt('Field:=3, Criteria1:=Array("1/15/2020", "3/1/2020"), Operator:=xlFilterValues'),
    "values_wildcard": filt('Field:=1, Criteria1:=Array("a*"), Operator:=xlFilterValues'),
    "time_criterion": filt('Field:=3, Criteria1:=">1/15/2020 5:00"'),
    "equals_serial": filt('Field:=3, Criteria1:="43845"'),
    "number_not_shown": filt('Field:=2, Criteria1:="<>1"'),
    "or_numbers": filt('Field:=2, Criteria1:="<0", Operator:=xlOr, Criteria2:=">50"'),
    "and_text": filt('Field:=1, Criteria1:="a*", Operator:=xlAnd, Criteria2:="<>apple"'),
    "operator_alone": filt('Field:=1, Criteria1:="apple", Operator:=xlAnd'),
    "criteria2_alone": filt('Field:=1, Criteria2:="apple"'),
    "mode_read_missing": "v = ws.AutoFilterMode\n",
    "filter_object_missing": "Set x = ws.AutoFilter\nv = (x Is Nothing)\n",
    "filters_count_after_off": filt() + filt() + "v = ws.AutoFilter Is Nothing\n",
    "range_hidden_rows_before": 'ws.Rows(6).Hidden = True\n' + filt('Field:=1, Criteria1:="banana"'),
    "range_hidden_then_show_all": 'ws.Rows(6).Hidden = True\n' + filt('Field:=1, Criteria1:="apple"') +
                                  "ws.ShowAllData\n",
})

#: F1:F9 as NUMBERS, with one cell swapped for text or an error.
WITH_TEXT = NUMBERS + 'ws.Range("F5").Value = "x"\n'
WITH_ERROR = NUMBERS + 'ws.Range("F5").Formula = "=1/0"\n'
WITH_BLANK = NUMBERS + 'ws.Range("F5").ClearContents\n'
LAYOUTS.update({
    "top_beside_text": WITH_TEXT + filt("Field:=1, Criteria1:=3, Operator:=xlTop10Items", "F1:F9"),
    "top_beside_error": WITH_ERROR + filt("Field:=1, Criteria1:=3, Operator:=xlTop10Items", "F1:F9"),
    "top_beside_blank": WITH_BLANK + filt("Field:=1, Criteria1:=3, Operator:=xlTop10Items", "F1:F9"),
    "average_beside_text": WITH_TEXT + filt("Field:=1, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic",
                                            "F1:F9"),
    "average_beside_error": WITH_ERROR + filt("Field:=1, Criteria1:=xlFilterAboveAverage, "
                                              "Operator:=xlFilterDynamic", "F1:F9"),
    "percent_35": NUMBERS + filt("Field:=1, Criteria1:=35, Operator:=xlTop10Percent", "F1:F9"),
    "percent_30": NUMBERS + filt("Field:=1, Criteria1:=30, Operator:=xlTop10Percent", "F1:F9"),
    "percent_10": NUMBERS + filt("Field:=1, Criteria1:=10, Operator:=xlTop10Percent", "F1:F9"),
    "percent_5": NUMBERS + filt("Field:=1, Criteria1:=5, Operator:=xlTop10Percent", "F1:F9"),
    "top_five_ties": NUMBERS + filt("Field:=1, Criteria1:=5, Operator:=xlTop10Items", "F1:F9"),
    "top_all": NUMBERS + filt("Field:=1, Criteria1:=20, Operator:=xlTop10Items", "F1:F9"),
    "top_zero": NUMBERS + filt("Field:=1, Criteria1:=0, Operator:=xlTop10Items", "F1:F9"),
    "top_fraction": NUMBERS + filt("Field:=1, Criteria1:=2.5, Operator:=xlTop10Items", "F1:F9"),
    "values_wildcards_three": filt('Field:=1, Criteria1:=Array("a*", "b*", "c*"), Operator:=xlFilterValues'),
    "hidden_match_shown": 'ws.Rows(2).Hidden = True\n' + filt('Field:=1, Criteria1:="apple"'),
    "trailing_space": filt('Field:=1, Criteria1:="apple "'),
    "equals_spaces": filt('Field:=4, Criteria1:="= "'),
    "not_number_blanks": filt('Field:=4, Criteria1:="<>5"'),
    "second_range_arrows": filt('Field:=1, Criteria1:="apple"') + NUMBERS + filt(reference="F1:F9"),
    "sub_range": filt('Field:=1, Criteria1:="apple"') + filt('Field:=2, Criteria1:=">1"', "A1:B13"),
    "less_than_date_text": filt('Field:=3, Criteria1:="<2/1/2020"'),
    "not_date": filt('Field:=3, Criteria1:="<>1/15/2020"'),
    "number_greater_text_cell": filt('Field:=1, Criteria1:=">100"'),
})


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, out As String", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             'out = failed & Show(v) & ";" & Dump(ws)', f"Case{index} = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name[:31]}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, TABLE + setup))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "table": TABLE,
              "layouts": [{"name": name, "setup": setup, "answers": answer}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
