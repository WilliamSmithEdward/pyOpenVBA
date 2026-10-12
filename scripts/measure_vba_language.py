"""Measure native VBA language semantics the sweep found the emulator getting wrong:
a String against a number in comparisons and Select Case, Err.Source, DatePart's weeks.

Each probe is a module whose ``Function Probe() As String`` the fixture records as a
value or a run-time error number (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_language.json'
HEAD = r'''
Private Function VStr() As Variant
VStr = "a"
End Function
Private Function SStr() As String
SStr = "a"
End Function
Private Function VNum() As Variant
VNum = 5
End Function
Private Function Inner3() As Long
Err.Raise 9999, "Inner3", "deep"
End Function
'''
CASES = {
    # A String against a number: a mismatch only when neither is a Variant.
    "string_var_lt_num": 'Dim s As String\ns = "a"\nOn Error Resume Next\nDim r As Boolean\nr = s < 0\nProbe = Err.Number & "|" & r',
    "variant_str_lt_num": 'Dim v As Variant\nv = "a"\nOn Error Resume Next\nDim r As Boolean\nr = v < 0\nProbe = Err.Number & "|" & r',
    "variant_str_gt_num": 'Dim v As Variant\nv = "a"\nProbe = (v > 0) & (v = 0) & (v <> 0) & (0 < v)',
    "literal_str_lt_num": 'On Error Resume Next\nDim r As Boolean\nr = "a" < 0\nProbe = Err.Number & "|" & r',
    "variant_both": 'Dim v As Variant, w As Variant\nv = "a": w = 0\nProbe = (v > w) & "|" & (w < v)',
    "variant_numeric_text_vs_num": 'Dim v As Variant\nv = "10"\nProbe = (v > 9) & "|" & (v = 10) & "|" & (v > 100)',
    "variant_numeric_text_vs_variant_num": 'Dim v As Variant, w As Variant\nv = "10": w = 100\nProbe = (v > w) & "|" & (v = 10)',
    "string_numeric_text_vs_variant_num": 'Dim s As String, w As Variant\ns = "10": w = 100\nProbe = (s > w) & "|" & (s < w)',
    "fn_variant_str": 'On Error Resume Next\nDim r As Boolean\nr = VStr() < 0\nProbe = Err.Number & "|" & r',
    "fn_string_str": 'On Error Resume Next\nDim r As Boolean\nr = SStr() < 0\nProbe = Err.Number & "|" & r',
    "variant_num_vs_string_var": 'Dim s As String\ns = "a"\nOn Error Resume Next\nDim r As Boolean\nr = VNum() < s\nProbe = Err.Number & "|" & r',
    "select_variant": 'Dim v As Variant, s As String\nv = "a"\nSelect Case v\nCase Is < 0: s = "lt"\nCase 1 To 10: s = "range"\nCase Else: s = "else"\nEnd Select\nProbe = s',
    "select_string": 'Dim v As String, s As String\nv = "a"\nOn Error Resume Next\nSelect Case v\nCase Is < 0: s = "lt"\nCase Else: s = "else"\nEnd Select\nProbe = Err.Number & "|" & s',
    "select_mixed_list": 'Dim s As String, x As Variant\nFor Each x In Array(-1, 0, 5, 10, 11, "a")\nSelect Case x\nCase Is < 0: s = s & "n"\nCase 0, 1 To 10: s = s & "r"\nCase Else: s = s & "e"\nEnd Select\nNext\nProbe = s',
    "array_element_variant": 'Dim a As Variant\na = Array("a", 1)\nProbe = (a(0) > a(1)) & "|" & (a(0) > 1)',
    "collection_item": 'Dim c As New Collection\nc.Add "a"\nOn Error Resume Next\nDim r As Boolean\nr = c(1) > 1\nProbe = Err.Number & "|" & r',
    "empty_vs_string_num": 'Dim v As Variant\nProbe = (v < "1") & "|" & (v = "") & "|" & (v < 1)',
    "date_vs_string": 'Dim d As Date, v As Variant\nd = #1/1/2024#\nv = "x"\nProbe = (v > d) & "|" & (d < v)',
    "bool_vs_variant_str": 'Dim v As Variant\nv = "True"\nProbe = (v = True) & "|" & (True = v)',
    "cell_value_text": 'Range("A1").Value = "a"\nOn Error Resume Next\nDim r As Boolean\nr = Range("A1").Value > 1\nProbe = Err.Number & "|" & r',
    # Err.Source of an error raised without one is the project's name.
    "err_source_runtime": 'On Error Resume Next\nDim x As Long\nx = 1 / 0\nProbe = Err.Number & "|" & Err.Source',
    "err_source_raise": 'On Error Resume Next\nErr.Raise 11\nProbe = Err.Description & "|" & Err.Source',
    "err_source_given": 'On Error Resume Next\nInner3\nProbe = Err.Number & "|" & Err.Source & "|" & Err.Description',
    "err_source_cleared": 'On Error Resume Next\nErr.Raise 5\nErr.Clear\nProbe = "[" & Err.Source & "]"',
    # DatePart's week of the year, with each first day and first week.
    "datepart_ww": 'Probe = DatePart("ww", #1/1/2021#) & "|" & DatePart("ww", #1/1/2021#, vbMonday, vbFirstFourDays) & "|" & DatePart("ww", #12/31/2024#, vbMonday, vbFirstFullWeek) & "|" & DatePart("ww", #3/4/2024#) & "|" & DatePart("ww", #1/3/2021#, vbMonday, vbFirstFullWeek)',
    "datepart_ww_more": 'Probe = DatePart("ww", #12/31/2020#, vbSunday, vbFirstFourDays) & "|" & DatePart("ww", #1/4/2021#, vbMonday, vbFirstFourDays) & "|" & DatePart("ww", #12/28/2026#, vbMonday, vbFirstFourDays) & "|" & DatePart("ww", #1/7/2024#, vbSaturday, vbFirstFullWeek)',
    "format_ww_first_day": 'Probe = Format(#1/1/2021#, "ww", vbMonday, vbFirstFourDays) & "|" & Format(#3/4/2024#, "w", vbMonday)',
    # Null as a condition is False; only a Boolean assignment is error 94.
    "null_in_conditions": 'Dim v As Variant, s As String\nv = Null\nIf v Then s = s & "a" Else s = s & "b"\nIf Not v Then s = s & "c" Else s = s & "d"\nDo While v\ns = s & "e"\nExit Do\nLoop\nDo Until v\ns = s & "f"\nExit Do\nLoop\nProbe = s',
    "null_in_iif_select": 'On Error Resume Next\nDim v As Variant, s As String\nv = Null\ns = IIf(v, "t", "f")\nSelect Case v\nCase 1: s = s & "1"\nCase Else: s = s & "e"\nEnd Select\nProbe = Err.Number & "|" & s',
    "null_elseif": 'Dim v As Variant, s As String\nv = Null\nIf False Then\ns = "x"\nElseIf v Then\ns = "y"\nElse\ns = "z"\nEnd If\nProbe = s',
    "null_while_wend": 'Dim n As Long, v As Variant\nv = Null\nWhile v\nn = n + 1\nWend\nProbe = n',
    "null_boolean_assign": 'Dim b As Boolean\nb = Null\nProbe = b',
    # ^ groups left to right; unary minus takes the power after it.
    "power_associativity": 'Probe = (2 ^ 3 ^ 2) & "|" & (-2 ^ -2) & "|" & (2 ^ -3 ^ 2) & "|" & (2 * 3 ^ 2) & "|" & (-3 ^ 2) & "|" & (2 ^ -1 + 1)',
    # Static procedures, On Error GoTo -1, ParamArray aliasing, a parenthesised first argument.
    "static_function": 'Probe = StaticBump() & StaticBump() & StaticBump()',
    "on_error_goto_minus1": 'On Error GoTo h\nErr.Raise 5\nExit Function\nh:\nOn Error GoTo -1\nOn Error Resume Next\nErr.Raise 6\nProbe = Err.Number',
    "on_error_goto_minus1_clears": 'On Error GoTo h\nErr.Raise 5\nExit Function\nh:\nOn Error GoTo -1\nProbe = Err.Number & "|" & Err.Description',
    "paramarray_byref": 'Dim a As Long, b As String\na = 5: b = "x"\nDoubleAll a, b\nProbe = a & "|" & b',
    "paramarray_parenthesised": 'Dim a As Long\na = 5\nDoubleAll (a), a + 0\nProbe = a',
    # DateDiff counts boundaries; DateAdd past the dates is error 5.
    "datediff_boundaries": 'Probe = DateDiff("h", #3/4/2024 11:59:59 PM#, #3/5/2024#) & DateDiff("n", #3/4/2024 1:00:59 PM#, #3/4/2024 1:01:00 PM#) & DateDiff("s", #3/4/2024 1:00:00 PM#, #3/4/2024 1:00:00 PM#) & DateDiff("d", #3/4/2024 11:59:59 PM#, #3/5/2024#) & DateDiff("h", #3/5/2024#, #3/4/2024 11:59:59 PM#)',
    "dateadd_past_years": 'Dim d As Date\nd = DateAdd("yyyy", 10000, #1/1/2024#)\nProbe = d',
    "dateadd_past_days": 'Dim d As Date\nd = DateAdd("d", 3000000, #1/1/2024#)\nProbe = d',
    "dateadd_before_years": 'Dim d As Date\nd = DateAdd("m", -30000, #1/1/2024#)\nProbe = d',
}
HEAD += r'''
Private Sub DoubleAll(ParamArray xs())
Dim i As Long
For i = LBound(xs) To UBound(xs)
xs(i) = xs(i) & xs(i)
Next
End Sub
Private Static Function StaticBump() As Long
Dim n As Long
n = n + 1
StaticBump = n
End Function
'''
ERL_OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_erl.json'
#: Erl reads the harness's own line numbers unless they are off.
ERL_CASES = {
    "erl_unnumbered": 'On Error Resume Next\nDim x As Long\nx = 1 / 0\nProbe = Erl & "|" & Err.Number',
    "erl_after_unnumbered": 'On Error Resume Next\n10 Dim x As Long\nx = 1 / 0\nProbe = Erl',
    "erl_handler": 'On Error GoTo h\n100 Dim y As Long\n110 y = 1\n120 Err.Raise 5\nExit Function\nh:\nProbe = Erl',
    "erl_after_clear": 'On Error Resume Next\n10 Err.Raise 5\nErr.Clear\nProbe = Erl',
    "erl_no_error": '10 Dim x As Long\n20 x = 1\nProbe = Erl',
    "erl_label_with_colon": 'On Error Resume Next\n40: Err.Raise 5\nProbe = Erl',
    "erl_after_executable_line": 'On Error Resume Next\n10 Dim x As Long\n15 x = 5\n20 x = 1 / 0\n30 Probe = Erl',
}


def macro(name: str) -> str:
    return HEAD + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)
    record({name: '\nFunction Probe() As String\n' + body + '\nEnd Function\n' for name, body in ERL_CASES.items()},
           ERL_OUT, line_numbers=False)


if __name__ == '__main__':
    main()
