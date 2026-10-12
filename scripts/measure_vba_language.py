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
}


def macro(name: str) -> str:
    return HEAD + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
