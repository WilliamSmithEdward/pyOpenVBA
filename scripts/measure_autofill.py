"""Range.AutoFill, as Excel's object model answers it.

Two kinds of probe run in one new workbook. A layout takes a sheet of its
own: a setup writes a few source cells and fills from them with AutoFill,
then a dump reads A1:E8 back -- each cell's formula, the type of its value,
how far its number sits from its own 15-digit spelling, its number format
and bold -- with the used range and what AutoFill returned or the error it
raised. A column case shares a sheet with the others: a text, or a short
list of values, goes at the top of a column of its own and is filled down
that column alone, and the column is read back with each cell's formula,
type, distance from its spelling and prefix character.

The layouts ask what each fill type makes of numbers, text, dates and
times, of patterns that mix them, of formulas and formats, filling down,
up, right and left, and which sources and destinations Excel refuses. The
column cases ask which texts count as a number to step and which as a
day, month or quarter name.

    python scripts/measure_autofill.py

writes tests/fixtures/autofill.json, which tests/test_excel_autofill.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "autofill.json"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
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

Private Function Off(v As Variant) As String
    If VarType(v) = vbDouble Then Off = CStr(v - CDbl(CStr(v)))
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:E8").Cells
        out = out & c.Formula & "~" & TypeName(c.Value) & "~" & Off(c.Value2) & "~" & c.NumberFormat & "~" & c.Font.Bold & "<>"
    Next
    Dump = out & ws.UsedRange.Address
End Function

Private Function Column(ws As Object, k As Long, depth As Long) As String
    Dim out As String, r As Long, c As Object
    For r = 1 To depth
        Set c = ws.Cells(r, k)
        out = out & c.Formula & "~" & TypeName(c.Value) & "~" & Off(c.Value2) & "~" & c.PrefixCharacter & "~" & _
            c.NumberFormat & "<>"
    Next
    Column = out
End Function
'''


def put(reference: str, value: str) -> str:
    """A VBA line writing ``value`` (a VBA expression) to ``reference``."""
    return f'ws.Range("{reference}").Value = {value}\n'


def formula(reference: str, text: str) -> str:
    return f'ws.Range("{reference}").Formula = "{text}"\n'


def bold(reference: str) -> str:
    return f'ws.Range("{reference}").Font.Bold = True\n'


def number_format(reference: str, code: str) -> str:
    return f'ws.Range("{reference}").NumberFormat = "{code}"\n'


def fill(source: str, destination: str, kind: str = "") -> str:
    extra = f", {kind}" if kind else ""
    return f'v = ws.Range("{source}").AutoFill(ws.Range("{destination}"){extra})'


def column_of(values: list[str], rows: int, kind: str = "") -> str:
    """Fill ``values`` (VBA expressions) down column A to ``rows`` rows."""
    return "".join(put(f"A{row}", value) for row, value in enumerate(values, start=1) if value) + fill(
        f"A1:A{len(values)}", f"A1:A{rows}", kind)


#: name -> setup; each works on ws, the layout's own sheet, and keeps what AutoFill returns in v.
LAYOUTS: dict[str, str] = {
    "number_default": put("A1", "1") + fill("A1", "A1:A5"),
    "number_series": put("A1", "1") + fill("A1", "A1:A5", "xlFillSeries"),
    "number_copy": put("A1", "1") + fill("A1", "A1:A5", "xlFillCopy"),
    "two_numbers": put("A1", "1") + put("A2", "3") + fill("A1:A2", "A1:A6"),
    "two_numbers_copy": put("A1", "1") + put("A2", "3") + fill("A1:A2", "A1:A6", "xlFillCopy"),
    "three_numbers": column_of(["1", "2", "4"], 7),
    "three_numbers_series": column_of(["1", "2", "4"], 7, "xlFillSeries"),
    "three_numbers_growth": column_of(["1", "2", "4"], 7, "xlGrowthTrend"),
    "three_numbers_linear": column_of(["1", "2", "4"], 7, "xlLinearTrend"),
    "decimals": put("A1", "0.5") + put("A2", "0.75") + fill("A1:A2", "A1:A6"),
    "descending": put("A1", "10") + put("A2", "7") + fill("A1:A2", "A1:A6"),
    "text": put("A1", '"abc"') + fill("A1", "A1:A4"),
    "text_series": put("A1", '"abc"') + fill("A1", "A1:A4", "xlFillSeries"),
    "text_number": put("A1", '"Item1"') + fill("A1", "A1:A5"),
    "text_number_copy": put("A1", '"Item1"') + fill("A1", "A1:A5", "xlFillCopy"),
    "text_number_two": put("A1", '"Item1"') + put("A2", '"Item4"') + fill("A1:A2", "A1:A6"),
    "numeric_text": put("A1", '"\'7"') + fill("A1", "A1:A4"),
    "numeric_text_two": put("A1", '"\'7"') + put("A2", '"\'9"') + fill("A1:A2", "A1:A5"),
    "mixed_pattern": put("A1", "1") + put("A2", '"a"') + fill("A1:A2", "A1:A6"),
    "date": put("A1", "#1/30/2020#") + fill("A1", "A1:A5"),
    "date_months": put("A1", "#1/31/2020#") + fill("A1", "A1:A5", "xlFillMonths"),
    "date_years": put("A1", "#2/29/2020#") + fill("A1", "A1:A5", "xlFillYears"),
    "date_weekdays": put("A1", "#1/2/2020#") + fill("A1", "A1:A7", "xlFillWeekdays"),
    "date_days": put("A1", "#1/30/2020#") + fill("A1", "A1:A5", "xlFillDays"),
    "date_series": put("A1", "#1/1/2020#") + fill("A1", "A1:A4", "xlFillSeries"),
    "two_dates": put("A1", "#1/1/2020#") + put("A2", "#1/8/2020#") + fill("A1:A2", "A1:A5"),
    "two_dates_months": put("A1", "#1/31/2020#") + put("A2", "#3/31/2020#") + fill("A1:A2", "A1:A5"),
    "time": put("A1", "#8:00:00 AM#") + fill("A1", "A1:A5"),
    "two_times": put("A1", "#8:00:00 AM#") + put("A2", "#8:30:00 AM#") + fill("A1:A2", "A1:A5"),
    "formula": put("B1", "2") + put("B2", "3") + formula("A1", "=B1*2+$C$1") + fill("A1", "A1:A4"),
    "formula_series": formula("A1", "=ROW()") + fill("A1", "A1:A4", "xlFillSeries"),
    "formula_two": formula("A1", "=B1") + formula("A2", "=B2+1") + fill("A1:A2", "A1:A6"),
    "formula_up": formula("A5", "=B5*2") + fill("A5", "A1:A5"),
    "formula_up_edge": formula("A2", "=A1*2") + fill("A2", "A1:A2"),
    "formula_right": formula("A1", "=$B$1+B$1+$B1+B1") + fill("A1", "A1:C1"),
    "formats": put("A1", "5") + bold("A1") + number_format("A1", "0.00") + fill("A1", "A1:A4"),
    "formats_only": put("A1", "5") + bold("A1") + put("A2", "9") + fill("A1", "A1:A3", "xlFillFormats"),
    "values_only": put("A1", "5") + bold("A1") + fill("A1", "A1:A3", "xlFillValues"),
    "values_two": column_of(["1", "3"], 5, "xlFillValues"),
    "values_formula": formula("A1", "=ROW()*2") + bold("A1") + fill("A1", "A1:A4", "xlFillValues"),
    "values_keep_format": put("A1", "5") + bold("A3") + number_format("A3", "0.00") + fill("A1", "A1:A3",
                                                                                          "xlFillValues"),
    "overwrite": put("A1", "1") + put("A3", '"old"') + bold("A3") + fill("A1", "A1:A4"),
    "format_pattern": put("A1", "1") + number_format("A1", "0.00") + put("A2", "2") + number_format("A2", "0%") +
                      bold("A2") + fill("A1:A2", "A1:A6"),
    "percent": put("A1", "0.1") + number_format("A1:A2", "0%") + put("A2", "0.2") + fill("A1:A2", "A1:A5"),
    "text_cell_number": number_format("A1", "@") + put("A1", '"5"') + fill("A1", "A1:A4"),
    "prefixed_item": put("A1", '"\'Item1"') + fill("A1", "A1:A4"),
    "up": put("A5", "10") + put("A4", "8") + fill("A4:A5", "A1:A5"),
    "up_single": put("A5", "10") + fill("A5", "A1:A5", "xlFillSeries"),
    "up_text_number": put("A4", '"Item9"') + fill("A4", "A1:A4"),
    "up_pattern": put("A4", "1") + put("A5", '"a"') + fill("A4:A5", "A1:A5"),
    "right": put("A1", "1") + put("B1", "2") + fill("A1:B1", "A1:E1"),
    "left": put("E1", '"Item5"') + fill("E1", "A1:E1"),
    "left_names": put("D1", '"Mon"') + put("E1", '"Wed"') + fill("D1:E1", "A1:E1"),
    "columns": put("A1", "1") + put("B1", '"x1"') + put("A2", "2") + put("B2", '"x3"') + fill("A1:B2", "A1:B5"),
    "pattern_three": column_of(["1", '"a"', "2"], 8),
    "pattern_run": column_of(["1", "2", '"a"'], 8),
    "pattern_two_texts": column_of(["1", '"a"', "3", '"b"'], 8),
    "pattern_text_numbers": column_of(['"Item1"', '"Item2"', '"x"'], 8),
    "pattern_number_item": column_of(["1", '"Item1"'], 6),
    "pattern_text_first": column_of(['"a"', "1"], 6),
    "pattern_blank": column_of(["1", "", "3"], 8),
    "pattern_blank_first": column_of(["", "1"], 6),
    "pattern_date_number": column_of(["#1/1/2020#", "5"], 6),
    "pattern_formula": put("B1", "10") + put("B3", "30") + put("B5", "50") + formula("A1", "=B1") +
                       column_of(["", "5"], 6),
    "pattern_formats": put("A1", "1") + bold("A1") + put("A2", "2") + fill("A1:A2", "A1:A6"),
    "pattern_numbers_split": column_of(["1", "2", '"a"', "10"], 8),
    "pattern_dates_run": column_of(["#1/1/2020#", "#1/8/2020#", '"x"'], 8),
    "pattern_mon_number": column_of(['"Mon"', "1"], 6),
    "pattern_numbers_texts": column_of(["1", "2", '"Item1"', '"Item3"'], 8),
    "growth_two": column_of(["2", "6"], 6, "xlGrowthTrend"),
    "growth_single": put("A1", "3") + fill("A1", "A1:A4", "xlGrowthTrend"),
    "growth_negative": column_of(["1", "-2"], 6, "xlGrowthTrend"),
    "growth_zero": column_of(["0", "5"], 6, "xlGrowthTrend"),
    "growth_fraction": column_of(["1", "3"], 8, "xlGrowthTrend"),
    "growth_three": column_of(["1", "2", "5"], 8, "xlGrowthTrend"),
    "growth_text": put("A1", '"Item1"') + fill("A1", "A1:A4", "xlGrowthTrend"),
    "growth_pattern": column_of(["1", '"a"', "4"], 8, "xlGrowthTrend"),
    "linear_single": put("A1", "3") + fill("A1", "A1:A4", "xlLinearTrend"),
    "linear_text": put("A1", '"Item1"') + fill("A1", "A1:A4", "xlLinearTrend"),
    "linear_pattern": column_of(["1", '"a"'], 6, "xlLinearTrend"),
    "months_number": put("A1", "5") + fill("A1", "A1:A4", "xlFillMonths"),
    "months_pair": column_of(["#1/1/2020#", "#1/8/2020#"], 6, "xlFillMonths"),
    "years_pair": column_of(["#1/15/2020#", "#3/15/2020#"], 6, "xlFillYears"),
    "weekdays_text": put("A1", '"abc"') + fill("A1", "A1:A4", "xlFillWeekdays"),
    "weekdays_up": put("A5", "#1/6/2020#") + fill("A5", "A1:A5", "xlFillWeekdays"),
    "weekdays_saturday": put("A1", "#1/4/2020#") + fill("A1", "A1:A5", "xlFillWeekdays"),
    "weekdays_pair": column_of(["#1/6/2020#", "#1/8/2020#"], 7, "xlFillWeekdays"),
    "days_number": put("A1", "5") + fill("A1", "A1:A4", "xlFillDays"),
    "days_text_number": put("A1", '"Item1"') + fill("A1", "A1:A4", "xlFillDays"),
    "series_mon": put("A1", '"Mon"') + fill("A1", "A1:A4", "xlFillSeries"),
    "series_time": put("A1", "#8:00:00 AM#") + fill("A1", "A1:A4", "xlFillSeries"),
    "days_time": put("A1", "#8:00:00 AM#") + fill("A1", "A1:A4", "xlFillDays"),
    "copy_date": put("A1", "#1/1/2020#") + fill("A1", "A1:A4", "xlFillCopy"),
    "invalid_type": put("A1", "1") + fill("A1", "A1:A4", "99"),
    "flash_fill": put("A1", '"abc"') + fill("A1", "A1:A4", "xlFlashFill"),
    "tenths": column_of(["0.1", "0.2"], 8),
    "cross_zero": column_of(["-0.3", "-0.2"], 8),
    "thirds": column_of(["1 / 3", "2 / 3"], 8),
    "least_squares": column_of(["0.1", "0.25", "0.3"], 8),
    "four_points": column_of(["1", "3", "2", "5"], 8),
    "big": column_of(["1E+15", "1000000000000001#"], 6),
    "date_time_pair": column_of(["#1/1/2020 6:00:00 AM#", "#1/2/2020 7:00:00 AM#"], 6),
    "date_time_single": put("A1", "#1/1/2020 6:00:00 AM#") + fill("A1", "A1:A4"),
    "date_months_time": put("A1", "#1/31/2020 6:00:00 AM#") + fill("A1", "A1:A4", "xlFillMonths"),
    "late_time": put("A1", "#11:00:00 PM#") + fill("A1", "A1:A4"),
    "three_dates": column_of(["#1/1/2020#", "#1/2/2020#", "#1/4/2020#"], 7),
    "month_ends": column_of(["#1/31/2020#", "#2/29/2020#"], 6),
    "month_steps": put("A1", '"Jan"') + put("A2", '"Mar"') + put("B1", '"Mon"') + put("B2", '"Wed"') +
                   put("C1", "#1/15/2020#") + put("C2", "#2/15/2020#") + fill("A1:C2", "A1:C6"),
    "years": column_of(["#1/1/2020#", "#1/1/2021#"], 5),
    "leap_years": column_of(["#2/29/2020#", "#2/28/2021#"], 5),
    "backward_dates": column_of(["#3/31/2020#", "#1/31/2020#"], 5),
    "same": column_of(["1", "2"], 2),
    "outside": put("A1", "1") + fill("A1", "B1:B5"),
    "smaller": put("A1", "1") + put("A2", "2") + fill("A1:A2", "A1:A1"),
    "middle_source": put("A2", "1") + put("A3", "2") + fill("A2:A3", "A1:A5"),
    "narrow_destination": put("A1", "1") + put("B1", "2") + fill("A1:B1", "A1:A5"),
    "other_sheet": put("A1", "1") + 'v = ws.Range("A1").AutoFill(ws.Parent.Worksheets.Add().Range("A1:A3"))',
    "multi_source": put("A1", "1") + put("A3", "3") + fill("A1,A3", "A1:A5"),
    "multi_destination": put("A1", "1") + fill("A1", "A1:A3,A5:A6"),
    "destination_missing": put("A1", "1") + 'v = ws.Range("A1").AutoFill()',
    "destination_text": put("A1", "1") + 'v = ws.Range("A1").AutoFill("A1:A5")',
    "blank_source": fill("A1", "A1:A3"),
    "diagonal": put("A1", "1") + fill("A1", "A1:C3"),
    "diagonal_series": put("A1", "1") + fill("A1", "A1:C3", "xlFillSeries"),
    "diagonal_text": put("A1", '"Item1"') + fill("A1", "A1:C3"),
    "diagonal_two": put("A1", "1") + put("A2", "2") + fill("A1:A2", "A1:C4"),
    "diagonal_row": put("A1", "1") + put("B1", "2") + fill("A1:B1", "A1:D3"),
    "diagonal_offset": put("B2", "5") + fill("B2", "B2:D4"),
    "diagonal_tall": put("A1", "1") + fill("A1", "A1:B5"),
    "diagonal_wide": put("A1", "1") + fill("A1", "A1:E2"),
    "row_leading": put("A1", '"1a"') + put("B1", '"1 apple"') + put("C1", '"12ab34"') + put("D1", '"a1b2"') +
                   put("E1", '"a-1"') + fill("A1:E1", "A1:E4"),
    "row_ordinals": put("A1", '"11th"') + put("B1", '"21st"') + put("C1", '"2nd"') + put("D1", '"v1.5"') +
                    put("E1", '"x99"') + fill("A1:E1", "A1:E4"),
    "row_quarters": put("A1", '"Qtr3"') + put("B1", '"Quarter 3"') + put("C1", '"q4"') + put("D1", '"Q5"') +
                    put("E1", '"1st Qtr"') + fill("A1:E1", "A1:E4"),
    "row_two": put("A1", '"2nd"') + put("B1", '"q4"') + fill("A1:B1", "A1:B4"),
    "row_second_alone": put("B1", '"2nd"') + fill("B1", "B1:B4"),
    "row_pair": put("A1", '"x"') + put("B1", '"2nd"') + fill("A1:B1", "A1:B4"),
    "row_numbers": put("A1", "1") + put("B1", '"Item1"') + put("C1", '"Mon"') + fill("A1:C1", "A1:C4"),
    "row_right": put("A1", '"2nd"') + put("A2", '"q4"') + put("A3", '"Item1"') + fill("A1:A3", "A1:C3"),
    "format_month": put("A1", "#1/15/2020#") + number_format("A1", "mmm-yy") + fill("A1", "A1:A4"),
    "format_hours": put("A1", "0.25") + number_format("A1", "[h]:mm") + fill("A1", "A1:A4"),
    "format_minutes": put("A1", "0.25") + number_format("A1", "mm:ss") + fill("A1", "A1:A4"),
    "format_year": put("A1", "#1/15/2020#") + number_format("A1", "yyyy") + fill("A1", "A1:A4"),
    "format_percent": put("A1", "0.25") + number_format("A1", "0%") + fill("A1", "A1:A4"),
    "format_percent_series": put("A1", "0.25") + number_format("A1", "0%") + fill("A1", "A1:A4", "xlFillSeries"),
    "format_text_date": number_format("A1", "@") + put("A1", '"1/15/2020"') + fill("A1", "A1:A4"),
    "format_days_time": put("A1", "#8:00:00 AM#") + fill("A1", "A1:A4", "xlFillWeekdays"),
    "format_date_growth": put("A1", "#1/1/2020#") + put("A2", "#1/2/2020#") + fill("A1:A2", "A1:A5",
                                                                                  "xlGrowthTrend"),
    "format_date_linear": put("A1", "#1/1/2020#") + put("A2", "#1/3/2020#") + fill("A1:A2", "A1:A5",
                                                                                  "xlLinearTrend"),
    "format_date_months_number": put("A1", "#1/31/2020#") + put("A2", "5") + fill("A1:A2", "A1:A6", "xlFillMonths"),
    "booleans": put("A1", "True") + put("B1", "True") + put("B2", "False") + put("C1", "CVErr(xlErrNA)") +
                fill("A1:C2", "A1:C5"),
    "boolean_series": put("A1", "True") + fill("A1", "A1:A4", "xlFillSeries"),
    "error_series": put("A1", "CVErr(xlErrDiv0)") + fill("A1", "A1:A4", "xlFillSeries"),
    "text_run_blank": column_of(['"Item1"', "", '"Item3"'], 8),
    "names_run_blank": column_of(['"Mon"', "", '"Wed"'], 8),
    "numbers_blanks": column_of(["1", "", "", "4"], 8),
    "blanks_then_number": column_of(["", "", "5"], 8),
    "number_blank_text": column_of(["1", "", '"a"'], 8),
    "formula_run": formula("A1", "=1") + put("A2", "2") + put("A3", "3") + fill("A1:A3", "A1:A8"),
}


def row_of(values: list[str], rows: int = 4, kind: str = "") -> str:
    """Fill ``values`` (VBA expressions) across row 1 down to ``rows`` rows."""
    columns = "ABCDE"[: len(values)]
    return "".join(put(f"{column}1", value) for column, value in zip(columns, values, strict=True) if value) + fill(
        f"A1:{columns[-1]}1", f"A1:{columns[-1]}{rows}", kind)


ROWS: dict[str, list[str]] = {
    "numbers": ["1", "2"], "number_text": ["1", '"x"'], "numbers_apart": ["1", "5"], "text_number": ['"x"', "1"],
    "number_item": ["1", '"Item1"'], "items": ['"Item1"', '"Item2"'], "items_apart": ['"Item1"', '"Thing1"'],
    "items_case": ['"Item1"', '"item5"'], "days": ['"Mon"', '"Tue"'], "day_month": ['"Mon"', '"Jan"'],
    "day_lists": ['"Mon"', '"Monday"'], "ordinals": ['"1st"', '"2nd"'], "ordinal_item": ['"1st"', '"Item1"'],
    "quarters": ['"Q1"', '"Q2"'], "dates": ["#1/1/2020#", "#1/2/2020#"], "number_date": ["1", "#1/1/2020#"],
    "number_gap": ["1", "", "2"], "numbers_text": ["1", "2", '"x"', "3"], "item_gap": ['"Item1"', '"x"', '"Item2"'],
    "ordinals_apart": ['"11th"', '"x"', '"21st"'], "ordinals_late": ['"2nd"', '"3rd"'], "equal": ["1", "1"],
    "equal_items": ['"Item1"', '"Item1"'], "times": ["#8:00:00 AM#", "#9:00:00 AM#"],
    "items_uneven": ['"Item1"', '"Item3"', '"Item4"'], "prefixes": ['"a1"', '"b2"'], "three": ["1", "2", "3"],
    "halves": ["1.5", "2.5"], "q_items": ['"q4"', '"Q5"'], "q_apart": ['"q4"', '"x"', '"Q5"'],
    "texts": ['"a"', '"b"'], "number_formula": ["1", '"=A1"'], "numbers_uneven": ["1", "2", "4"],
    "days_apart": ['"Mon"', '"Wed"'], "months": ['"Jan"', '"Feb"'], "ordinals_text": ['"11th"', '"21st"', '"x"'],
}
ROW_KINDS: dict[str, tuple[list[str], str]] = {
    "numbers_series": (["1", "2"], "xlFillSeries"), "numbers_linear": (["1", "2"], "xlLinearTrend"),
    "numbers_values": (["1", "2"], "xlFillValues"), "items_series": (['"Item1"', '"Item2"'], "xlFillSeries"),
    "dates_days": (["#1/1/2020#", "#1/2/2020#"], "xlFillDays"),
}
for row_name, row_values in ROWS.items():
    LAYOUTS[f"row_{row_name}"] = row_of(row_values)
for row_name, (row_values, row_kind) in ROW_KINDS.items():
    LAYOUTS[f"row_{row_name}"] = row_of(row_values, kind=row_kind)
LAYOUTS.update({
    "block_numbers": put("A1", "1") + put("B1", "2") + put("A2", "3") + put("B2", "4") + fill("A1:B2", "A1:B5"),
    "block_mixed": put("A1", "1") + put("B1", "2") + put("A2", '"a"') + put("B2", '"b"') + fill("A1:B2", "A1:B5"),
    "right_numbers": put("A1", "1") + put("A2", "2") + fill("A1:A2", "A1:C2"),
    "right_items": put("A1", '"Item1"') + put("A2", '"Item2"') + fill("A1:A2", "A1:C2"),
    "right_number_text": put("A1", "1") + put("A2", '"x"') + fill("A1:A2", "A1:C2"),
    "format_runs": put("A1", "1") + number_format("A1", "0.00") + put("A2", "2") + number_format("A2", "0.0") +
                   put("B1", "1") + put("B2", "2") + number_format("B2", "0") + put("C1", "1") +
                   number_format("C1:C2", "#,##0") + put("C2", "2") + put("D1", "1") + number_format("D1", "0%") +
                   put("D2", "2") + fill("A1:D2", "A1:D6"),
    "blank_over": put("A1", "1") + put("A3", '"old"') + put("A4", '"old"') + bold("A4") + fill("A1:A2", "A1:A4"),
    "blank_single_over": put("A2", '"old"') + fill("A1", "A1:A3"),
    "block_perpendicular": put("A1", "1") + put("B1", '"x"') + put("A2", "2") + put("B2", "5") +
                           fill("A1:B2", "A1:B6"),
    "block_items": put("A1", '"Item1"') + put("B1", '"Item2"') + put("A2", '"x"') + put("B2", '"y"') +
                   fill("A1:B2", "A1:B6"),
    "row_formats": put("A1", "1") + number_format("A1", "0.00") + put("B1", "2") + number_format("B1", "0%") +
                   put("C1", "1") + put("D1", "2") + number_format("D1", "0.00") + fill("A1:D1", "A1:D4"),
    "row_date_formats": put("A1", "#1/1/2020#") + put("B1", "#1/2/2020#") + number_format("B1", "d-mmm") +
                        put("C1", "#1/1/2020#") + number_format("C1", "mmm-yy") + fill("A1:C1", "A1:C4"),
})

#: Pairs of numbers under two number formats, filled down a column of their own: (first, second).
FORMAT_PAIRS = [
    ("0.00", "0%"), ("0%", "0.00"), ("General", "0%"), ("0%", "0%"), ("0.00", "#,##0"), ("0.00", "$#,##0.00"),
    ("0.00", "0.00E+00"), ("General", "$#,##0.00"), ("0%", "$#,##0.00"), ("General", "# ?/?"), ("General", "h:mm"),
    ("h:mm", "General"), ("m/d/yyyy", "h:mm"), ("General", "@"), ("m/d/yyyy", "d-mmm"), ("m/d/yyyy", "mmm-yy"),
    ("m/d/yyyy", "General"), ("h:mm", "[h]:mm"), ("$#,##0.00", "$#,##0"), ("0.00E+00", "General"),
]

#: Dates and times under a number format, each filled down a column of its own: (format, values, depth, kind).
FORMATTED: list[tuple[str, list[str], int, str]] = [
    ("mmmm yyyy", ["#1/15/2020#"], 5, ""), ("mm/yyyy", ["#1/15/2020#"], 5, ""), ("mmm", ["#1/15/2020#"], 5, ""),
    ("mmmm", ["#1/15/2020#"], 5, ""), ("yyyy-mm", ["#1/15/2020#"], 5, ""), ("dd", ["#1/15/2020#"], 5, ""),
    ("ddd", ["#1/15/2020#"], 5, ""), ("d-mmm", ["#1/15/2020#"], 5, ""), ("mmm d", ["#1/15/2020#"], 5, ""),
    ("dddd", ["#1/15/2020#"], 5, ""), ("mmm-yy", ["#1/15/2020#", "#3/15/2020#"], 6, ""),
    ("yyyy", ["#1/1/2020#", "#1/1/2021#"], 6, ""), ("h", ["0.25"], 5, ""), ("[m]", ["0.25"], 5, ""),
    ("s", ["0.25"], 5, ""), ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#"], 5, ""), ("0.00", ["#1/15/2020#"], 5, ""),
    ("General", ["43845"], 5, ""), ("mmm-yy", ["#1/31/2020#"], 5, ""), ("mmm-yy", ["#1/15/2020#"], 5, "xlFillDays"),
    ("mmm-yy", ["#1/15/2020#"], 5, "xlFillSeries"), ("yy", ["#1/15/2020#"], 5, ""), ("mmmmm", ["#1/15/2020#"], 5, ""),
    ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#"], 5, "xlFillDays"),
    ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#"], 5, "xlFillWeekdays"),
    ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#"], 5, "xlFillYears"),
    ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#"], 5, "xlFillSeries"),
    ("m/d/yyyy h:mm", ["#1/15/2020 6:00:00 AM#", "#1/17/2020 6:00:00 AM#"], 6, "xlFillMonths"),
    ("m/d/yyyy", ["#1/4/2020#", "#1/6/2020#"], 6, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/4/2020#", "#1/5/2020#"], 6, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/3/2020#", "#1/6/2020#"], 6, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/6/2020#", "#1/7/2020#", "#1/8/2020#"], 8, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/6/2020#", "#1/9/2020#"], 6, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/10/2020#", "#1/13/2020#"], 6, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/4/2020#"], 5, "xlFillWeekdays"), ("m/d/yyyy", ["#1/5/2020#"], 5, "xlFillWeekdays"),
    ("m/d/yyyy", ["#1/15/2020#", "#2/15/2020#"], 6, "xlFillYears"),
    ("m/d/yyyy", ["#1/15/2020#", "#1/15/2022#"], 6, "xlFillYears"),
    ("m/d/yyyy", ["#1/15/2020#", "#1/15/2021#"], 6, "xlFillMonths"),
    ("m/d/yyyy", ["#1/15/2020#", "#1/16/2020#", "#1/18/2020#"], 8, "xlFillMonths"),
    ("m/d/yyyy", ["#1/15/2020#", "#1/16/2020#", "#1/18/2020#"], 8, "xlFillDays"),
    ("m/d/yyyy", ["#1/31/2020#", "#2/29/2020#", "#3/31/2020#"], 8, ""),
    ("m/d/yyyy", ["#1/15/2020#", "#2/15/2020#", "#4/15/2020#"], 8, ""),
]

#: Texts filled with a type other than the default: (value, kind).
TYPED_TEXTS = [
    ("Item1", "xlFillSeries"), ("1st", "xlFillSeries"), ("Q1", "xlFillSeries"), ("'7", "xlFillSeries"),
    ("Mon", "xlLinearTrend"), ("Jan", "xlGrowthTrend"), ("Item1", "xlFillValues"), ("Item1", "xlFillFormats"),
    ("Mon", "xlFillCopy"), ("Jan", "xlFillMonths"), ("Mon", "xlFillWeekdays"), ("Mon", "xlFillDays"),
    ("1 apple", "xlFillSeries"), ("Q1", "xlLinearTrend"),
]

#: Pairs of values filled down a column of their own to six rows.
MORE_LISTS = [
    ['"Item 1"', '"Item2"'], ['"Item1"', '"Item  3"'], ['"item1"', '"ITEM3"'], ['"Item1x"', '"Item2x"'],
    ['"1 apple"', '"2 Apple"'], ['"1 apple"', '"2  apple"'], ['"Q1"', '"Qtr2"'], ['"Mon"', '"Tue."'],
    ['"1st"', '"2 x"'], ['"Item1"', '"Item-2"'], ["1", '"1"'], ["1", "#1/1/1900#"],
    ['"Item1"', '"Item.2"'], ['"Item1"', '"Item_2"'], ['"Item1"', '"Item/2"'], ['"Item1"', '"Item#2"'],
    ['"Item1"', '"Item:2"'], ['"Item-1"', '"Item 2"'], ['"Apr"', '"May"'], ['"May"', '"Jun"'], ['"May"', '"June"'],
    ['"Item10"', '"Item09"'], ['"Item00"', '"Item5"'], ['"Item5"', '"Item00"'],
]

#: More single texts filled four rows deep, with a fill type: (value, kind).
MORE_TYPED = [
    ("may", ""), ("MAY", ""), ("1 Q", ""), ("1 Quarter", ""), ("1 qtr", ""), ("2 QTR", ""), ("Thu", "xlFillWeekdays"),
    ("Fri", "xlFillWeekdays"), ("Mon", "xlFillMonths"), ("Jan", "xlFillDays"), ("Jan", "xlFillWeekdays"),
    ("Jan", "xlFillYears"), ("Item1", "xlFillWeekdays"), ("Item00", ""), ("Item0", "xlFillSeries"),
    ("4 Quarter", ""), ("Quarter4", ""), ("q 4", ""), ("Q.4", ""), ("Q_4", ""), ("Q/4", ""), ("Q:4", ""),
    ("Q#4", ""), ("Q,4", ""), ("Qtr-4", ""), ("Q - 3", ""), ("Q. 3", ""), ("Quarter-4", ""), ("Quarter.4", ""),
]

#: Dates filled with a date fill type: (values, depth, kind).
MORE_DATES = [
    (["#1/6/2020#", "#1/20/2020#"], 6, "xlFillWeekdays"), (["#1/6/2020#", "#2/3/2020#"], 6, "xlFillWeekdays"),
    (["#1/6/2020#", "#1/7/2020#", "#1/9/2020#"], 8, "xlFillWeekdays"), (["#1/1/2020#", "#3/5/2020#"], 6, "xlFillMonths"),
    (["#1/1/2020#", "#1/8/2020#", "#1/15/2020#"], 8, "xlFillMonths"), (["#1/1/2020#", "#3/5/2021#"], 6, "xlFillYears"),
    (["#1/1/2020#", "#1/8/2020#", "#1/15/2020#"], 8, "xlFillYears"), (["#1/31/2020#", "#2/29/2020#"], 6, "xlFillMonths"),
    (["#1/6/2020#", "#1/7/2020#"], 6, "xlFillDays"), (["#1/1/2020#", "#1/8/2020#", "#1/15/2020#"], 8, "xlFillWeekdays"),
    (["#1/8/2020#", "#1/6/2020#"], 6, "xlFillWeekdays"), (["#1/7/2020#", "#1/6/2020#"], 6, "xlFillWeekdays"),
    (["#3/1/2020#", "#1/1/2020#"], 6, "xlFillMonths"), (["#1/1/2020#", "#2/1/2020#"], 6, "xlFillYears"),
]

#: Texts each filled down four rows of a column of their own: a leading ' keeps one as text.
TEXTS = [
    "1st", "2nd", "3rd", "4th", "5th", "10th", "11th", "12th", "13th", "21st", "22nd", "23rd", "101st", "111th",
    "1nd", "2st", "1th", "1ST", "1St", "1stx", "1st Qtr", "2nd Qtr", "4th Qtr", "1st Period", "2nd Period",
    "3rd place", "1a", "2a", "'1 a", "2 apple", "5 apple", "10 apple", "0 apple", "1 Qtr", "1  apple", "1.5 apple",
    "-1 apple", "1-apple", "1,apple", "1_a", "12 34", "1 2 3", "Q1", "Q2", "Q3", "Q4", "Q0", "Q5", "Q10", "q1", "q2",
    "q3", "q4", "Qtr1", "Qtr4", "QTR2", "qtr2", "Qtr5", "Quarter1", "Quarter 4", "quarter 2", "Quarter5", "Q 1",
    "Q-1", "Qtr 2", "Quarter  2", "Q1 2020", "Q1-2020", "FY Q1", "1Q", "Item0", "Item-1", "Item 01", "Item007",
    "Item9", "Item99", "Item999999999", "Item9999999999", "Item12345678901234567890", "A9", "AB12", "IV65536",
    "XFD1048576", "x0", "v1.9", "v1.99", "Item1.5", "Item 1.5", "Item_1", "Item(1)", "(1)", "#1", "Item1 ",
    "Item 1a", "Item1Item2", "1a1", "a1b2c3", "'1", "'007", "'-5", "'1.5", "'1e3", "'Item1", "'1-1", "Mon",
    "Monday", "MON", "mon", "MONDAY", "Tues", "Tue", "Thu", "Thur", "Thurs", "Sat", "Sun", "Sunday", "Sept", "Sep",
    "May", "June", "Jun", "July", "December", "Dec", "jan", "JAN", "Janu", "Mo", "M", "Mond", "Mon.", "Mon 1",
    "Month1", "abc", "a b", "Total", "x", " Item1", "Item" + "1" * 15, "Item" + "9" * 16,
]

#: Short lists of values (VBA expressions), each filled down a column of its own to six rows.
LISTS = [
    ['"Item1"', '"Item3"'], ['"Item3"', '"Item1"'], ['"Item1"', '"Item1"'], ['"Item1"', '"Thing2"'],
    ['"Item1"', '"x"'], ['"x"', '"Item1"'], ['"Mon"', '"Wed"'], ['"Wed"', '"Mon"'], ['"Mon"', '"Mon"'],
    ['"Mon"', '"Monday"'], ['"Jan"', '"Mar"'], ['"Jan"', '"Dec"'], ['"Q1"', '"Q3"'], ['"Q1"', '"Q2"'],
    ['"Qtr1"', '"Qtr2"'], ['"1st"', '"3rd"'], ['"1st"', '"2nd"'], ['"a"', '"b"'], ['"1 apple"', '"3 apple"'],
    ['"Item1"', '"item2"'], ['"Item01"', '"Item3"'], ['"\'1"', '"\'3"'], ['"\'1"', "3"], ["1", '"\'3"'],
    ['"Mon"', '"Tue"'], ['"Mon"', '"Thu"'], ['"Item5"', '"Item2"'], ['"Item2"', '"Item5"'], ['"x1"', '"y2"'],
    ['"Item1"', "5"], ['"Item1"', '"Item2"', '"Item4"'], ['"Mon"', '"Tue"', '"Thu"'], ['"Item1"', '"Item1.5"'],
    ['"Item9"', '"Item10"'], ['"Item10"', '"Item9"'], ['"A1"', '"B1"'], ['"Item2"', '"Item1"'],
    ['"Jan"', '"Feb"', '"Mar"'], ['"Item1"', '"x"', '"Item2"'], ['"Q4"', '"Q3"'], ['"Item1"', '"Item5"', '"Item3"'],
    ['"1st Qtr"', '"2nd Qtr"'], ['"Mon"', '"Tue"', '"Wed"'], ['"Item1"', '"Item2"', '"Item3"'], ['"Item3"', '"Item2"'],
    ['"\'5"', '"\'3"'], ['"1 apple"', '"0 apple"'], ['"Q2"', '"Q1"'], ['"Jan"', '"Jan"'], ['"January"', '"Feb"'],
    ['"Mon"', '"Tue"', '"Thu"', '"Fri"'], ['"Item1"', '"Item3"', '"Item5"'], ['"1st"', '"2nd"', '"4th"'],
    ['"2nd"', '"1st"'], ['"Item1"', "", '"Item3"'], ['"Mon"', "", '"Wed"'], ["1", "", "", "4"],
    ['"Item1"', '"Item 2"'], ['"Item01"', '"Item02"'], ['"Item9"', '"Item09"'], ['"a1"', '"A2"'],
    ['"1 x"', '"2 y"'], ['"Q1 2020"', '"Q2 2020"'], ['"Mon."', '"Tue"'],
]

#: Texts filled eight rows deep, to see whether a list wraps round.
LONG_TEXTS = [
    "Q0", "Q 1", "Q-1", "FY Q1", "FY Q4", "1 Qtr", "3 Qtr", "Q4 ", "AQ4", "Qtr.4", "XQtr4", "4th Quarter",
    "4th qtr", "4 Qtr", "5th Qtr", "Quarter 04", "Q04", "q04", "Qtr 04", "Q3a", "QQ4", "Qu4", "Quart4",
]

#: More texts filled four rows deep.
MORE_TEXTS = [
    "Item2147483647", "Item2147483648", "Item1000000000", "Item4294967295", "Item4294967296", "Item99999999",
    "Item 999999999", "mON", "MOn", "tUE", "Jan.", "Mon,", "Mon-", "Mon ", " Mon", "MONday", "Monday.", "Sept.",
    "Febr", "Feb.", "Tues.", "Wednesday", "Wed", "Th", "Sa", "0th", "1st ", "1stQtr", "1st-Qtr", "4TH QTR", "11st",
    "12nd", "13rd", "112th", "1st1", "1 ", "01 apple", "007 x", "9 apple", "99 x", "'9", "'99", "'0.9", "'1,000",
    "'$5", "'5%", "'1E+3", "a1 ", "Item1  ", "1st  Qtr", "Jan 2020x", "Mon1", "MonX", "Monday1",
]

#: Numbers and dates (VBA expressions), each with how deep to fill and the fill type.
NUMERIC: list[tuple[list[str], int, str]] = [
    (["1 / 3", "2 / 3"], 12, ""), (["0.333333333333333", "0.666666666666667"], 12, ""), (["1 / 7", "2 / 7"], 12, ""),
    (["2 / 3", "4 / 3"], 12, ""), (["2 / 3"], 8, "xlFillSeries"), (["1 / 3"], 8, "xlFillSeries"),
    (["0.1", "0.2"], 30, ""), (["1 / 3", "2 / 3", "1"], 12, ""), (["0.7", "0.8"], 12, ""), (["1.1", "1.2"], 12, ""),
    (["1E+15", "1000000000000003#"], 8, ""), (["123456789012345#", "123456789012346#"], 8, ""),
    (["1E-20", "2E-20"], 8, ""), (["-1 / 3", "-2 / 3"], 12, ""), (["5 / 3"], 8, "xlFillSeries"),
    (["0.1", "0.3"], 12, ""), (["1", "1 + 2 ^ -50"], 8, ""), (["10 / 3", "20 / 3"], 12, ""),
    (["1 / 3", "1", "2"], 10, ""), (["1", "2", "3", "5", "8"], 12, ""), (["10", "20", "31"], 10, ""),
    (["1 / 3"], 8, "xlLinearTrend"), (["0.1", "0.2", "0.3"], 12, ""), (["1 / 3", "2 / 3", "3 / 3"], 12, ""),
    (["1", "3"], 12, "xlGrowthTrend"), (["2", "3"], 12, "xlGrowthTrend"), (["1", "1.1"], 12, "xlGrowthTrend"),
    (["3", "1"], 8, "xlGrowthTrend"), (["1", "2", "4", "8"], 10, "xlGrowthTrend"), (["1", "10"], 8, "xlGrowthTrend"),
    (["0.5", "0.25"], 8, "xlGrowthTrend"), (["-1", "-2"], 6, "xlGrowthTrend"), (["2", "2"], 6, "xlGrowthTrend"),
    (["1", "2", "5"], 10, "xlGrowthTrend"), (["2"], 6, "xlGrowthTrend"),
    (["#1/1/2020 6:00:00 AM#", "#1/1/2020 7:00:00 AM#"], 8, ""), (["#1/1/2020#", "#1/2/2020 1:00:00 AM#"], 8, ""),
    (["#1/1/2020 6:00:00 AM#", "#1/2/2020 6:00:00 AM#"], 8, ""),
    (["#1/1/1900 6:00:00 AM#", "#1/2/1900 7:00:00 AM#"], 8, ""), (["#6:00:00 AM#", "#7:00:00 AM#"], 8, ""),
    (["#1/1/2020 6:00:00 AM#", "#1/8/2020 6:00:00 AM#"], 8, ""),
    (["#1/15/2020 6:00:00 AM#", "#2/15/2020 6:00:00 AM#"], 8, ""), (["#1/29/2021#", "#2/28/2021#"], 8, ""),
    (["#1/30/2020#", "#2/29/2020#"], 8, ""), (["#2/28/2021#", "#3/28/2021#"], 8, ""),
    (["#2/28/2021#", "#3/31/2021#"], 8, ""), (["#1/31/2020#", "#1/31/2021#"], 6, ""),
    (["#1/1/2020#", "#1/1/2020#"], 5, ""), (["#1/1/2020#", "#12/31/2019#"], 6, ""),
    (["#1/15/2020#", "#1/15/2020#", "#2/15/2020#"], 8, ""), (["#1/1/2020#", "#2/1/2020#", "#3/1/2020#"], 8, ""),
    (["#1/1/2020#", "#1/3/2020#", "#1/5/2020#"], 8, ""), (["#1/6/2020#", "#1/7/2020#"], 8, "xlFillWeekdays"),
    (["#1/6/2020#", "#1/8/2020#", "#1/10/2020#"], 9, "xlFillWeekdays"), (["#1/1/2020#", "#1/8/2020#"], 6, "xlFillDays"),
    (["#1/1/2020#", "#3/1/2020#"], 6, "xlFillMonths"), (["#1/1/2020#", "#1/1/2021#"], 6, "xlFillYears"),
    (["#1/1/2020#", '"x"'], 6, "xlFillDays"), (["#1/31/2020#", "5"], 6, "xlFillMonths"),
    (["#12:00:00 AM#"], 5, ""), (["#1:30:00 PM#"], 5, ""), (["#1/1/2020#", "#1/3/2020#"], 6, "xlFillWeekdays"),
    (["#1/6/2020#", "#1/13/2020#"], 6, "xlFillWeekdays"), (["#1/1/2020#", "#1/3/2020#"], 6, "xlFillMonths"),
    (["#1/1/2020#", "#1/3/2020#"], 6, "xlFillSeries"), (["#1/31/2020#", "#3/31/2020#"], 6, "xlFillSeries"),
    (["#1/31/2020#", "#3/31/2020#"], 6, "xlFillDays"), (["#1/1/2020#", "#1/2/2020#", "#1/4/2020#"], 7, "xlLinearTrend"),
    (["#1/1/2020#", "#1/2/2020#", "#1/4/2020#"], 7, "xlFillSeries"), (["1", "2", "4"], 7, "xlFillValues"),
    (["1", "2", "4"], 7, "xlFillDays"), (["#1/1/2020#"], 5, "xlGrowthTrend"),
    (["1", "2", "3"], 8, ""), (["10", "20", "30"], 8, ""), (["0.5", "1", "1.5"], 8, ""), (["2020", "2021", "2022"], 8, ""),
    (["100", "90", "80"], 8, ""), (["1.5", "2.5", "3.5", "4.5"], 10, ""), (["-3", "-1", "1"], 8, ""),
    (["0.1", "0.2", "0.3", "0.4"], 10, ""), (["1 / 7", "2 / 7", "3 / 7"], 10, ""), (["0.3", "0.6", "0.9"], 10, ""),
    (["1.1", "2.2", "3.3"], 10, ""), (["1 / 3", "1", "5 / 3"], 10, ""), (["2 / 3", "4 / 3", "2"], 10, ""),
    (["0.7", "0.8", "0.9"], 10, ""), (["1 / 3", "2 / 3", "1"], 12, "xlFillSeries"), (["1", "2", "3"], 6, "xlLinearTrend"),
    (["1 / 9", "2 / 9", "1 / 3"], 10, ""), (["1 / 3", "2 / 3", "1", "4 / 3"], 12, ""), (["0.25", "0.5", "0.75"], 8, ""),
    (["1E+15", "2E+15", "3E+15"], 8, ""), (["1 / 11", "2 / 11", "3 / 11"], 10, ""), (["5 / 7", "1", "9 / 7"], 10, ""),
]

BATCH = 40


def text_expression(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def column_cases() -> list[dict[str, object]]:
    cases: list[dict[str, object]] = []
    for text in TEXTS:
        cases.append({"values": [text_expression(text)], "depth": 4, "kind": ""})
    for values in LISTS:
        cases.append({"values": values, "depth": 6 if len(values) < 3 else 8, "kind": ""})
    for text in LONG_TEXTS:
        cases.append({"values": [text_expression(text)], "depth": 8, "kind": ""})
    for text in MORE_TEXTS:
        cases.append({"values": [text_expression(text)], "depth": 4, "kind": ""})
    for values, depth, kind in NUMERIC:
        cases.append({"values": values, "depth": depth, "kind": kind})
    for code, values, depth, kind in FORMATTED:
        cases.append({"values": values, "depth": depth, "kind": kind, "format": code})
    for text, kind in TYPED_TEXTS:
        cases.append({"values": [text_expression(text)], "depth": 4, "kind": kind})
    for values in MORE_LISTS:
        cases.append({"values": values, "depth": 6, "kind": ""})
    for first, second in FORMAT_PAIRS:
        cases.append({"values": ["1", "2"], "depth": 6, "kind": "", "formats": [first, second]})
    for text, kind in MORE_TYPED:
        cases.append({"values": [text_expression(text)], "depth": 6, "kind": kind})
    for values, depth, kind in MORE_DATES:
        cases.append({"values": values, "depth": depth, "kind": kind})
    for k, case in enumerate(cases, start=1):
        values: list[str] = case["values"]  # type: ignore[assignment]
        extra = f", {case['kind']}" if case["kind"] else ""
        lines = [f'ws.Range(ws.Cells(1, {k}), ws.Cells({len(values)}, {k})).NumberFormat = "{case["format"]}"'] \
            if "format" in case else []
        for row, code in enumerate(case.get("formats", []), start=1):  # type: ignore[arg-type]
            lines.append(f'ws.Cells({row}, {k}).NumberFormat = "{code}"')
        lines += [f"ws.Cells({row}, {k}).Value = {value}" for row, value in enumerate(values, start=1) if value]
        lines.append(f"ws.Range(ws.Cells(1, {k}), ws.Cells({len(values)}, {k})).AutoFill "
                     f"ws.Range(ws.Cells(1, {k}), ws.Cells({case['depth']}, {k})){extra}")
        case["setup"] = "\n".join(lines)
        case["column"] = k
    return cases


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, out As String", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             'out = failed & Show(v) & ";" & Dump(ws)', f"Case{index} = out", "End Function"]
    return "\n".join(lines) + "\n"


def batch_code(index: int, cases: list[dict[str, object]]) -> str:
    lines = [f"Private Function Columns{index}(ws As Object) As String", "Dim out As String", "On Error Resume Next"]
    for case in cases:
        lines += ["Err.Clear", *str(case["setup"]).splitlines(),
                  f'out = out & Err.Number & ":" & Column(ws, {case["column"]}, {case["depth"]}) & "|"']
    lines += [f"Columns{index} = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name[:31]}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup))
    cases = column_cases()
    build += ["Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))", 'ws.Name = "column cases"']
    for index in range(0, len(cases), BATCH):
        build.append(f"out = out & Columns{index // BATCH}(ws)")
        bodies.append(batch_code(index // BATCH, cases[index:index + BATCH]))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")
    layout_answers, column_answers = answers[: len(LAYOUTS)], answers[len(LAYOUTS): len(LAYOUTS) + len(cases)]
    record = {"helper": HELPER,
              "layouts": [{"name": name, "setup": setup, "answers": answer}
                          for (name, setup), answer in zip(LAYOUTS.items(), layout_answers, strict=True)],
              "columns": [{"values": case["values"], "kind": case["kind"],
                           "format": case.get("format", "") or case.get("formats", ""),
                           "setup": case["setup"], "column": case["column"], "depth": case["depth"],
                           "answer": answer}
                          for case, answer in zip(cases, column_answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, layout_answers, strict=True):
        head, _, dump = answer.partition(";")
        if head.startswith("S"):
            head, _, dump = f"{head};{dump}".partition(";")
            head2, _, dump = dump.partition(";")
            head = f"{head};{head2}"
        parts = dump.split("<>")
        cells = [f"{'ABCDE'[i % 5]}{i // 5 + 1}={cell}" for i, cell in enumerate(parts[:40])
                 if not cell.startswith("~Empty~~General~False")]
        print(f"{name}: {head} | {' | '.join(cells)} | used {parts[-1]}")
    for case, answer in zip(cases, column_answers, strict=True):
        error, _, cells = answer.partition(":")
        shown = [cell for cell in cells.split("<>") if cell]
        label = (f"{' / '.join(case['values'])} {case['kind']} {case.get('format', '')}"  # type: ignore[arg-type]
                 f"{case.get('formats', '')}")
        print(f"{label}: {error} | {' | '.join(shown)}")


if __name__ == "__main__":
    main()
