"""What Range.TextToColumns leaves in the cells, as live Excel does it.

Each case writes some lines of text down column A of a fresh sheet, splits
them with TextToColumns -- every delimiter flag given, so that no case
depends on what an omitted one comes to -- and reads back each cell of
A1:T10 that holds something: its value's type and text, its formula and
its number format; or ! and the error the call raised. The record keeps
the year the cases ran in, which a date read without one falls in.

    python scripts/measure_text_to_columns.py

writes tests/fixtures/text_to_columns.json, which
tests/test_excel_text_to_columns.py replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "text_to_columns.json"


def _dates(count: int, kind: int) -> str:
    """FieldInfo giving ``count`` columns, numbered in order, all of one ``kind``."""
    return "Array(" + ", ".join(f"Array({column}, {kind})" for column in range(1, count + 1)) + ")"


#: Fields a date column reads otherwise than a three-part date: two parts, with each separator and a year in
#: either place; runs of digits; month names and the three-part forms no order reads. Every line has 20 fields,
#: the columns the dump reads. Some read in the current year, which the record keeps.
SHAPES = [
    "1.5,1/5,13/5,5/13,1.50,12.31,1-5,1 5,5/2020,2020/5,1/2020,12/2020,2020/12,5/20,20/5,5.2020,2020.5,05/2020,5/99,"
    "99/5",
    "5,42,123,1234,12345,123456,1234567,12345678,20200501,05012020,01052020,200501,010520,0105,1.5.2020,31.12.2020,"
    "1.2.20,00123,0,99",
    "May 1 2020,1 May 2020,May 2020,May 1,1 May,Jan-2020,May-1-2020,1-May-2020,1 2 2020,1/2-2020,1/2/202,001/2/2020,"
    "31/4/2020,29/2/2021,29/2/2020,0/2/2020,1/0/2020,1/2/30,1/2/1,1/2/100",
]
#: More of them: separators with no part before them, doubled, trailing, or with spaces round them; a month's name
#: run into digits, written out, in lower or upper case; the calendar's ends; and more runs of digits and spills.
SHAPES_MORE = [
    "-5,/5,1//2/2020,1/2/,1 / 2 / 2020,1 /2/2020,1/ 2/2020,1-/2/2020,May2020,5May2020,January 5 2020,may 5 2020,"
    "MAY-5-2020,5-may-2020,Sept 5 2020,2/29/1900,29/2/1900,1/2/1899,12/31/9999,1/1/10000",
    "1052020,1012020,10520,11220,31220,105,512,0512,1205,3112,5/2020/1,123/4/2020,1/22020,12.2020,2020.12,1-2020,"
    "2020-1,1 2020,2020 1,01.05",
]

#: Each case: the lines written down column A from A1, the arguments TextToColumns is given beyond the delimiter
#: flags, which are all False unless named here, and any setup before the lines are written.
CASES: dict[str, tuple[list[str], str, str]] = {
    "comma": (["a,b,c", "1,2,3", "x,,y", " sp , x ", 'a,"b,c",d', '"q""x",z', "trail,"], "Comma:=True", ""),
    "tab": (["a\tb", "1\t2", "\tlead"], "Tab:=True", ""),
    "semicolon": (["a;b;c", "1;;2"], "Semicolon:=True", ""),
    "space": (["a b  c", " lead"], "Space:=True", ""),
    "space_consecutive": (["a b  c", "  lead", "x   "], "Space:=True, ConsecutiveDelimiter:=True", ""),
    "comma_consecutive": (["a,,b", ",,c", 'a,,"",b'], "Comma:=True, ConsecutiveDelimiter:=True", ""),
    "other": (["a|b|c"], 'Other:=True, OtherChar:="|"', ""),
    "other_long": (["a|b|xc"], 'Other:=True, OtherChar:="|x"', ""),
    "several": (["a,b;c d"], "Comma:=True, Semicolon:=True, Space:=True", ""),
    "typing": (["1,2.5,1e3,1/2/2020,TRUE,$5,50%,(5),5-,00123,1 1/2,12:30,a1,-0,1.5E+308"], "Comma:=True", ""),
    "typing_more": (["Jan 5,5-Jan,2020-01-05,true,FALSE, 7 ,1,000,#N/A,$1,234.50,12:30 PM"], "Comma:=True", ""),
    "trailing_minus_off": (["5-,6"], "Comma:=True, TrailingMinusNumbers:=False", ""),
    "trailing_minus_on": (["5-,6"], "Comma:=True, TrailingMinusNumbers:=True", ""),
    "field_text": (["00123,1/2/2020,5"], "Comma:=True, FieldInfo:=Array(Array(1, 2), Array(2, 2))", ""),
    "field_skip": (["a,b,c"], "Comma:=True, FieldInfo:=Array(Array(2, 9))", ""),
    "field_dmy": (["1/2/2020,13/2/2020"], "Comma:=True, FieldInfo:=Array(Array(1, 4), Array(2, 4))", ""),
    "field_ymd": (["2020/2/1,2020/13/1"], "Comma:=True, FieldInfo:=Array(Array(1, 5), Array(2, 5))", ""),
    "field_general_text": (["a,1"], "Comma:=True, FieldInfo:=Array(Array(1, 1), Array(2, 1))", ""),
    "decimal": (["1.234,5;7,25;1,000"], 'Semicolon:=True, DecimalSeparator:=",", ThousandsSeparator:="."', ""),
    "qualifier_none": (['a,"b,c",d'], "Comma:=True, TextQualifier:=-4142", ""),
    "qualifier_single": (["a,'b,c',d"], "Comma:=True, TextQualifier:=2", ""),
    "destination": (["a,b", "1,2"], 'Comma:=True, Destination:=ws.Range("D3")', ""),
    "destination_other_sheet": (["a,b"], "Comma:=True, Destination:=ws.Parent.Worksheets.Add.Range(\"A1\")", ""),
    "fixed": (["abc123def", "xy45", "longer text here"],
              "DataType:=2, FieldInfo:=Array(Array(0, 1), Array(3, 1), Array(5, 1))", ""),
    "fixed_text": (["00123abc"], "DataType:=2, FieldInfo:=Array(Array(0, 2), Array(5, 1))", ""),
    # Text that reads as a formula, written with an apostrophe so the cell holds the text.
    "formula_text": (["'=1+1,a", "'+2,-3", "'=A1,=B2"], "Comma:=True", ""),
    "field_positional": (["00123,00456,00789"], "Comma:=True, FieldInfo:=Array(Array(3, 2))", ""),
    "field_order": (["00123,00456,00789"], "Comma:=True, FieldInfo:=Array(Array(3, 2), Array(1, 1))", ""),
    "destination_other_sheet_offset": (["a,b"], 'Comma:=True, Destination:=ws.Parent.Worksheets.Add.Range("C5")',
                                       ""),
    "text_formatted_source": (["1,2"], "Comma:=True", 'ws.Range("A1:C1").NumberFormat = "@"'),
    "overwrite": (["a,b"], "Comma:=True", 'ws.Range("B1").Value = "old"'),
    "clear_by_empty": (["a,,c"], "Comma:=True", 'ws.Range("B1").Value = "old"'),
    "short_line": (["a,b,c", "x"], "Comma:=True", 'ws.Range("C2").Value = "old"\nws.Range("B2").Value = "old"'),
    "empty_rows": (["a,b", "", "c,d"], "Comma:=True", ""),
    "number_source": (["x"], "Comma:=True", 'ws.Range("A2").Value = 123.5'),
    "no_delimiter": (["a,b"], "", ""),
    "two_columns": (["a,b"], 'Comma:=True', 'ws.Range("B1").Value = "c,d"'),
    "whole_column": (["a,b", "c,d"], "Comma:=True", ""),
    "formatted_source": (["1,2"], "Comma:=True", 'ws.Range("A1:C1").NumberFormat = "0.00"'),
    # Fields a date column is given that are no date in its order, or no date at all. Every date has its year, so
    # that nothing depends on the year the case runs in.
    "field_dmy_other": (["abc,5,1.5,12:30,5-Jan-2020,TRUE,=1+1,1/2/2020 10:00,1/2/20,1-2-2020, 1/2/2020 ,2020/1/2"],
                        "Comma:=True, FieldInfo:=" + _dates(12, 4), ""),
    "field_mdy": (["13/2/2020,2/30/2020,2/3/2020,2020/2/3,2/3/2020 10:00,2/3/99"],
                  "Comma:=True, FieldInfo:=" + _dates(6, 3), ""),
    "field_ymd_short": (["1/2/3,20/1/2,99/12/31,30/1/2,2020/1/2"], "Comma:=True, FieldInfo:=" + _dates(5, 5), ""),
    "field_orders": (["2/2020/1,1/2020/2,2020/1/2"],
                     "Comma:=True, FieldInfo:=Array(Array(1, 6), Array(2, 7), Array(3, 8))", ""),
    **{f"shapes_{order}": (SHAPES, "Comma:=True, FieldInfo:=" + _dates(20, kind), "")
       for order, kind in (("general", 1), ("mdy", 3), ("dmy", 4), ("ymd", 5), ("myd", 6), ("dym", 7), ("ydm", 8))},
    **{f"shapes_more_{order}": (SHAPES_MORE, "Comma:=True, FieldInfo:=" + _dates(20, kind), "")
       for order, kind in (("general", 1), ("mdy", 3), ("dmy", 4), ("ymd", 5), ("myd", 6), ("dym", 7), ("ydm", 8))},
}
#: Cases called on a range other than the lines written: the whole column, or two columns.
TARGETS = {"two_columns": 'ws.Range("A1:B1")', "whole_column": 'ws.Columns(1)'}
_HELPER = r'''Private Function Dump(ws As Object) As String
    Dim c As Object, out As String, shown As String
    For Each c In ws.Range("A1:T10")
        If Not IsEmpty(c.Value) Or c.HasFormula Then
            If IsError(c.Value) Then
                shown = "Error`" & CStr(CLng(c.Value))
            Else
                shown = TypeName(c.Value) & "`" & Replace(Replace(CStr(c.Value), vbCr, "\r"), vbLf, "\n")
            End If
            out = out & c.Address(False, False) & "=" & shown & "`" & c.Formula & "`" & c.NumberFormat & "^"
        End If
    Next
    Dump = out
End Function
'''


def vba_text(text: str) -> str:
    """A string as a VBA expression: quotes doubled, a tab as vbTab."""
    return " & vbTab & ".join('"' + piece.replace('"', '""') + '"' for piece in text.split("\t"))


def call(name: str, lines: list[str], arguments: str) -> str:
    """The TextToColumns call of a case, every delimiter flag given."""
    flags = dict.fromkeys(("Tab", "Semicolon", "Comma", "Space", "Other"), "False")
    given = [part.strip() for part in arguments.split(", ") if part.strip()] if arguments else []
    named = {part.split(":=", 1)[0] for part in given}
    extra = [f"{flag}:={value}" for flag, value in flags.items() if flag not in named]
    if "ConsecutiveDelimiter" not in named:
        extra.append("ConsecutiveDelimiter:=False")
    if "TextQualifier" not in named:
        extra.append("TextQualifier:=1")
    if "DataType" not in named:
        extra.append("DataType:=1")
    target = TARGETS.get(name, f'ws.Range("A1:A{len(lines)}")')
    return f"{target}.TextToColumns " + ", ".join(given + extra)


def body(name: str) -> str:
    lines, arguments, setup = CASES[name]
    writes = [f'ws.Range("A{row}").Value = {vba_text(text)}' for row, text in enumerate(lines, start=1) if text]
    return "\n".join([setup, *writes, call(name, lines, arguments)]).strip()


def main() -> None:
    record: dict[str, dict[str, str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in CASES:
            code = (_HELPER + "Public Function Report() As String\nDim ws As Object\nApplication.DisplayAlerts = False\n"
                    "Set ws = ActiveWorkbook.Worksheets.Add\nOn Error Resume Next\n" + body(name) + "\n"
                    'If Err.Number <> 0 Then Report = "!" & Err.Number Else Report = Dump(ws)\n'
                    "Application.DisplayAlerts = True\nEnd Function\n")
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, f"{name}: {result.outcome} {result.message}"
            record[name] = {"body": body(name), "cells": str(result.value)}
            print(name, result.value, flush=True)
    # A two-part date reads in the year the case ran in, which the replay holds the model at.
    OUT.write_text(json.dumps({"helper": _HELPER, "year": datetime.date.today().year, "cases": record}, indent=1)
                   + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
