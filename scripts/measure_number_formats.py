"""What a number format shows: Range.Text and the TEXT function over a spread of codes and values.

Each format code is set on one cell in a column 60 characters wide, wide
enough that the width never decides what shows except where a format
fills it; every value is written to that cell in turn and read back
through Range.Text, and passed with the code to WorksheetFunction.Text.
The codes cover every kind a format has -- digit placeholders, thousands
and scaling commas, percent, exponents, fractions, currency and
accounting, sections and conditions, text, dates, times and elapsed
times -- and the values the edges that matter to each: signs, zero,
rounding halves, tiny and huge numbers, the 29 February 1900 Excel
counts, the last day it knows, and text.

    python scripts/measure_number_formats.py

writes tests/fixtures/number_formats.json, which
tests/test_excel_number_formats.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "number_formats.json"

CODES = [
    # digits, decimals, thousands and scaling
    "General", "0", "0.0", "0.00", "0.000", "#", "#.#", "#.##", "0.#", "#,##0", "#,##0.00", "#,##0.000", "0,",
    "0.0,,", '#,##0,"K"', "000", "00.00", "?.?", "??.??", "#.???", "0.0?", "#,###", "#,#", "0,0", "00,000.0",
    ".00", "0.", "#.", "#0.0#", "0.00_);(0.00)", "#,##0_);(#,##0)", "#,##0.00_);[Red](#,##0.00)",
    # sections and conditions
    "0;(0)", "0;-0;\"zero\"", "0;-0;;@", "0;;", ";;;", "0.00;;;", "[Red]0", "[Blue]0;[Red]-0",
    "[>=100]0;[<100]0.0", '[<0]"neg";"pos"', "[>100]\"big\";[>10]\"mid\";0", "0;\"neg\";\"zero\";\"text\"",
    "General;General;General", "[Red]General", "0.00;General",
    # percent and exponents
    "0%", "0.0%", "0.00%", "%0", "0.00E+00", "0.0E+0", "##0.0E+0", "0.00E-00", "0E+0", "#.##E+00", "00.00E+00",
    # fractions
    "# ?/?", "# ??/??", "# ???/???", "?/?", "??/??", "# ?/2", "# ?/4", "# ?/8", "# ??/16", "# ?/10", "# ??/100",
    "0 ?/?", "#/#", "0/0",
    # currency and accounting
    "$#,##0", "$#,##0.00", "$#,##0_);($#,##0)", "$#,##0.00_);[Red]($#,##0.00)",
    '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "[$$-409]#,##0.00", "#,##0 [$USD]", "#,##0.00 [$EUR]",
    # text and literals
    "@", '"x"@', '@" units"', '0;0;0;"text: "@', '0 "units"', '"$"0', "0\\x", "(0)", "0 %", "0_)", "_(0_)",
    '"a""b"0', "\\@0", "0-0", "000-00-0000", "(###) ###-####",
    # dates
    "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "m/d/yy", "mm/dd/yyyy", "dd/mm/yyyy", "yyyy-mm-dd", "d", "dd",
    "ddd", "dddd", "m", "mm", "mmm", "mmmm", "mmmmm", "yy", "yyyy", "y", "yyy", "mmmm d, yyyy",
    "dddd, mmmm dd, yyyy", "[$-409]mmmm d, yyyy", "d/m/yyyy", "m/d",
    # times
    "h", "hh", "h:mm", "hh:mm", "h:mm:ss", "hh:mm:ss", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm A/P", "h:mm a/p",
    "hh:mm AM/PM", "mm:ss", "mm:ss.0", "mm:ss.00", "mm:ss.000", "[h]:mm:ss", "[h]:mm", "[mm]:ss", "[ss]",
    "[ss].00", "h:mm:ss.0", "m/d/yyyy h:mm", "m/d/yyyy h:mm:ss AM/PM", "s", "ss", "[hh]:mm", "h AM/PM",
    "AM/PM h", "yyyy-mm-dd hh:mm:ss", "[h]:mm:ss.00",
    # fills, which the width decides
    "* 0", "0*-", "*x0.00",
]

#: The values each code is shown with, as VBA expressions.
VALUES = ["0", "1", "-1", "0.5", "-0.5", "1234.5678", "-1234.5678", "0.001", "1234567.891", "45000.5625", "0.25",
          "1E+15", "1E-10", "99.995", "0.005", "1.5", "2.5", "-2.5", "60", "61", "2958465", "2958466", "12.3456",
          '"abc"', "True", "-0.4", "-0.004", "-0.01"]

STEP, ITEM = "<^>", "<|>"

#: One code through every value: Range.Text of the cell, then the TEXT function, for each.
SHOW = f'''Private Function Shown(ws As Object, code As String, values As Variant) As String
    Dim out As String, c As Object, v As Variant, i As Long
    Set c = ws.Range("A1")
    On Error Resume Next
    c.Clear
    Err.Clear
    c.NumberFormat = code
    If Err.Number <> 0 Then out = out & "S" & Err.Number & "{ITEM}"
    For i = LBound(values) To UBound(values)
        Err.Clear
        c.Value = values(i)
        v = Empty
        v = c.Text
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "{ITEM}" Else out = out & v & "{ITEM}"
        Err.Clear
        v = Empty
        v = Application.WorksheetFunction.Text(values(i), code)
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "{ITEM}" Else out = out & v & "{ITEM}"
    Next
    On Error GoTo 0
    Shown = out & "{STEP}"
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def main() -> None:
    lines = [SHOW, "Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Dim values As Variant", "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "ws.Columns(1).ColumnWidth = 60", f"values = Array({', '.join(VALUES)})"]
    lines += [f"out = out & Shown(ws, {vba_text(code)}, values)" for code in CODES]
    lines += ["wb.Close False", "Probe = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(lines) + "\n", "Probe", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    shown = str(result.value).split(STEP)[: len(CODES)]
    record: dict[str, object] = {"values": VALUES, "codes": []}
    codes: list[dict[str, object]] = []
    for code, answer in zip(CODES, shown, strict=True):
        items = answer.split(ITEM)[:-1]
        failed = items[0] if items and items[0].startswith("S") and len(items) > 2 * len(VALUES) else ""
        items = items[1:] if failed else items
        codes.append({"code": code, "failed": failed, "text": items[0::2], "function": items[1::2]})
    record["codes"] = codes
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for entry in codes:
        print(f"{entry['code']:40} {entry['failed']} {entry['text'][:6]}")


if __name__ == "__main__":
    main()
