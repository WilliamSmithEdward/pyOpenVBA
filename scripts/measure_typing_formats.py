"""What typing into a cell does to the number format the cell already has, and to its prefix character.

A string typed into a General cell brings its own format -- "5%" 0%,
"1/2/2020" m/d/yyyy -- and scripts/measure_value_typing.py records which.
A cell with a format of its own keeps it for some typed values and takes
the typed one for others. The matrix types each value into a fresh cell
holding each of Excel's built-in formats and a few custom ones, and reads
the value's type, the number under it and the format the cell ends with;
the chains type one value and then another into the same cell, so the
first format is one typing gave. The sequences write to one cell in turn:
what keeps and what clears the apostrophe a cell was typed with, what a
Text cell makes of what is written to it, what Value2 and FormulaR1C1
type, and fractions typed alone. The last section reads back formats
Excel rewrites as it stores them.

    python scripts/measure_typing_formats.py

writes tests/fixtures/typing_formats.json, which tests/test_excel_value_typing.py
replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "typing_formats.json"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''

#: The formats a cell holds before the typing: Excel's built-in ones, then custom ones.
FORMATS = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "$#,##0_);($#,##0)", "$#,##0_);[Red]($#,##0)",
    "$#,##0.00_);($#,##0.00)", "$#,##0.00_);[Red]($#,##0.00)", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??",
    "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm",
    "#,##0_);(#,##0)", "#,##0_);[Red](#,##0)", "#,##0.00_);(#,##0.00)", "#,##0.00_);[Red](#,##0.00)",
    '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "mm:ss", "[h]:mm:ss", "mm:ss.0", "##0.0E+0", "@",
    "$#,##0.00", "$#,##0", "0.0", "0.000%", "yyyy-mm-dd", "dddd", "[$-409]h:mm:ss AM/PM", '0 "units"',
    "#,##0 [$USD]", "[Red]0.00", "0.00;-0.00", "[h]:mm", "d", "mmmm", "hh", "0.00_);(0.00)", "# ?/8", "yyyy",
    "[$-409]m/d/yyyy", "m/d/yy", "mm/dd/yyyy", "ss", '"$"#,##0.00', "[$$-409]#,##0.00", "0.00%;[Red]-0.00%",
]

#: What is typed, as VBA expressions: strings that bring a format, and values a macro computes.
TYPED = [
    '"5"', '"1,000"', '"1,000.5"', '"5%"', '"5.5%"', '"$5"', '"$5.5"', '"1e3"', '"1 1/2"', '"1 3/16"',
    '"1/2/2020"', '"1/2"', '"Jan 2020"', '"2-Jan-2020"', '"12:30"', '"12:30:45"', '"12:30 PM"', '"12:30:45 AM"',
    '"25:00"', '"12:30:45.5"', '"1/2/2020 12:30"', '"1:60"', "#1/2/2020#", "#12:30:00 PM#",
    "#1/2/2020 12:30:00 PM#", "CCur(5.5)", "5", '"abc"', '"TRUE"', '"13/4"', '"1/0"', '"3/4/5"',
]

#: Writes to one cell in turn, each followed by reads; a step that is not an assignment runs as it stands.
SEQUENCES = [
    ['"\'5"', '"7"'],
    ['"\'5"', "True"],
    ['"\'5"', "CVErr(2042)"],
    ['"\'5"', '"#N/A"'],
    ['"\'5"', "#1/2/2020#"],
    ['"\'5"', '"1/2/2020"'],
    ['"\'5"', '""'],
    ['"\'5"', "Empty"],
    ['"\'5"', '"abc"', "c.ClearContents", "5", '"q"'],
    ['"\'5"', '"abc"', "c.ClearContents", '"7"', '"q"'],
    ['"\'5"', "c.Clear", '"q"'],
    ['"\'5"', "c.ClearFormats", '"q"'],
    ['"\'5"', 'c.NumberFormat = "0.00"', '"q"', "5"],
    ['"\'5"', 'c.Formula = "abc"'],
    ['"\'5"', 'c.Value2 = "abc"'],
    ['"\'5"', 'c.Formula = "=1+1"', '"q"'],
    ['c.NumberFormat = "@"', '"\'5"', '"abc"', "5", '"q"'],
    ['c.NumberFormat = "@"', '"=1+1"', 'c.Formula = "=1+1"', 'c.FormulaR1C1 = "=1+1"', '"\'=1+1"', '""', '" 5 "'],
    ['c.NumberFormat = "0%"', '"\'5"', '"7"'],
    ['"\'5"', 'c.Formula = "=""a"""', '"q"'],
    # A Text cell keeps what is typed as it is typed, and turns what a macro computes into text.
    ['c.NumberFormat = "@"', '""', "Empty", '"5"', '""'],
    ['c.NumberFormat = "@"', "CCur(-5.5)"],
    ['c.NumberFormat = "@"', "#12:30:45 PM#", "#1/2/2020 12:30:45 PM#", "#1/2/1900#", "#1/2/1899#",
     "#1/2/1899 6:00:00 AM#", "CDec(1.5)", "CSng(1.5)", "CLng(5)", "CVErr(2007)"],
    ['c.NumberFormat = "@"', "c.Value2 = #1/2/2020#", "c.Value2 = CCur(5.5)", 'c.FormulaR1C1 = "5%"'],
    # Value2 turns a Date and a Currency into a Double, in an array as well.
    ['ws.Range("B2:C2").Value2 = Array(#1/2/2020#, CCur(5.5))', 'Set c = ws.Range("C2")'],
    ['ws.Range("B2:C2").Value = Array(#12:30:00 PM#, "\'5")', 'Set c = ws.Range("C2")'],
    ['c.FormulaR1C1 = "5%"', 'c.FormulaR1C1 = "1/2/2020"', 'c.FormulaR1C1 = "\'5"'],
    # A fraction typed alone, into cells that read one and cells that read a date.
    ['c.NumberFormat = "0%"', '"3/16"'],
    ['c.NumberFormat = "$#,##0_);($#,##0)"', '"3/16"'],
    ['c.NumberFormat = "0.00"', '"-1/2"', '"1-2"', '"1/2/3"', '"12/31"', '"1/13"', '"0/2"', '"1/2 "', '" 1/2"'],
    ['"-1/2"'],
    ['c.NumberFormat = "# ?/?"', '"3/16"', '"-1/2"'],
    ['c.NumberFormat = "h:mm"', '"1-2"'],
    # Early 1900, where Excel's calendar counts a 29 February that VBA's does not, and before it.
    ["#1/2/1899#", "#12/30/1899#", "#12/31/1899#", "#1/1/1900#", "#2/28/1900#", "#3/1/1900#", "#12:00:00 AM#"],
    ['"1/1/1900"', '"2/28/1900"', '"2/29/1900"', '"3/1/1900"', '"12/31/1899"', '"1/2/1899"', '"12/31/9999"',
     '"1/1/10000"', '"0:00"', '"24:00"', '"-1:00"', '"1:2:3:4"', '"9999:00"', '"10000:00"'],
    ['c.NumberFormat = "@"', "#12:00:00 AM#", "#1/2/2020 12:00:00 AM#", "#12/31/1899#", "#12/30/1899#"],
    # Fractions with longer denominators, and improper ones.
    ['"1 1/100"', '"1 12/345"', '"1 13/4"', '"0/2"', '"1 0/2"', '"1 1/2/3"'],
    ['c.NumberFormat = "0.00"', '"1/100"', '"1/2345"', '"12/345"', '"0/0"', '"-0/2"'],
    # How far each part of a typed time may run, each into a fresh cell.
    *[[f'"{text}"'] for text in ("0:9999", "0:10000", "0:0:9999", "0:0:10000", "1:60:00", "100:00", "13:00 PM",
                                 "0:30 AM", "12:60 PM", "1 13/4", "12:30:60", "1:2:3.25", "1:02 pm", "1:02PM",
                                 "12:30 P", "12:30 a", "1:2.5", "0:0:0.5", "1:00:00 AM", "25:00:00", "9999:59:59",
                                 "1:2:3.123456", "12:30:45.55")],
]
READS = ["Show(c.Value)", "Show(c.Value2)", "c.Text", "c.Formula", "c.NumberFormat", "c.PrefixCharacter",
         "c.HasFormula"]

STEP = "^"


def read_lines(expressions: list[str]) -> list[str]:
    lines: list[str] = []
    for expression in expressions:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return lines


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


#: The matrix and the chains as loops over the typed values, which keeps the module small enough to add.
#: Each typing reads back the value's type, the number under it and the format.
LOOPS = f'''Private Function Matrix(ws As Object, code As String, items As Variant) As String
    Dim out As String, c As Object, i As Long
    Set c = ws.Range("B2")
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        c.Clear
        Err.Clear
        c.NumberFormat = code
        c.Value = items(i)
        If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"
        out = out & TypeName(c.Value) & ";" & Show(c.Value2) & ";" & c.NumberFormat & "{STEP}"
    Next
    On Error GoTo 0
    Matrix = out
End Function

Private Function Chain(ws As Object, first As Variant, items As Variant) As String
    Dim out As String, c As Object, i As Long
    Set c = ws.Range("B2")
    On Error Resume Next
    For i = LBound(items) To UBound(items)
        c.Clear
        Err.Clear
        c.Value = first
        c.Value = items(i)
        If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"
        out = out & TypeName(c.Value) & ";" & Show(c.Value2) & ";" & c.NumberFormat & "{STEP}"
    Next
    On Error GoTo 0
    Chain = out
End Function
'''


def sequence_procedure(index: int, steps: list[str]) -> list[str]:
    lines = [f"Private Function Sequence{index}(ws As Object) As String",
             "Dim out As String, v As Variant, c As Object", 'Set c = ws.Range("B2")', "On Error Resume Next"]
    for step in steps:
        statement = step if step.startswith(("c.", "ws.", "Set ")) else f"c.Value = {step}"
        lines += ["Err.Clear", statement, 'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"',
                  *read_lines(READS), f'out = out & "{STEP}"']
    return [*lines, "On Error GoTo 0", f"Sequence{index} = out", "End Function"]


#: Formats set through NumberFormat, to see which Excel rewrites as it stores them.
NORMALIZED = ['"$"#,##0.00', '0"-"0', '0" "0', '"abc"0', "0\\-0", '#,##0"."00', "[red]0.00", "general", "0.00e+00",
              "M/D/YYYY", "h:MM", "MM:SS", "[H]:MM", "am/pm h", "h AM/PM", "h a/p", "0.00_)", '"$"0', "$0", "\\$0",
              "0%", "0 %", "[>100]0;0.00", "[Blue][<0]0;0", "dd/mm/yyyy", "yyyy/mm/dd", "@", "\\@", '"@"', "0;-0;;@"]


def normalized_procedure() -> list[str]:
    lines = ["Private Function Normalized(ws As Object) As String", "Dim out As String, c As Object",
             'Set c = ws.Range("B2")', "On Error Resume Next"]
    for code in NORMALIZED:
        lines += ["c.Clear", "Err.Clear", f"c.NumberFormat = {vba_text(code)}",
                  'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"', f'out = out & c.NumberFormat & "{STEP}"']
    return [*lines, "On Error GoTo 0", "Normalized = out", "End Function"]


def steps(answer: str) -> list[str]:
    return [part for part in answer.split(STEP) if part]


def main() -> None:
    head = [HELPER, LOOPS, "Public Function Probe() As String",
            "Dim wb As Object, ws As Object, out As String, items As Variant, i As Long",
            "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            "ws.Columns(2).ColumnWidth = 40", f"items = Array({', '.join(TYPED)})"]
    body: list[str] = []
    for code in FORMATS:
        head.append(f'out = out & Matrix(ws, {vba_text(code)}, items) & "|"')
    head += ['out = out & "~CHAINS~"', "For i = LBound(items) To UBound(items)",
             'out = out & Chain(ws, items(i), items) & "|"', "Next"]
    head.append('out = out & "~SEQUENCES~"')
    for index, sequence in enumerate(SEQUENCES):
        head += ["ws.Cells.Clear", f'out = out & Sequence{index}(ws) & "|"']
        body += sequence_procedure(index, sequence)
    head += ['out = out & "~NORMALIZED~" & Normalized(ws)', "wb.Close False", "Probe = out", "End Function"]
    body += normalized_procedure()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(head + body) + "\n", "Probe", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    matrix, rest = str(result.value).split("~CHAINS~", 1)
    chains, rest = rest.split("~SEQUENCES~", 1)
    sequenced, normalized = rest.split("~NORMALIZED~", 1)
    matrix_answers = matrix.split("|")[: len(FORMATS)]
    chain_answers = chains.split("|")[: len(TYPED)]
    sequence_answers = sequenced.split("|")[: len(SEQUENCES)]
    record = {"helper": HELPER, "year": datetime.date.today().year, "typed": TYPED, "reads": READS,
              "matrix": [{"code": code, "answers": steps(answer)}
                         for code, answer in zip(FORMATS, matrix_answers, strict=True)],
              "chains": [{"first": first, "answers": steps(answer)}
                         for first, answer in zip(TYPED, chain_answers, strict=True)],
              "sequences": [{"steps": sequence, "answers": steps(answer)}
                            for sequence, answer in zip(SEQUENCES, sequence_answers, strict=True)],
              "normalized": [{"code": code, "stored": stored}
                             for code, stored in zip(NORMALIZED, normalized.split(STEP)[: len(NORMALIZED)], strict=True)]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for code, answer in zip(FORMATS, matrix_answers, strict=True):
        print(f"{code:40} {answer}")
    for first, answer in zip(TYPED, chain_answers, strict=True):
        print(f"{first:24} {answer}")
    for sequence, answer in zip(SEQUENCES, sequence_answers, strict=True):
        print(sequence, answer)
    for entry in record["normalized"]:
        print(f"{entry['code']:20} {entry['stored']}")


if __name__ == "__main__":
    main()
