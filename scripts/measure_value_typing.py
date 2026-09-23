"""What a string written through Range.Value becomes in a cell, as Excel types it.

Writing a string to a cell is typing it: "5" becomes the number, "TRUE"
the Boolean, "1/2/2020" a date, and a leading apostrophe keeps the rest
as text. Each probe writes one string into a cell of a fresh sheet and
reads back what the cell holds: the value's type and text, what it shows,
its formula, its prefix character and number format, and whether it is
empty. A read records the value's type and text, and E<number> for an
error, as the other probes do.

Four more sections ask what the rules leave open: how a number reads, and
shows, through each of a set of formats; what typing does to a cell that
already has a format, written through Value, Value2 or Formula; what
repeated writes to one cell leave behind, the prefix character and the
format included; and what an array written to a block types.

    python scripts/measure_value_typing.py

writes tests/fixtures/value_typing.json, which tests/test_excel_value_typing.py
replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "value_typing.json"

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

#: The strings written, as VBA expressions.
TEXTS = [
    '""', '" "', '"  "', 'vbTab', 'Chr(160)', '" a"', '"a "', '" 5"', '"5 "', '" 5 "', '"5"', '"-5"', '"+5"', '"5-"',
    '"(5)"', '"1,000"', '"1,000.5"', '"1 000"', '"5%"', '" 5%"', '"5 %"', '"$5"', '"$1,000.50"', '"-$5"', '"1e3"',
    '"1E+03"', '"1e400"', '".5"', '"5."', '"0005"', '"5.50"', '"1/2"', '"1/2/2020"', '"2020-01-02"', '"Jan 2, 2020"',
    '"2 Jan 2020"', '"12:30"', '"12:30:45"', '"1:2:3"', '"12:30 PM"', '"1/2/2020 12:30"', '"TRUE"', '"true"',
    '" TRUE"', '"FALSE "', '"Yes"', '"#N/A"', '"#DIV/0!"', '"\'"', '"\'5"', '"\'\'5"', '"\'abc"', '"\' 5"',
    '"\'TRUE"', '"\'=1+1"', '"=1+1"', '" =1+1"', '"a" & vbLf & "b"', '"1 1/2"', '"5 1/4"', '"- 5"', '"5E"',
    '"&H10"', '"0x10"', '"1.2.3"', '"12/31"', '"31/12/2020"', '"13/1/2020"', '"2/29/2021"', '"1-2-2020"',
    '"5.5%"', '"-5%"', '"(5%)"', '"$-5"', '"($5)"', '"1,00"', '"1,0000"', '"1,000,000"', '",5"', '"1,000%"',
    '"$1e3"', '"1.5e3"', '"1e-3"', '"1/4"', '"3/16"', '"1 3/16"', '"0 1/2"', '"-1 1/2"', '"#n/a"', '"#VALUE!"',
    '"#REF!"', '"#NAME?"', '"#NUM!"', '"#NULL!"', '"#SPILL!"', '"12:30:45 AM"', '"25:00"', '"1:60"', '"12:30am"',
    '"12 PM"', '"2020/1/2"', '"1/2/20"', '"1/2/1920"', '"1/2/29"', '"1/2/30"', '"Jan 2020"', '"January 2"',
    '"2-Jan"', '"2-Jan-2020"', '"Mon, Jan 2, 2020"', '"5 5"', '"$ 5"', '"5$"', '"1.5.2020"',
    '"1/2020"', '"1-2"', '"Jan-2020"', '"2 January"', '"January 2020"', '"Jan 2 2020"', '"2020-1-2 12:30"',
    '"1/2/2020 1:2:3 PM"', '"12:30:45.5"', '"$5.5"', '"$5%"', '"1,234,5"', '"1234,567"', '"$.5"', '"-.5"',
    '"5 1/10"', '"1/0"', '"1 1/0"', '"TRUE "', '"#N/A "', '"1 2/3"',
    # More than fifteen significant digits.
    '"123456789012345678"', '"1234567890123456789"', '"0.12345678901234567"', '"12345678901234.5678"',
    '"1.99999999999999999"', '"99999999999999999"', '"-123456789012345678"', '"1.23456789012345678E+20"',
    # More than one space where typing takes one.
    '"1/2/2020  12:30"', '"1/2/2020   1:00 PM"', '"12:30  PM"', '"1/2/2020 12:30  PM"', '"Jan  2, 2020"',
    '"Jan 2,  2020"', '"2  Jan  2020"', '"1  1/2"', '"$  5"', '"-  5"', '"5  %"', '"Jan  2020"', '"12  PM"',
    # Values a macro computes rather than types.
    "#1/2/2020#", "#1/2/2020 12:30:00 PM#", "#12:30:00 PM#", "CCur(5.5)", "CCur(1000)", "CDec(1.5)", "CInt(5)",
    "CSng(1.5)", "True", "CVErr(2042)",
]

READS = ["TypeName(c.Value)", "Show(c.Value)", "c.Text", "c.Formula", "c.PrefixCharacter", "c.NumberFormat",
         "IsEmpty(c.Value)", "c.HasFormula"]

#: Number formats a number is read through: which make Value a Date or a Currency, and what Text shows.
FORMATS = ["General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "@",
           "$#,##0", "$#,##0_);[Red]($#,##0)", "$#,##0.00_);($#,##0.00)", '_($* #,##0.00_);_($* (#,##0.00);'
           '_($* "-"??_);_(@_)', '"$"0', "#,##0 [$USD]", "m/d/yyyy", "d-mmm", "mmm-yy",
           "d-mmm-yy", "yyyy", "dddd", "mmmm", "m", "mm", "mmm", "d", "h", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm",
           "[mm]:ss", "h:mm AM/PM", "m/d/yyyy h:mm", "s", "[$-409]h:mm:ss AM/PM", "0.00;[Red]-0.00", '0 "units"']
FORMAT_READS = ["TypeName(c.Value)", "Show(c.Value)", "Show(c.Value2)", "c.Text"]
#: The last five are more than a Date or a Currency holds, one way or the other.
FORMAT_VALUES = ["45000.5625", "-1234.5", "0.25", "-1", "-657435", "2958466", "1E+20", "922337203685478"]

#: Formats a cell holds before a string is typed into it, and what is typed.
PRESET_FORMATS = ["General", "@", "0.00", "#,##0", "0%", "$#,##0.00", "0.00E+00", "# ?/?", "m/d/yyyy", "h:mm"]
PRESET_TEXTS = ['"5"', '"5%"', '"$5"', '"1,000"', '"1e3"', '"1 1/2"', '"1/2/2020"', '"12:30"', '"TRUE"', '"\'5"',
                '"#N/A"', '"abc"', "#1/2/2020#", "CCur(5.5)", "5", "True"]
#: Which property the typing goes through; Value2 and Formula only into a General cell and a Text one.
PRESET_WRITERS = {"Value": PRESET_FORMATS, "Value2": ["General", "@"], "Formula": ["General", "@"]}
PRESET_READS = ["TypeName(c.Value)", "Show(c.Value)", "Show(c.Value2)", "c.Text", "c.Formula", "c.PrefixCharacter",
                "c.NumberFormat"]

#: Writes to one cell in turn, read after each; ClearContents clears the cell's contents between.
SEQUENCES = [
    ['"\'5"', '"abc"', '"\'x"', "5", '"\'y"', '"=1+1"', '"\'z"', "ClearContents", '"q"'],
    ['"5%"', '"7"', '"$5"', '"1/2/2020"', '"12:30"', '"1e3"', '"1 1/2"', '"TRUE"', '"abc"', '"5"'],
    ['"1/2/2020"', '"12:30"', '"1/2/2020 12:30"', '"5"', '"Jan 2020"', '"2-Jan"'],
    ['"$5"', '"$5.5"', '"5%"', '"1,000.5"', '"1e3"'],
    ['"12:30"', '"1:2:3"', '"25:00"', '"12:30 PM"', '"12:30:45.5"'],
    ['"1 1/2"', '"1 3/16"', '"0 1/2"', '"5%"'],
    ["#1/2/2020#", "CCur(5.5)", "#12:30:00 PM#", "5"],
]
SEQUENCE_READS = ["Show(c.Value)", "c.Text", "c.NumberFormat", "c.PrefixCharacter"]

#: Arrays written to a block at once, and a single string written to a block.
ARRAYS = ['ws.Range("B2:E2").Value = Array("5%", "1/2/2020", "\'5", "$5")',
          'ws.Range("B2:E2").Value2 = Array("5%", "1/2/2020", "\'5", "$5")',
          'ws.Range("B2:E2").Value = Array(#1/2/2020#, CCur(5.5), "1,000", "1 1/2")',
          'ws.Range("B2:E2").Value = "5%"',
          # A block read of a cell its Date or Currency cannot hold, kept in D2 when it does not fail.
          'ws.Range("B2").Value = 2958466: ws.Range("B2").NumberFormat = "m/d/yyyy": '
          'ws.Range("D2").Value = TypeName(ws.Range("B2:C2").Value)',
          'ws.Range("B2").Value = 2958466: ws.Range("B2").NumberFormat = "m/d/yyyy": '
          'ws.Range("D2").Value = TypeName(ws.Range("B2:C2").Value2)',
          'ws.Range("C2").Value = 1E+20: ws.Range("C2").NumberFormat = "$#,##0.00": '
          'ws.Range("D2").Value = TypeName(ws.Range("B2:C2").Value)']
ARRAY_READS = ["Show(c.Value)", "c.NumberFormat", "c.PrefixCharacter"]


#: Between the steps of one case; no answer holds it.
STEP = "^"


def read_lines(expressions: list[str]) -> list[str]:
    lines: list[str] = []
    for expression in expressions:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return lines


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def procedure(index: int, text: str) -> list[str]:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim out As String, v As Variant, c As Object",
             'Set c = ws.Range("B2")', "On Error Resume Next", "Err.Clear", f"c.Value = {text}",
             'If Err.Number <> 0 Then out = "S" & Err.Number & ";"']
    return [*lines, *read_lines(READS), "On Error GoTo 0", f"Case{index} = out", "End Function"]


def format_procedure(index: int, code: str) -> list[str]:
    lines = [f"Private Function Format{index}(ws As Object) As String", "Dim out As String, v As Variant, c As Object"]
    for value in FORMAT_VALUES:
        lines += ['Set c = ws.Range("B2")', "c.Clear", "On Error Resume Next", "Err.Clear", f"c.Value = {value}",
                  f"c.NumberFormat = {vba_text(code)}", 'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"',
                  *read_lines(FORMAT_READS), "On Error GoTo 0", f'out = out & "{STEP}"']
    return [*lines, f"Format{index} = out", "End Function"]


def presets() -> list[tuple[str, str, str]]:
    return [(writer, code, text) for writer, codes in PRESET_WRITERS.items() for code in codes for text in PRESET_TEXTS]


def preset_procedure(index: int, writer: str, code: str, text: str) -> list[str]:
    lines = [f"Private Function Preset{index}(ws As Object) As String", "Dim out As String, v As Variant, c As Object",
             'Set c = ws.Range("B2")', "On Error Resume Next", "Err.Clear", f"c.NumberFormat = {vba_text(code)}",
             f"c.{writer} = {text}", 'If Err.Number <> 0 Then out = "S" & Err.Number & ";"']
    return [*lines, *read_lines(PRESET_READS), "On Error GoTo 0", f"Preset{index} = out", "End Function"]


def sequence_procedure(index: int, steps: list[str]) -> list[str]:
    lines = [f"Private Function Sequence{index}(ws As Object) As String",
             "Dim out As String, v As Variant, c As Object", 'Set c = ws.Range("B2")', "On Error Resume Next"]
    for step in steps:
        lines += ["Err.Clear", "c.ClearContents" if step == "ClearContents" else f"c.Value = {step}",
                  'If Err.Number <> 0 Then out = out & "S" & Err.Number & ";"', *read_lines(SEQUENCE_READS),
                  f'out = out & "{STEP}"']
    return [*lines, "On Error GoTo 0", f"Sequence{index} = out", "End Function"]


def array_procedure(index: int, write: str) -> list[str]:
    lines = [f"Private Function Array{index}(ws As Object) As String", "Dim out As String, v As Variant, c As Object",
             "On Error Resume Next", "Err.Clear", write, 'If Err.Number <> 0 Then out = "S" & Err.Number & ";"']
    for reference in ("B2", "C2", "D2", "E2"):
        lines += [f'Set c = ws.Range("{reference}")', *read_lines(ARRAY_READS), f'out = out & "{STEP}"']
    return [*lines, "On Error GoTo 0", f"Array{index} = out", "End Function"]


def steps(answer: str) -> list[str]:
    return [part for part in answer.split(STEP) if part]


def main() -> None:
    # Column B is wide enough that no answer's Text overflows into ####.
    head = [HELPER, "Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
            "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Columns("B:E").ColumnWidth = 40']
    body: list[str] = []
    for index, text in enumerate(TEXTS):
        head += ["ws.Cells.Clear", f'out = out & Case{index}(ws) & "|"']
        body += procedure(index, text)
    head.append('out = out & "~FORMATS~"')
    for index, code in enumerate(FORMATS):
        head.append(f'out = out & Format{index}(ws) & "|"')
        body += format_procedure(index, code)
    head.append('out = out & "~PRESETS~"')
    for index, (writer, code, text) in enumerate(presets()):
        head += ["ws.Cells.Clear", f'out = out & Preset{index}(ws) & "|"']
        body += preset_procedure(index, writer, code, text)
    head.append('out = out & "~SEQUENCES~"')
    for index, sequence in enumerate(SEQUENCES):
        head += ["ws.Cells.Clear", f'out = out & Sequence{index}(ws) & "|"']
        body += sequence_procedure(index, sequence)
    head.append('out = out & "~ARRAYS~"')
    for index, write in enumerate(ARRAYS):
        head += ["ws.Cells.Clear", f'out = out & Array{index}(ws) & "|"']
        body += array_procedure(index, write)
    head += ["wb.Close False", "Probe = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(head + body) + "\n", "Probe", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    typed, rest = str(result.value).split("~FORMATS~", 1)
    formatted, rest = rest.split("~PRESETS~", 1)
    preset, rest = rest.split("~SEQUENCES~", 1)
    sequenced, arrayed = rest.split("~ARRAYS~", 1)
    answers = typed.split("|")[: len(TEXTS)]
    shown = formatted.split("|")[: len(FORMATS)]
    preset_answers = preset.split("|")[: len(presets())]
    sequence_answers = sequenced.split("|")[: len(SEQUENCES)]
    array_answers = arrayed.split("|")[: len(ARRAYS)]
    # A date typed without a year falls in the year it was typed, which a replay has to know.
    record = {"helper": HELPER, "year": datetime.date.today().year, "reads": READS, "format_reads": FORMAT_READS,
              "format_values": FORMAT_VALUES,
              "preset_reads": PRESET_READS, "sequence_reads": SEQUENCE_READS, "array_reads": ARRAY_READS,
              "cases": [{"text": text, "answers": answer} for text, answer in zip(TEXTS, answers, strict=True)],
              "formats": [{"code": code, "answers": steps(answer)} for code, answer in zip(FORMATS, shown, strict=True)],
              "presets": [{"writer": writer, "code": code, "text": text, "answers": answer}
                          for (writer, code, text), answer in zip(presets(), preset_answers, strict=True)],
              "sequences": [{"steps": sequence, "answers": steps(answer)}
                            for sequence, answer in zip(SEQUENCES, sequence_answers, strict=True)],
              "arrays": [{"write": write, "answers": steps(answer)}
                         for write, answer in zip(ARRAYS, array_answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for text, answer in zip(TEXTS, answers, strict=True):
        print(f"{text:24} {answer}")
    for code, answer in zip(FORMATS, shown, strict=True):
        print(f"{code:40} {answer}")
    for (writer, code, text), answer in zip(presets(), preset_answers, strict=True):
        print(f"{writer:8} {code:12} {text:14} {answer}")
    for sequence, answer in zip(SEQUENCES, sequence_answers, strict=True):
        print(sequence, answer)
    for write, answer in zip(ARRAYS, array_answers, strict=True):
        print(write, answer)


if __name__ == "__main__":
    main()
