"""Range.FillDown, FillUp, FillRight and FillLeft, as Excel's object model answers them.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills a few cells, formats some and fills a range one way, then a dump of
A1:E6 -- each cell's formula, number format, bold and fill colour -- and
the used range. The layouts ask what a fill copies (values, formulas with
each kind of reference, formats, blanks over what was there), where a
one-row or one-column range takes its source from, what several areas
and merged cells do, and what the method returns. A read records the
value's type and text, and E<number> for an error, as the other probes do.

    python scripts/measure_fill.py

writes tests/fixtures/fill.json, which tests/test_excel_fill.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "fill.json"

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

Private Function Dump(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:E6").Cells
        out = out & c.Formula & "~" & c.NumberFormat & "~" & c.Font.Bold & "~" & c.Interior.ColorIndex & "/"
    Next
    Dump = out & ws.UsedRange.Address
End Function
'''

#: name -> setup; each setup works on ws, the layout's own sheet, and ends with the fill it asks about.
LAYOUTS: dict[str, str] = {
    "down_block": 'ws.Range("A1").Value = 1\nws.Range("B1").Formula = "=A1*2"\nws.Range("C1").Value = "x"\n'
                  'ws.Range("D1").Value = 3\nws.Range("D1").NumberFormat = "0.00"\nws.Range("D1").Font.Bold = True\n'
                  'ws.Range("A3").Value = "old"\nws.Range("E4").Value = 9\nv = ws.Range("A1:E5").FillDown',
    "down_one_row": 'ws.Range("A1").Value = 1\nws.Range("B1").Formula = "=A1*2"\nv = ws.Range("A2:B2").FillDown',
    "down_one_cell": 'ws.Range("B1").Value = 5\nv = ws.Range("B2").FillDown',
    "down_top_row": 'ws.Range("A1").Value = 1\nv = ws.Range("A1:B1").FillDown',
    "down_references": 'ws.Range("A1").Formula = "=$C$1+C$1+$C1+C1"\nws.Range("C1:C4").Value = 1\n'
                       'ws.Range("B2:B3").Value = 9\nv = ws.Range("A1:B3").FillDown',
    "down_formats": 'ws.Range("A1").Value = 1\nws.Range("A1").Interior.ColorIndex = 6\n'
                    'ws.Range("A1").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                    'ws.Range("A2:A3").Interior.ColorIndex = 3\nv = ws.Range("A1:A3").FillDown',
    "up_block": 'ws.Range("A5").Value = 1\nws.Range("B5").Formula = "=A5*2"\nws.Range("A1").Value = "old"\n'
                'v = ws.Range("A1:B5").FillUp',
    "up_one_row": 'ws.Range("A3").Value = 1\nws.Range("B3").Formula = "=A3*2"\nv = ws.Range("A2:B2").FillUp',
    "right_block": 'ws.Range("A1").Value = 1\nws.Range("A2").Formula = "=A1*2"\nws.Range("C2").Value = "old"\n'
                   'v = ws.Range("A1:D2").FillRight',
    "right_one_column": 'ws.Range("A1").Value = 1\nws.Range("A2").Formula = "=A1*2"\nv = ws.Range("B1:B2").FillRight',
    "left_block": 'ws.Range("D1").Value = 1\nws.Range("D2").Formula = "=D1*2"\nv = ws.Range("A1:D2").FillLeft',
    "left_one_column": 'ws.Range("C1").Value = 1\nws.Range("C2").Formula = "=C1*2"\nv = ws.Range("B1:B2").FillLeft',
    "several_areas": 'ws.Range("A1").Value = 1\nws.Range("C1").Value = 2\nv = ws.Range("A1:A3,C1:C3").FillDown',
    "merged_source": 'ws.Range("A1:B1").Merge\nws.Range("A1").Value = 1\nv = ws.Range("A1:B3").FillDown',
    "text_is_not_a_series": 'ws.Range("A1").Value = "Item 1"\nws.Range("B1").Value = #1/2/2020#\n'
                            'v = ws.Range("A1:B3").FillDown',
    "prefix_and_error": 'ws.Range("A1").Value = "\'5"\nws.Range("B1").Formula = "=1/0"\nv = ws.Range("A1:B2").FillDown',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String, v As Variant, out As String",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             'out = failed & Show(v) & ";" & Dump(ws)', f"Case{index} = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup))
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "layouts": [{"name": name, "setup": setup, "answers": answer}
                                            for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
