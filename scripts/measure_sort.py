"""Range.Sort and Worksheet.Sort, as Excel's object model answers them.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills a small table and sorts it one way or another, then a dump of
A1:E7 -- each cell's formula, number format and bold -- with the used
range and what the sort returned or the error it raised. The layouts ask
what moves with a row, how formulas follow their rows, how a header is
guessed, what a single cell sorts, how keys, orders, orientation and
case combine, what text sorted as numbers does, and how the Sort object
a recorded macro uses does the same things. The order text sorts in is
scripts/measure_sort_order.py's.

    python scripts/measure_sort.py

writes tests/fixtures/sort.json, which tests/test_excel_sort.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "sort.json"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "Object"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:E7").Cells
        out = out & c.Formula & "~" & c.NumberFormat & "~" & c.Font.Bold & "<>"
    Next
    Dump = out & ws.UsedRange.Address
End Function
'''

#: A table with a header row: names, scores and a formula on each row.
TABLE = ('ws.Range("A1:C1").Value = Array("Name", "Score", "Double")\n'
         'ws.Range("A2:B2").Value = Array("pear", 3)\nws.Range("A3:B3").Value = Array("Apple", 1)\n'
         'ws.Range("A4:B4").Value = Array("fig", 2)\nws.Range("A5:B5").Value = Array("apple", 5)\n'
         'ws.Range("C2:C5").Formula = "=B2*2"\nws.Range("B3").Font.Bold = True\nws.Range("B4").NumberFormat = "0.00"\n')
#: The same rows without a header.
ROWS = ('ws.Range("A1:B1").Value = Array("pear", 3)\nws.Range("A2:B2").Value = Array("Apple", 1)\n'
        'ws.Range("A3:B3").Value = Array("fig", 2)\nws.Range("A4:B4").Value = Array("apple", 5)\n')

def column(letter: str, *values: object) -> str:
    """Writes of ``values`` down a column from row 1, as VBA; None leaves a cell blank."""
    lines: list[str] = []
    for row, value in enumerate(values, start=1):
        if value is None:
            continue
        text = f'"{value}"' if isinstance(value, str) else ("True" if value is True else str(value))
        lines.append(f'ws.Range("{letter}{row}").Value = {text}')
    return "\n".join(lines) + "\n"


#: name -> setup; each works on ws, the layout's own sheet, and keeps a result in v.
LAYOUTS: dict[str, str] = {
    "by_name": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes)',
    "by_score_descending": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("B1"), Order1:=xlDescending, '
                                   'Header:=xlYes)',
    "key_as_text": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:="B1", Header:=xlYes)',
    "key_column": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Columns("B"), Header:=xlYes)',
    "match_case": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("A1"), Header:=xlYes, MatchCase:=True)',
    "header_no": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("B1"), Header:=xlNo)',
    "header_guess_text_over_numbers": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("B1"), Header:=xlGuess)',
    "header_guess_omitted": TABLE + 'v = ws.Range("A1:C5").Sort(Key1:=ws.Range("B1"))',
    "header_guess_all_text": ROWS + 'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_bold": ROWS + 'ws.Range("A1:B1").Font.Bold = True\n'
                                'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_numbers": column("A", 4, 2, 3, 1) + 'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_number_over_text": column("A", 9, "b", "a", "c") +
                                     'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_text_over_text_numbers": column("A", "Key", "b", "a", "c") + column("B", "Value", 2, 1, 3) +
                                           'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_italic": ROWS + 'ws.Range("A1").Font.Italic = True\n'
                                  'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_fill": ROWS + 'ws.Range("B1").Interior.ColorIndex = 6\n'
                                'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_bold_elsewhere": ROWS + 'ws.Range("A3").Font.Bold = True\n'
                                          'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_number_format": column("A", 5, 3, 4, 1) + 'ws.Range("A1").NumberFormat = "0.00"\n'
                                  'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_date_over_numbers": column("A", "1/2/2020", 3, 4, 1) +
                                      'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_blank_first": column("A", None, "b", "a", "c") + column("B", None, 2, 1, 3) +
                                'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_text_over_blank": column("A", "h", None, "b", "a") +
                                    'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_caps": column("A", "NAME", "b", "a", "c") +
                         'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_formula": column("A", 4, 2, 3) + 'ws.Range("A4").Formula = "=1"\nws.Range("A1").Formula = "=9"\n'
                            'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_boolean": column("A", True, 2, 3, 1) +
                            'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_second_row_text": column("A", "h", 2, "x", 1) +
                                    'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_border": ROWS + 'ws.Range("A1").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                                  'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_alignment": ROWS + 'ws.Range("A1").HorizontalAlignment = xlCenter\n'
                                     'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_font_color": ROWS + 'ws.Range("A1").Font.Color = RGB(255, 0, 0)\n'
                                      'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_font_size": ROWS + 'ws.Range("A1").Font.Size = 14\n'
                                     'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_underline": ROWS + 'ws.Range("A1").Font.Underline = xlUnderlineStyleSingle\n'
                                     'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_font_name": ROWS + 'ws.Range("A1").Font.Name = "Courier New"\n'
                                     'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_error": column("A", "#N/A", 2, 3, 1) +
                          'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_partly_blank": column("A", None, "b", "a", "c") + column("B", "V", 2, 1, 3) +
                                 'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_caps_both": column("A", "NAME", "B", "A", "C") +
                              'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_caps_digits": column("A", "ID1", "b", "a", "c") +
                                'v = ws.Range("A1:A4").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "header_guess_one_row": column("A", "x") + column("B", 5) +
                            'v = ws.Range("A1:B1").Sort(Key1:=ws.Range("A1"), Header:=xlGuess)',
    "text_as_numbers_forms":column("A", "'$5", "'5%", "'1,000", "' 7", "'(3)", "'1/2", "'TRUE", "'6") +
                             'v = ws.Range("A1:A8").Sort(Key1:=ws.Range("A1"), Header:=xlNo, '
                             'DataOption1:=xlSortTextAsNumbers)',
    "sort_object_guess": TABLE + 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B2:B5")\n'
                         'ws.Sort.SetRange ws.Range("A1:C5")\nws.Sort.Header = xlGuess\nws.Sort.Apply\n'
                         'v = ws.Sort.Header',
    "sort_object_no_header": TABLE + 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B1:B5")\n'
                             'ws.Sort.SetRange ws.Range("A1:C5")\nws.Sort.Header = xlNo\nws.Sort.Apply\nv = ws.Sort.Header',
    "rows_no_header": ROWS +'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("B1"), Order1:=xlDescending, Header:=xlNo)',
    "two_keys": column("A", "b", "a", "b", "a", "c") + column("B", 2, 9, 1, 3, 0) +
                'v = ws.Range("A1:B5").Sort(Key1:=ws.Range("A1"), Order1:=xlAscending, '
                'Key2:=ws.Range("B1"), Order2:=xlDescending, Header:=xlNo)',
    "three_keys": column("A", 1, 1, 1, 2, 1, 2) + column("B", "x", "y", "x", "x", "y", "x") +
                  column("C", 3, 2, 1, 6, 5, 4) +
                  'v = ws.Range("A1:C6").Sort(Key1:=ws.Range("A1"), Key2:=ws.Range("B1"), Order2:=xlDescending, '
                  'Key3:=ws.Range("C1"), Header:=xlNo)',
    "single_cell": ROWS + 'ws.Range("D1").Value = 1\nv = ws.Range("A2").Sort(Key1:=ws.Range("A2"), Header:=xlNo)',
    "single_cell_region": ROWS + 'v = ws.Range("B3").Sort(Key1:=ws.Range("B3"), Header:=xlNo)',
    "orientation_rows": 'ws.Range("A1:D1").Value = Array("d", "b", "a", "c")\n'
                        'ws.Range("A2:D2").Value = Array(4, 2, 1, 3)\n'
                        'v = ws.Range("A1:D2").Sort(Key1:=ws.Range("A1"), Orientation:=xlSortRows, Header:=xlNo)',
    "formulas_follow": column("A", 3, 1, 4, 2) +
                       'ws.Range("E1").Value = 100\nws.Range("B1").Formula = "=A1*10"\nws.Range("B2").Formula = "=A1+A2"\n'
                       'ws.Range("B3").Formula = "=$A$1+A3"\nws.Range("B4").Formula = "=E1+A4"\n'
                       'ws.Range("C1").Formula = "=SUM(A1:A4)"\n'
                       'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("A1"), Header:=xlNo)',
    "blanks_last": column("A", 3, None, 1, None, 2) + column("B", "c", "x", "a", "y", "b") +
                   'v = ws.Range("A1:B5").Sort(Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo)',
    "text_as_numbers": 'ws.Range("A1:A5").NumberFormat = "@"\n' + column("A", "10", "9", "100", "b", "2") +
                       'v = ws.Range("A1:A5").Sort(Key1:=ws.Range("A1"), Header:=xlNo, DataOption1:=xlSortTextAsNumbers)',
    "text_as_numbers_mixed": column("A", "'10", 9, "'100", "b", 2, "'1e1") +
                             'v = ws.Range("A1:A6").Sort(Key1:=ws.Range("A1"), Header:=xlNo, '
                             'DataOption1:=xlSortTextAsNumbers)',
    "text_as_text": 'ws.Range("A1:A5").NumberFormat = "@"\n' + column("A", "10", "9", "100", "b", "2") +
                    'v = ws.Range("A1:A5").Sort(Key1:=ws.Range("A1"), Header:=xlNo)',
    "mixed_types": column("A", True, "b", 2, "#N/A", "a", 1) +
                   'v = ws.Range("A1:A6").Sort(Key1:=ws.Range("A1"), Header:=xlNo)',
    "mixed_types_descending": column("A", True, "b", 2, "#N/A", "a", 1, None, "#DIV/0!") +
                              'v = ws.Range("A1:A8").Sort(Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo)',
    "key_outside": ROWS + 'v = ws.Range("A1:B4").Sort(Key1:=ws.Range("D1"), Header:=xlNo)',
    "no_key": ROWS + 'v = ws.Range("A1:B4").Sort(Header:=xlNo)',
    "sort_object": TABLE + 'ws.Sort.SortFields.Clear\n'
                   'ws.Sort.SortFields.Add Key:=ws.Range("B2:B5"), SortOn:=xlSortOnValues, Order:=xlDescending, '
                   'DataOption:=xlSortNormal\nws.Sort.SetRange ws.Range("A1:C5")\nws.Sort.Header = xlYes\n'
                   'ws.Sort.MatchCase = False\nws.Sort.Orientation = xlTopToBottom\nws.Sort.Apply\n'
                   'v = ws.Sort.SortFields.Count',
    "sort_object_add2": TABLE + 'ws.Sort.SortFields.Clear\n'
                        'ws.Sort.SortFields.Add2 Key:=ws.Range("A2:A5"), SortOn:=xlSortOnValues, Order:=xlAscending\n'
                        'With ws.Sort\n.SetRange ws.Range("A1:C5")\n.Header = xlYes\n.MatchCase = True\n.Apply\n'
                        'End With\nv = ws.Sort.MatchCase',
    "sort_object_two_fields": column("A", "b", "a", "b", "a", "c") + column("B", 2, 9, 1, 3, 0) +
                              'ws.Sort.SortFields.Clear\n'
                              'ws.Sort.SortFields.Add Key:=ws.Range("A1:A5")\n'
                              'ws.Sort.SortFields.Add Key:=ws.Range("B1:B5"), Order:=xlDescending\n'
                              'ws.Sort.SetRange ws.Range("A1:B5")\nws.Sort.Header = xlNo\nws.Sort.Apply\n'
                              'v = ws.Sort.SortFields.Count',
    "sort_object_fields_kept": ROWS + 'ws.Sort.SortFields.Clear\nws.Sort.SortFields.Add Key:=ws.Range("B1:B4")\n'
                               'ws.Sort.SetRange ws.Range("A1:B4")\nws.Sort.Apply\nv = ws.Sort.SortFields.Count',
    "sort_object_header_default": TABLE + 'ws.Sort.SortFields.Clear\n'
                                  'ws.Sort.SortFields.Add Key:=ws.Range("B1:B5")\n'
                                  'ws.Sort.SetRange ws.Range("A1:C5")\nws.Sort.Apply\nv = ws.Sort.Header',
}


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
        head, _, dump = answer.partition(";")
        parts = dump.split("<>")
        cells = [f"{'ABCDE'[i % 5]}{i // 5 + 1}={cell}" for i, cell in enumerate(parts[:35])
                 if not cell.startswith("~General~False")]
        print(f"{name}: {head} | {' | '.join(cells)} | used {parts[35]}")


if __name__ == "__main__":
    main()
