"""Range.RemoveDuplicates, as Excel's object model answers it.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills a few cells and removes duplicates from them one way or another,
then a dump of A1:E8 -- each cell's formula, the type of its value, its
number format and bold -- with the used range and what the call returned
or the error it raised. The layouts ask which values count as the same
(case, text against numbers, formats, formulas, blanks, errors), which
columns are compared, how a header is kept or guessed, what moves when a
row goes and what stays, and which arguments Excel refuses.

    python scripts/measure_remove_duplicates.py

writes tests/fixtures/remove_duplicates.json, which
tests/test_excel_remove_duplicates.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "remove_duplicates.json"

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

Private Function Dump(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:E8").Cells
        out = out & c.Formula & "~" & TypeName(c.Value) & "~" & c.NumberFormat & "~" & c.Font.Bold & "<>"
    Next
    Dump = out & ws.UsedRange.Address
End Function
'''


def column(letter: str, *values: object) -> str:
    """Writes of ``values`` (VBA expressions, or None for a blank) down a column from row 1."""
    return "".join(f'ws.Range("{letter}{row}").Value = {value}\n'
                   for row, value in enumerate(values, start=1) if value is not None)


def remove(reference: str, arguments: str) -> str:
    return f'ws.Range("{reference}").RemoveDuplicates {arguments}'


PAIRS = column("A", "1", "1", "1", "2", "2") + column("B", '"a"', '"b"', '"a"', '"a"', '"a"') + \
    column("C", "10", "20", "30", "40", "50")

#: name -> setup; each works on ws, the layout's own sheet, and may keep a result in v.
LAYOUTS: dict[str, str] = {
    "basic": column("A", "1", "2", "1", "3", "2", "4") + remove("A1:A6", "Columns:=1, Header:=xlNo"),
    "header_yes": column("A", '"id"', "1", "2", "1") + remove("A1:A4", "Columns:=1, Header:=xlYes"),
    "header_yes_same": column("A", "1", "1", "2", "1") + remove("A1:A4", "Columns:=1, Header:=xlYes"),
    "header_no_text": column("A", '"id"', '"id"', "2") + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "header_guess_text": column("A", '"id"', "1", "1", "2") + remove("A1:A4", "Columns:=1, Header:=xlGuess"),
    "header_guess_numbers": column("A", "5", "1", "1", "5") + remove("A1:A4", "Columns:=1, Header:=xlGuess"),
    "header_guess_texts": column("A", '"x"', '"y"', '"y"', '"x"') + remove("A1:A4", "Columns:=1, Header:=xlGuess"),
    "header_missing": column("A", '"x"', '"x"', "1") + remove("A1:A3", "Columns:=1"),
    "case": column("A", '"a"', '"A"', '"b"', '"B"', '"b"') + remove("A1:A5", "Columns:=1, Header:=xlNo"),
    "number_text": column("A", "1", '"\'1"', "1", '"\'1"') + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "formats": column("A", "1", "1", "1", "2") + 'ws.Range("A2").NumberFormat = "0.00"\n' +
               'ws.Range("A3").Font.Bold = True\n' + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "date_number": column("A", "43831", "#1/1/2020#", "43831") + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "floats": column("A", "0.1 + 0.2", "0.3", "0.30000000000000004") + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "near_one": column("A", "1", "1 + 2 ^ -52", "1") + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "percent": column("A", "0.5", "0.5") + 'ws.Range("A1").NumberFormat = "0%"\n' +
               remove("A1:A2", "Columns:=1, Header:=xlNo"),
    "two_columns": PAIRS + remove("A1:C5", "Columns:=Array(1, 2), Header:=xlNo"),
    "second_column": PAIRS + remove("A1:C5", "Columns:=2, Header:=xlNo"),
    "first_column": PAIRS + remove("A1:C5", "Columns:=1, Header:=xlNo"),
    "all_columns": PAIRS + remove("A1:C5", "Columns:=Array(1, 2, 3), Header:=xlNo"),
    "narrow_range": PAIRS + remove("A1:B5", "Columns:=1, Header:=xlNo"),
    "blanks": column("A", None, "1", None, "1", "2") + remove("A1:A5", "Columns:=1, Header:=xlNo"),
    "blank_rows": column("A", "1", None, "2", None) + column("B", '"a"', None, '"b"', None) +
                  remove("A1:B4", "Columns:=Array(1, 2), Header:=xlNo"),
    "blank_one_column": column("A", "1", "1", "1") + column("B", '"a"', None, None) +
                        remove("A1:B3", "Columns:=Array(1, 2), Header:=xlNo"),
    "formulas": 'ws.Range("A1").Formula = "=1+0"\n' + column("A", None, "1") +
                'ws.Range("A3").Formula = "=0+1"\n' + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "formula_text": 'ws.Range("A1").Formula = "=""1"""\n' + column("A", None, "1", '"\'1"') +
                    remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "formulas_move": column("A", "1", "2", "1", "3") + 'ws.Range("B1:B4").Formula = "=A1*10"\n' +
                     column("D", "7") + 'ws.Range("C1:C4").Formula = "=$D$1+B1"\n' +
                     remove("A1:C4", "Columns:=1, Header:=xlNo"),
    "references_in": column("A", "1", "1", "2", "3") + 'ws.Range("E1").Formula = "=A4"\n' +
                     'ws.Range("E2").Formula = "=SUM(A1:A4)"\n' + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "errors": column("A", "CVErr(xlErrNA)", "CVErr(xlErrNA)", "CVErr(xlErrDiv0)") +
              remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "booleans": column("A", "True", "True", "False", '"TRUE"') + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "bold_moves": column("A", "1", "2", "1", "3") + 'ws.Range("A4").Font.Bold = True\n' +
                  'ws.Range("A4").NumberFormat = "0.00"\n' + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "bold_kept_row": column("A", "1", "2", "1", "3") + 'ws.Range("A3").Font.Bold = True\n' +
                     remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "outside_untouched": column("A", "1", "1", "2") + column("C", '"x"', '"y"', '"z"') +
                         remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "below_untouched": column("A", "1", "1", "2", None, "9") + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "trailing_space": column("A", '"a"', '"a "', '" a"', '"a"') + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "long_text": column("A", 'String(300, "x")', 'String(300, "x")', 'String(299, "x") & "y"') +
                 remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "accents": column("A", "ChrW(233)", "ChrW(201)", '"e"') + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "ligature": column("A", "ChrW(230)", '"ae"', "ChrW(223)", '"ss"') + remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "hyphen": column("A", '"co-op"', '"coop"', '"co op"') + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "columns_outside": PAIRS + remove("A1:B5", "Columns:=3, Header:=xlNo"),
    "columns_zero": PAIRS + remove("A1:B5", "Columns:=0, Header:=xlNo"),
    "columns_missing": PAIRS + remove("A1:B5", "Header:=xlNo"),
    "columns_text": PAIRS + remove("A1:B5", 'Columns:="A", Header:=xlNo'),
    "columns_repeated": PAIRS + remove("A1:C5", "Columns:=Array(1, 1), Header:=xlNo"),
    "columns_order": PAIRS + remove("A1:C5", "Columns:=Array(2, 1), Header:=xlNo"),
    "header_bad": PAIRS + remove("A1:B5", "Columns:=1, Header:=5"),
    "multi_area": column("A", "1", "1") + column("C", "1", "1") + remove("A1:A2,C1:C2", "Columns:=1, Header:=xlNo"),
    "whole_column": column("A", "1", "2", "1") + remove("A:A", "Columns:=1, Header:=xlNo"),
    "single_cell": column("A", "1", "1") + remove("A1", "Columns:=1, Header:=xlNo"),
    "return_value": column("A", "1", "1") + 'v = ws.Range("A1:A2").RemoveDuplicates(Columns:=1, Header:=xlNo)',
    "numbers_order": column("A", "3", "1", "3", "2", "1") + remove("A1:A5", "Columns:=1, Header:=xlNo"),
    "negative_zero": column("A", "0", "-0", '"0"') + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "big_numbers": column("A", "123456789012345#", "123456789012346#", "123456789012345#") +
                   remove("A1:A3", "Columns:=1, Header:=xlNo"),
}


def formatted(first: tuple[str, str], second: tuple[str, str]) -> str:
    """Two values (VBA expressions) under two number formats in A1:A2, then duplicates removed."""
    codes = [code.replace('"', '""') for code in (first[1], second[1])]
    return (f'ws.Range("A1").NumberFormat = "{codes[0]}"\nws.Range("A2").NumberFormat = "{codes[1]}"\n' +
            column("A", first[0], second[0]) + remove("A1:A2", "Columns:=1, Header:=xlNo"))


LAYOUTS.update({
    "key_general_zero": formatted(("1", "General"), ("1", "0")),
    "key_rounded_display": formatted(("1.5", "0"), ("2", "0")),
    "key_two_decimals": formatted(("1", "0.00"), ("1", "#,##0.00")),
    "key_general_decimal": formatted(("0.5", "General"), ("0.5", "0.0")),
    "key_date_formats": formatted(("#1/1/2020#", "m/d/yyyy"), ("#1/1/2020#", "mm/dd/yyyy")),
    "key_date_time": formatted(("43831.5", "m/d/yyyy"), ("43831", "m/d/yyyy")),
    "key_text_format": formatted(('"abc"', "General"), ('"abc"', "@")),
    "key_same_shown": formatted(("0.1 + 0.2", "0.0"), ("0.3", "0.0")),
    "key_currency": formatted(("1", "$#,##0.00"), ("1", "General")),
    "key_percents": formatted(("0.5", "0%"), ("0.5", "0.0%")),
    "key_shows_true": formatted(("1", "\\T\\R\\U\\E"), ('"TRUE"', "General")),
    "key_number_shows_text": formatted(("1", '"x"'), ('"x"', "General")),
    "key_general_shows_same": formatted(("1", "General"), ("1", "0;-0;0")),
    "key_bool_lower": column("A", "True", '"true"') + remove("A1:A2", "Columns:=1, Header:=xlNo"),
    "key_error_text": column("A", "CVErr(xlErrNA)", '"#N/A"') + remove("A1:A2", "Columns:=1, Header:=xlNo"),
    "key_number_bool": column("A", "1", "True") + remove("A1:A2", "Columns:=1, Header:=xlNo"),
    "key_bool_number_text": column("A", "False", '"0"', '"\'0"') + remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "formula_refs_moved": column("A", "1", "1", "2", "3") + 'ws.Range("B3").Formula = "=A$1+A3"\n' +
                          'ws.Range("B4").Formula = "=A1"\n' + remove("A1:B4", "Columns:=1, Header:=xlNo"),
    "formula_refs_below": column("A", "1", "1", "2") + column("D", None, None, None, None, "5") +
                          'ws.Range("B3").Formula = "=D5+A3"\n' + remove("A1:B3", "Columns:=1, Header:=xlNo"),
    "formula_to_removed": column("A", "1", "1", "2") + 'ws.Range("B1").Formula = "=A2&A3"\n' +
                          remove("A1:B3", "Columns:=1, Header:=xlNo"),
    "formula_first_row": column("A", "1", "2", "2") + 'ws.Range("B3").Formula = "=A1"\n' +
                         remove("A1:B3", "Columns:=1, Header:=xlNo"),
    "columns_zero_array": column("A", "1", "1") + remove("A1:A2", "Columns:=Array(0), Header:=xlNo"),
    "columns_negative": column("A", "1", "1") + remove("A1:A2", "Columns:=-1, Header:=xlNo"),
    "columns_partly_outside": PAIRS + remove("A1:B5", "Columns:=Array(1, 3), Header:=xlNo"),
    "columns_fraction": PAIRS + remove("A1:C5", "Columns:=1.7, Header:=xlNo"),
    "columns_numeric_text": PAIRS + remove("A1:C5", 'Columns:="2", Header:=xlNo'),
    "columns_empty_array": PAIRS + remove("A1:C5", "Columns:=Array(), Header:=xlNo"),
    "header_guess_bold": column("A", '"x"', '"y"', '"y"') + 'ws.Range("A1").Font.Bold = True\n' +
                         remove("A1:A3", "Columns:=1, Header:=xlGuess"),
    "header_guess_bold_numbers": column("A", "1", "2", "2") + 'ws.Range("A1").Font.Bold = True\n' +
                                 remove("A1:A3", "Columns:=1, Header:=xlGuess"),
    "header_guess_one_column": column("A", '"id"', "1", "1") + column("B", '"x"', '"y"', '"y"') +
                               remove("A1:B3", "Columns:=2, Header:=xlGuess"),
    "header_guess_caps": column("A", '"NAME"', '"y"', '"y"') + remove("A1:A3", "Columns:=1, Header:=xlGuess"),
    "empty_range": remove("A1:A3", "Columns:=1, Header:=xlNo"),
    "blank_bold": column("A", "1", None, None, "2") + 'ws.Range("A3").Font.Bold = True\n' +
                  remove("A1:A4", "Columns:=1, Header:=xlNo"),
    "single_cell_region": column("A", "1", "1", "2") + column("B", '"a"', '"a"', '"b"') +
                          remove("B2", "Columns:=1, Header:=xlNo"),
    "numbers_as_formulas_text": 'ws.Range("A1").Formula = "=1=1"\n' + column("A", None, "True", '"TRUE"') +
                                remove("A1:A3", "Columns:=1, Header:=xlNo"),
})

#: xlGuess where a header and no header leave different things: the first row repeats further down.
GUESS = "Columns:=1, Header:=xlGuess"
LAYOUTS.update({
    "guess_bold_text": column("A", '"x"', '"y"', '"x"') + 'ws.Range("A1").Font.Bold = True\n' + remove("A1:A3", GUESS),
    "guess_plain_texts": column("A", '"x"', '"y"', '"x"') + remove("A1:A3", GUESS),
    "guess_bold_number": column("A", "1", "2", "1") + 'ws.Range("A1").Font.Bold = True\n' + remove("A1:A3", GUESS),
    "guess_caps": column("A", '"NAME"', '"y"', '"NAME"') + remove("A1:A3", GUESS),
    "guess_text_number": column("A", '"id"', "1", '"id"') + remove("A1:A3", GUESS),
    "guess_format": column("A", "1", "2", "1") + 'ws.Range("A1").NumberFormat = "0.00"\n' +
                    'ws.Range("A3").NumberFormat = "0.00"\n' + remove("A1:A3", GUESS),
    "guess_second_column": column("A", '"x"', '"y"', '"x"') + column("B", '"id"', "1", "1") + remove("A1:B3", GUESS),
    "guess_blank_second": column("A", '"x"', None, '"x"') + remove("A1:A3", GUESS),
    "guess_numbers": column("A", "1", "2", "1") + remove("A1:A3", GUESS),
    "guess_one_row": column("A", "1") + column("B", "1") + remove("A1:B1", GUESS),
    "guess_fill": column("A", '"x"', '"y"', '"x"') + 'ws.Range("A1").Interior.Color = 255\n' + remove("A1:A3", GUESS),
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
        parts = answer.split(";")
        head = ";".join(parts[:2]) if answer.startswith("S") else parts[0]
        cells = parts[-1].split("<>")
        shown = [f"{'ABCDE'[i % 5]}{i // 5 + 1}={cell}" for i, cell in enumerate(cells[:40])
                 if not cell.startswith("~Empty~General~False")]
        print(f"{name}: {head} | {' | '.join(shown)} | used {cells[-1]}")


if __name__ == "__main__":
    main()
