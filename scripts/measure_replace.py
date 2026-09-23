"""Range.Replace, as Excel's object model answers it.

Every layout sits on a sheet of its own in one new workbook: a setup that
fills a few cells and replaces something, then a dump of A1:D4 -- each
cell's formula, value type, number format and prefix character -- the
used range and what Replace returned. The layouts ask what Replace
matches (part or whole, case, wildcards and the ~ escape, formulas as
text), what it leaves behind (a replaced number, date, Boolean or error
typed again, an emptied cell, a prefixed or Text cell), where it looks
(a single cell, several areas), in what order it works, which a formula
it cannot enter shows by stopping there, and what it leaves for the next
Find. A second procedure puts 20 numbers into 26 formats and reads back
the text Replace edits in each -- what the formula bar shows -- by
replacing each digit in turn with q, which leaves text behind.

    python scripts/measure_replace.py

writes tests/fixtures/replace.json, which tests/test_excel_replace.py replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "replace.json"

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
    For Each c In ws.Range("A1:D4").Cells
        out = out & c.Formula & "~" & TypeName(c.Value) & "~" & c.NumberFormat & "~" & c.PrefixCharacter & "<>"
    Next
    Dump = out & ws.UsedRange.Address
End Function

Private Function Where(r As Object, what As String, inside As Long) As String
    Dim hit As Object
    Set hit = r.Find(what, LookIn:=inside, LookAt:=xlPart)
    If hit Is Nothing Then Where = "Nothing" Else Where = hit.Address(False, False)
End Function
'''

A_TO_C = 'ws.Range("A1").Value = "apple"\nws.Range("A2").Value = "pineapple"\nws.Range("A3").Value = "grape"\n'

FORMATTED = ('ws.Range("A1").Value = 1234\nws.Range("A1").NumberFormat = "#,##0"\nws.Range("A2").Value = #1/2/2020#\n'
             'ws.Range("A3").Value = 0.5\nws.Range("A3").NumberFormat = "0%"\nws.Range("A4").Value = 0.75\n'
             'ws.Range("A4").NumberFormat = "h:mm AM/PM"\n')

#: What Find matches in a formatted constant, looking in formulas and then in values.
FIND_TEXTS = ["1,234", "1234", "2020", "43832", "50%", "0.5", "PM", "0.75"]

#: name -> setup; each works on ws, the layout's own sheet, and keeps what Replace returns in v.
LAYOUTS: dict[str, str] = {
    "part": A_TO_C + 'v = ws.Range("A1:A3").Replace("apple", "pear")',
    "whole": A_TO_C + 'v = ws.Range("A1:A3").Replace("apple", "pear", LookAt:=xlWhole)',
    "case_kept": 'ws.Range("A1").Value = "Apple"\nws.Range("A2").Value = "apple"\n'
                 'v = ws.Range("A1:A2").Replace("apple", "x", MatchCase:=True)',
    "case_ignored": 'ws.Range("A1").Value = "Apple"\nws.Range("A2").Value = "APPLE pie"\n'
                    'v = ws.Range("A1:A2").Replace("apple", "x", MatchCase:=False)',
    "every_occurrence": 'ws.Range("A1").Value = "aaa"\nws.Range("A2").Value = "aaaa"\n'
                        'v = ws.Range("A1:A2").Replace("aa", "b")',
    "wildcards": 'ws.Range("A1").Value = "abc"\nws.Range("A2").Value = "axc"\nws.Range("A3").Value = "xxabcxxabc"\n'
                 'ws.Range("A4").Value = "a*c"\nv = ws.Range("A1:A4").Replace("a?c", "Z")',
    "star": 'ws.Range("A1").Value = "xxabcxxabc"\nws.Range("A2").Value = "abc"\nv = ws.Range("A1:A2").Replace("a*c", "-")',
    "star_only": 'ws.Range("A1").Value = "abc"\nws.Range("A2").Value = "abc"\nv = ws.Range("A1").Replace("*", "Z")\n'
                 'v = ws.Range("A2").Replace("?", "Z")',
    "escape": 'ws.Range("A1").Value = "a*c"\nws.Range("A2").Value = "abc"\nv = ws.Range("A1:A2").Replace("a~*c", "Y")',
    "replacement_literal": 'ws.Range("A1").Value = "abc"\nv = ws.Range("A1").Replace("b", "~*?")',
    "numbers_typed_again": 'ws.Range("A1").Value = 123\nws.Range("A2").Value = "1a"\nws.Range("A3").Value = "abc"\n'
                           'ws.Range("A4").Value = 1.5\nv = ws.Range("A1").Replace("2", "5")\n'
                           'v = ws.Range("A2").Replace("a", "")\nv = ws.Range("A3").Replace("abc", "5%")\n'
                           'v = ws.Range("A4").Replace(".", ",")',
    "numeric_arguments": 'ws.Range("A1").Value = 123\nv = ws.Range("A1").Replace(2, 5)',
    "formula": 'ws.Range("B1").Value = 2\nws.Range("C1").Value = 3\nws.Range("A1").Formula = "=B1+1"\n'
               'v = ws.Range("A1").Replace("B1", "C1")',
    "formula_whole": 'ws.Range("B1").Value = 2\nws.Range("A1").Formula = "=B1+1"\n'
                     'v = ws.Range("A1").Replace("=B1+1", "5", LookAt:=xlWhole)',
    "formula_result_not_searched": 'ws.Range("B1").Value = 2\nws.Range("A1").Formula = "=B1+1"\n'
                                   'v = ws.Range("A1:A2").Replace("3", "9")',
    "invalid_formula": 'ws.Range("B1").Value = 2\nws.Range("A1").Formula = "=B1+1"\nws.Range("A2").Value = "a+b"\n'
                       'v = ws.Range("A1:A2").Replace("+", "+)")',
    "text_to_formula": 'ws.Range("A1").Value = "x1+1"\nws.Range("A2").NumberFormat = "@"\nws.Range("A2").Value = "x1+1"\n'
                       'v = ws.Range("A1:A2").Replace("x", "=")',
    "emptied": 'ws.Range("A1").Value = "abc"\nws.Range("A2").Value = "abcd"\nv = ws.Range("A1:A2").Replace("abc", "")',
    "date": 'ws.Range("A1").Value = #1/2/2020#\nws.Range("A2").Value = #1/2/2020#\n'
            'v = ws.Range("A1").Replace("2020", "2021")\nv = ws.Range("A2").Replace("4", "5")',
    "percent": 'ws.Range("A1").Value = 0.5\nws.Range("A1").NumberFormat = "0%"\nv = ws.Range("A1").Replace("0.", "0.0")',
    "time": 'ws.Range("A1").Value = 0.75\nws.Range("A1").NumberFormat = "h:mm AM/PM"\n'
            'v = ws.Range("A1").Replace("PM", "AM")',
    "boolean": 'ws.Range("A1").Value = True\nws.Range("A2").Value = True\nv = ws.Range("A1").Replace("TRUE", "FALSE")\n'
               'v = ws.Range("A2").Replace("T", "X")',
    "error": 'ws.Range("A1").Value = "#N/A"\nv = ws.Range("A1").Replace("N/A", "DIV/0!")',
    "prefixed": 'ws.Range("A1").Value = "\'123"\nv = ws.Range("A1").Replace("1", "9")',
    "text_cell": 'ws.Range("A1").NumberFormat = "@"\nws.Range("A1").Value = "123"\nv = ws.Range("A1").Replace("1", "9")',
    "formatted_number": 'ws.Range("A1").Value = 1234\nws.Range("A1").NumberFormat = "#,##0"\n'
                        'v = ws.Range("A1").Replace("2", "7")',
    "single_cell_scope": 'ws.Range("A1").Value = "cat"\nws.Range("B2").Value = "cat"\nv = ws.Range("A1").Replace("cat", "dog")',
    "single_cell_miss": 'ws.Range("A1").Value = "x"\nws.Range("B2").Value = "cat"\nv = ws.Range("A1").Replace("cat", "dog")',
    "several_areas": 'ws.Range("A1").Value = "cat"\nws.Range("B2").Value = "cat"\nws.Range("C3").Value = "cat"\n'
                     'v = ws.Range("A1,C3").Replace("cat", "dog")',
    "whole_sheet": 'ws.Range("A1").Value = "cat"\nws.Range("D4").Value = "cat"\nv = ws.Cells.Replace("cat", "dog")',
    "nothing_found": 'ws.Range("A1").Value = "x"\nv = ws.Range("A1:A2").Replace("zzz", "y")',
    "empty_what": 'ws.Range("A1").Value = "x"\nv = ws.Range("A1:B2").Replace("", "y")',
    "number_whole": 'ws.Range("A1").Value = 12\nws.Range("A2").Value = 123\n'
                    'v = ws.Range("A1:A2").Replace("12", "7", LookAt:=xlWhole)',
    # What Replace leaves behind for Find, and takes from it: Find starts after the first cell.
    "leaves_look_at": 'ws.Range("A1").Value = "q"\nws.Range("A2").Value = "pq"\n'
                      'v = ws.Range("A1:A2").Replace("zz", "y", LookAt:=xlWhole)\n'
                      'v = ws.Range("A1:A2").Find("q").Address(False, False)',
    "leaves_order": 'ws.Range("A1").Value = "x"\nws.Range("B1").Value = "q"\nws.Range("A2").Value = "q"\n'
                    'v = ws.Range("A1:B2").Replace("zz", "y", SearchOrder:=xlByColumns)\n'
                    'v = ws.Range("A1:B2").Find("q").Address(False, False)',
    "leaves_case": 'ws.Range("A1").Value = "x"\nws.Range("A2").Value = "Q"\nws.Range("A3").Value = "q"\n'
                   'v = ws.Range("A1:A3").Replace("zz", "y", MatchCase:=True)\n'
                   'v = ws.Range("A1:A3").Find("q").Address(False, False)',
    "leaves_what": 'ws.Range("A1").Value = "x"\nws.Range("A2").Value = "a"\nws.Range("A3").Value = "b"\n'
                   'Set f = ws.Range("A1:A3").Find("a")\nv = ws.Range("A1:A3").Replace("b", "bb")\n'
                   'v = ws.Range("A1:A3").FindNext(ws.Range("A2")).Address(False, False)',
    "takes_look_at": 'ws.Range("A1").Value = "pq"\nws.Range("A2").Value = "q"\n'
                     'Set f = ws.Range("A1:A2").Find("zz", LookAt:=xlWhole)\n'
                     'v = ws.Range("A1:A2").Replace("q", "r")',
    "takes_case": 'ws.Range("A1").Value = "Q"\nws.Range("A2").Value = "q"\n'
                  'Set f = ws.Range("A1:A2").Find("zz", MatchCase:=True)\n'
                  'v = ws.Range("A1:A2").Replace("q", "r")',
    "find_formatted": FORMATTED + "v = " + ' & "," & '.join(
        f'Where(ws.Range("A1:A4"), "{text}", {inside})' for inside in ("xlFormulas", "xlValues") for text in FIND_TEXTS),
    "replace_formatted": FORMATTED + 'v = ws.Range("A1:A4").Replace("1,234", "x")',
    "values_then_replace": 'ws.Range("A1").Value = 1234\nws.Range("A1").NumberFormat = "#,##0"\n'
                           'Set f = ws.Range("A1:A2").Find("zz", LookIn:=xlValues)\n'
                           'v = ws.Range("A1:A2").Replace("1,234", "x")',
    "leaves_look_in": 'ws.Range("A1").Value = 1234\nws.Range("A1").NumberFormat = "#,##0"\n'
                      'Set f = ws.Range("A1:A2").Find("zz", LookIn:=xlValues)\n'
                      'v = ws.Range("A1:A2").Replace("zz", "y")\nSet f = ws.Range("A1:A2").Find("1,234")\n'
                      'If f Is Nothing Then v = "Nothing" Else v = f.Address(False, False)',
    # Each cell is its own question about * and ~: the lazy middle star, the greedy trailing one.
    "stars": "\n".join(f'ws.Range("{cell}").Value = "{text}"\nv = ws.Range("{cell}").Replace("{what}", "Z"{extra})'
                       for cell, text, what, extra in (
                           ("A1", "abcabc", "*c", ""), ("B1", "abab", "a*", ""), ("C1", "abcabc", "*b*", ""),
                           ("D1", "aXbYcaXbYc", "a*b*c", ""), ("A2", "abc", "**", ""), ("B2", "abc", "*?", ""),
                           ("C2", "abc", "?*", ""), ("D2", "abcb", "b*", ""), ("A3", "abcbc", "a*c", ""),
                           ("B3", "xa", "a*", ""), ("C3", "bab", "*a", ""), ("D3", "abc", "c*", ""),
                           ("A4", "aaa", "a*a", ""), ("B4", "abcabc", "*", ", LookAt:=xlWhole"),
                           ("C4", "ab", "a~", ""), ("D4", "a~b", "~", ""))),
    "tildes": "\n".join(f'ws.Range("{cell}").Value = "{text}"\nv = ws.Range("{cell}").Replace("{what}", "Z")'
                        for cell, text, what in (
                            ("A1", "a~b", "a~"), ("B1", "ab", "a~"), ("C1", "a~~b", "~~"), ("D1", "a?b", "~?"),
                            ("A2", "abc", "~a"), ("B2", "a~b", "~b"), ("C2", "ab", "b~"), ("D2", "ab~", "b~"),
                            ("A3", "a~b", "~"), ("B3", "abc", "*~"))),
    # A replacement that leaves a formula Excel cannot read stops the Replace there: which cells
    # went before it shows the order Replace works in.
    "invalid_rows": 'ws.Range("A1").Value = "c+d"\nws.Range("B1").Value = "a+b"\nws.Range("A2").Formula = "=1+1"\n'
                    'ws.Range("B2").Value = "e+f"\nv = ws.Range("A1:B2").Replace("+", "+)")',
    "invalid_columns": 'ws.Range("A1").Value = "c+d"\nws.Range("B1").Value = "a+b"\nws.Range("A2").Formula = "=1+1"\n'
                       'ws.Range("B2").Value = "e+f"\nv = ws.Range("A1:B2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "invalid_mid_row": 'ws.Range("A1").Value = "a+b"\nws.Range("B1").Formula = "=1+1"\nws.Range("C1").Value = "c+d"\n'
                       'ws.Range("A2").Value = "e+f"\nv = ws.Range("A1:C2").Replace("+", "+)")',
    "invalid_mid_column": 'ws.Range("A1").Value = "a+b"\nws.Range("A2").Formula = "=1+1"\nws.Range("A3").Value = "c+d"\n'
                          'ws.Range("B1").Value = "e+f"\n'
                          'v = ws.Range("A1:B3").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "invalid_first_row": 'ws.Range("A1").Formula = "=1+1"\nws.Range("B1").Value = "a+b"\nws.Range("A2").Value = "c+d"\n'
                         'v = ws.Range("A1:B2").Replace("+", "+)")',
    "invalid_first_column": 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = "a+b"\nws.Range("B1").Value = "c+d"\n'
                            'v = ws.Range("A1:B2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "invalid_diagonal": 'ws.Range("A1").Value = "a+b"\nws.Range("B2").Formula = "=1+1"\nws.Range("C3").Value = "c+d"\n'
                        'ws.Range("A3").Value = "e+f"\nws.Range("C1").Value = "g+h"\nv = ws.Range("A1:C3").Replace("+", "+)")',
    "invalid_diagonal_columns": 'ws.Range("A1").Value = "a+b"\nws.Range("B2").Formula = "=1+1"\n'
                                'ws.Range("C3").Value = "c+d"\nws.Range("A3").Value = "e+f"\nws.Range("C1").Value = "g+h"\n'
                                'v = ws.Range("A1:C3").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "invalid_long_row": 'ws.Range("A1").Value = "a+b"\nws.Range("B1").Value = "c+d"\nws.Range("C1").Formula = "=1+1"\n'
                        'ws.Range("D1").Value = "e+f"\nws.Range("A2").Value = "g+h"\nv = ws.Range("A1:D2").Replace("+", "+)")',
    "invalid_long_column": 'ws.Range("A1").Value = "a+b"\nws.Range("A2").Value = "c+d"\nws.Range("A3").Formula = "=1+1"\n'
                           'ws.Range("A4").Value = "e+f"\nws.Range("B1").Value = "g+h"\n'
                           'v = ws.Range("A1:B4").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "invalid_areas": 'ws.Range("A1").Value = "a+b"\nws.Range("B2").Formula = "=1+1"\nws.Range("C3").Value = "c+d"\n'
                     'v = ws.Range("C3,B2,A1").Replace("+", "+)")',
    "invalid_areas_columns": 'ws.Range("A1").Value = "a+b"\nws.Range("B2").Formula = "=1+1"\nws.Range("C3").Value = "c+d"\n'
                             'v = ws.Range("C3,B2,A1").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "columns_all_valid": 'ws.Range("A1").Value = "a+b"\nws.Range("A2").Value = "c+d"\nws.Range("B1").Value = "e+f"\n'
                         'ws.Range("B2").Value = "g+h"\nv = ws.Range("A1:B2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "columns_block_areas": 'ws.Range("A1").Value = "a+b"\nws.Range("A2").Value = "c+d"\nws.Range("C1").Formula = "=1+1"\n'
                           'ws.Range("C2").Value = "e+f"\n'
                           'v = ws.Range("A1:A2,C1:C2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "columns_block_areas_first": 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = "c+d"\n'
                                 'ws.Range("C1").Value = "e+f"\nws.Range("C2").Value = "g+h"\n'
                                 'v = ws.Range("A1:A2,C1:C2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "rows_block_areas": 'ws.Range("A1").Value = "a+b"\nws.Range("B1").Formula = "=1+1"\nws.Range("A3").Value = "c+d"\n'
                        'ws.Range("B3").Value = "e+f"\nv = ws.Range("A1:B1,A3:B3").Replace("+", "+)")',
    "rows_block_areas_second": 'ws.Range("A1").Value = "a+b"\nws.Range("B1").Value = "c+d"\n'
                               'ws.Range("A3").Formula = "=1+1"\nws.Range("B3").Value = "e+f"\n'
                               'v = ws.Range("A3:B3,A1:B1").Replace("+", "+)")',
    "columns_one_row": 'ws.Range("A1").Formula = "=1+1"\nws.Range("B1").Value = "a+b"\nws.Range("C1").Value = "c+d"\n'
                       'v = ws.Range("A1:C1").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "rows_one_column": 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = "a+b"\nws.Range("A3").Value = "c+d"\n'
                       'v = ws.Range("A1:A3").Replace("+", "+)")',
    "columns_offset": 'ws.Range("B2").Formula = "=1+1"\nws.Range("B3").Value = "a+b"\nws.Range("C2").Value = "c+d"\n'
                      'v = ws.Range("B2:C3").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "rows_offset": 'ws.Range("B2").Value = "a+b"\nws.Range("C2").Formula = "=1+1"\nws.Range("B3").Value = "c+d"\n'
                   'v = ws.Range("B2:C3").Replace("+", "+)")',
    "rows_active_cell": 'ws.Range("A1").Value = "c+d"\nws.Range("B1").Value = "a+b"\nws.Range("A2").Formula = "=1+1"\n'
                        'ws.Range("B2").Value = "e+f"\nws.Activate\nws.Range("B2").Select\n'
                        'v = ws.Range("A1:B2").Replace("+", "+)")',
    "columns_active_cell": 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = "a+b"\nws.Range("B1").Value = "c+d"\n'
                           'ws.Range("B2").Value = "e+f"\nws.Activate\nws.Range("B2").Select\n'
                           'v = ws.Range("A1:B2").Replace("+", "+)", SearchOrder:=xlByColumns)',
    "text_cell_formats": "\n".join(f'ws.Range("A{row}").NumberFormat = "@"\nws.Range("A{row}").Value = "x"\n'
                                   f'v = ws.Range("A{row}").Replace("x", "{replacement}")'
                                   for row, replacement in ((1, "1,000"), (2, "1:60"), (3, "1 1/2"), (4, "$5"))),
    "same_text": 'ws.Range("A1").NumberFormat = "@"\nws.Range("A1").Value = "123"\nv = ws.Range("A1").Replace("1", "1")\n'
                 'ws.Range("A2").NumberFormat = "@"\nws.Range("A2").Value = "123"\nws.Range("A2").NumberFormat = "General"\n'
                 'v = ws.Range("A2").Replace("2", "2")\nws.Range("A3").Value = "\'abc"\nv = ws.Range("A3").Replace("abc", "")',
    "overlapping_areas": 'ws.Range("B2").Value = "a"\nv = ws.Range("A1:B2,B2:C3").Replace("a", "ab")',
    "merged_and_hidden": 'ws.Range("A1:B2").Merge\nws.Range("A1").Value = "cat"\nws.Range("A3").Value = "cat"\n'
                         'ws.Rows(3).Hidden = True\nv = ws.Range("A1:C4").Replace("cat", "dog")',
    "date_and_time": 'ws.Range("A1").Value = 43832.5\nws.Range("A1").NumberFormat = "m/d/yyyy h:mm"\n'
                     'v = ws.Range("A1").Replace("PM", "AM")\nws.Range("A2").Value = "1/2/2020  12:00"\n'
                     'ws.Range("A3").Value = "1/2/2020   1:00 PM"\nws.Range("A4").Value = 0.5\n'
                     'ws.Range("A4").NumberFormat = "m/d/yyyy"\nv = ws.Range("A4").Replace("PM", "AM")',
    "bad_arguments": 'ws.Range("A1").Value = "a"\nv = ws.Range("A1").Replace("a", "b", LookAt:=5)\n'
                     'v = v & "," & Err.Number\nErr.Clear\nws.Range("A2").Value = "a"\n'
                     'v = ws.Range("A2").Replace("a", "b", SearchOrder:=7)\nv = v & "," & Err.Number',
    "prefix_formula_number":'ws.Range("A1").Value = "\'x1+1"\nv = ws.Range("A1").Replace("x", "=")\n'
                             'ws.Range("A2").Value = "\'123"\nws.Range("A2").Value = 5\nv = ws.Range("A2").Replace("5", "6")',
    "formula_case": 'ws.Range("A1").Formula = "=B1+1"\nv = ws.Range("A1").Replace("b1", "c1")',
    "formula_to_text": 'ws.Range("A1").Formula = "=B1+1"\nv = ws.Range("A1").Replace("=", "")',
    "text_cell_brings": "\n".join(f'ws.Range("A{row}").NumberFormat = "@"\nws.Range("A{row}").Value = "{text}"\n'
                                  f'v = ws.Range("A{row}").Replace("{text}", "{replacement}")'
                                  for row, text, replacement in ((1, "abc", "5%"), (2, "x", "1/2/2020"),
                                                                 (3, "x", "'5"), (4, "x", "TRUE"))),
    "fraction_cell": 'ws.Range("A1").NumberFormat = "0.00"\nws.Range("A1").Value = "x"\n'
                     'v = ws.Range("A1").Replace("x", "1/2")\nws.Range("A2").Value = "x"\n'
                     'v = ws.Range("A2").Replace("x", "1 1/2")',
    "prefix_removed": 'ws.Range("A1").Value = "\'123"\nws.Range("A2").Value = "\'456"\n'
                      'v = Where(ws.Range("A1:A2"), "\'4", xlFormulas) & "," & Where(ws.Range("A1:A2"), "\'4", xlValues)\n'
                      'ws.Range("A1").Replace "\'", ""',
}

EDIT_TEXTS = '''Private Function EditText(c As Object) As String
    Dim d As Long
    For d = 0 To 9
        c.Replace CStr(d), "q", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
        If TypeName(c.Value2) = "String" Then
            EditText = Replace(c.Value2, "q", CStr(d))
            Exit Function
        End If
    Next
    EditText = "?" & c.Formula
End Function
'''

#: The values and formats whose text Replace edits, read by replacing each digit in turn with q.
TEXT_VALUES = ["0", "0.5", "0.75", "1", "1.25", "43832", "43832.5", "-1", "-0.25", "1234.5678", "0.055", "1E+20",
               "1E-20", "60", "2958465.5", "2958466", "0.500005787037037", "0.999999999", "12345678901234567", "0.1"]
TEXT_FORMATS = ["General", "0.00", "#,##0", "$#,##0.00_);($#,##0.00)", "0%", "0.00%", "0.00E+00", "# ?/?",
                "m/d/yyyy", "d-mmm-yy", "yyyy-mm-dd", "mmm", "dddd", "h:mm", "h:mm AM/PM", "h:mm:ss", "[h]:mm:ss",
                "[mm]:ss", "mm:ss.0", "m/d/yyyy h:mm", "@", "0\\%", "0.0%;[Red]-0.0%", "[$-409]mmmm d, yyyy;@", "ss",
                "h"]


def texts_code() -> str:
    """A procedure that puts every value in every format and reads back the text Replace edits."""
    values = ", ".join(TEXT_VALUES)
    formats = ", ".join(f'"{code}"' for code in TEXT_FORMATS)
    return EDIT_TEXTS + f'''Public Function Texts() As String
    Dim wb As Object, ws As Object, values As Variant, formats As Variant, r As Long, k As Long, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    values = Array({values})
    formats = Array({formats})
    For r = 0 To UBound(values)
        For k = 0 To UBound(formats)
            ws.Cells(r + 1, k + 1).Value = values(r)
            ws.Cells(r + 1, k + 1).NumberFormat = formats(k)
            out = out & EditText(ws.Cells(r + 1, k + 1)) & "^"
        Next
        out = out & "|"
    Next
    wb.Close False
    Texts = out
End Function
'''


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, f As Object, out As String", "On Error Resume Next", "Err.Clear",
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
                  # Every layout starts from the search settings a new session has.
                  'Set f = ws.Range("A1").Find("", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, '
                  "MatchCase:=False)",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup))
    build.insert(2, "Dim f As Object")
    build += ["wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        texts = excel.run_vba(texts_code(), "Texts", timeout=600.0)
        assert texts.ok, f"{texts.outcome}: {texts.message} {texts.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    rows = [row.split("^")[: len(TEXT_FORMATS)] for row in str(texts.value).split("|")[: len(TEXT_VALUES)]]
    record = {"helper": HELPER, "year": datetime.date.today().year,
              "layouts": [{"name": name, "setup": setup, "answers": answer}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)],
              "edit_texts": {"code": texts_code(), "values": TEXT_VALUES, "formats": TEXT_FORMATS, "texts": rows}}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        head, _, dump = answer.partition(";")
        cells = [cell for cell in dump.split("<>")[:-1] if not cell.startswith("~Empty~General~")]
        print(f"{name}: {head} | {' | '.join(cells)} | used {dump.split('<>')[-1]}")
    for value, row in zip(TEXT_VALUES, rows, strict=True):
        print(value, "=>", " ^ ".join(f"{code}:{text}" for code, text in zip(TEXT_FORMATS, row, strict=True)))


if __name__ == "__main__":
    main()
