"""Range.Subtotal and Range.RemoveSubtotal in live Excel.

A macro gives each case a sheet of its own in one new workbook, types a
list into it cell by cell, makes any setting the case needs and then
calls Subtotal, RemoveSubtotal or both, an error's number noted where a
call raises one. It reads back every cell the sheet holds with its
formula, value and type, bold, italic, number format and fill, each row's
outline level and whether it is hidden, the outline's summary row and
anything else the case asks for. A sheet of its own, scratch, takes the
reads and settings that leave Excel a number format for the first label
of a Subtotal (see _subtotals), which each case starts from General. The
workbook is saved, each sheet's sheetPr, sheetFormatPr, rows and
rowBreaks kept, opened again and every sheet read once more. Last, an
application's handlers log the Change and Calculate events a Subtotal and
a RemoveSubtotal raise, with a formula reading the list and without.

    python scripts/measure_subtotals.py

writes tests/fixtures/subtotals/: the workbook as Excel saved it and
subtotals.json, which tests/test_excel_subtotals.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "subtotals"

#: Lists are typed a cell at a time, rows split by "!" and cells by ",", an empty cell left as it is.
BASE = "Region,Product,Sales!East,A,10!East,B,20!West,A,30!West,C,40!North,A,5"
NEST = "Region,Product,Sales!East,A,10!East,A,20!East,B,5!West,A,30!West,C,40"
THREE = "A,B,C,V!x,p,1,1!x,p,1,2!x,p,2,3!x,q,1,4!y,p,1,5!y,p,2,6"
HEADED = 'ws.Range("A1:C1").Font.Bold = True: ws.Range("A1:C1").Interior.Color = RGB(200, 200, 255)'


def large() -> str:
    rows = ["Region,Product,Units,Sales"]
    for index in range(40):
        region = ("East", "North", "South", "West")[index // 10]
        product = ("Bolt", "Nut", "Screw")[index % 10 // 4]
        rows.append(f"{region},{product},{index % 7 + 1},{(index * 37) % 101}")
    return "!".join(rows)


#: name -> (the list typed, statements made before, the calls measured, statements made once the sheet is read).
CASES: dict[str, tuple[str, list[str], list[str], list[str]]] = {}

for function in ("xlSum", "xlCount", "xlAverage", "xlMax", "xlMin", "xlProduct", "xlCountNums", "xlStDev",
                 "xlStDevP", "xlVar", "xlVarP"):
    CASES["fn_" + function[2:].lower()] = (BASE, [], [f'ws.Range("A1").Subtotal 1, {function}, Array(3)'], [])

CASES.update({
    # Which cells make the list.
    "above": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove'], []),
    "below_named": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryBelow'], []),
    "cell_inside": (BASE, [], ['ws.Range("B3").Subtotal 1, xlSum, Array(3)'], []),
    "range_rows": (BASE, [], ['ws.Range("A1:C4").Subtotal 1, xlSum, Array(3)'], []),
    "range_columns": (BASE, [], ['ws.Range("B1:C6").Subtotal 1, xlSum, Array(2)'], []),
    "range_wider": (BASE, [], ['ws.Range("A1:E8").Subtotal 1, xlSum, Array(3)'], []),
    "whole_columns": (BASE, [], ['ws.Range("A:C").Subtotal 1, xlSum, Array(3)'], []),
    "whole_rows": (BASE, [], ['ws.Range("1:6").Subtotal 1, xlSum, Array(3)'], []),
    "header_only": ("Region,Product,Sales", [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "one_row": ("Region,Sales!East,10", [], ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "no_header": ("1,10!1,20!2,30", [], ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "empty_cell": ("", [], ['ws.Range("H10").Subtotal 1, xlSum, Array(1)'], []),
    "offset_list": ("!!!!,,Region,Sales!,,East,10!,,East,20!,,West,30", [],
                    ['ws.Range("C5").Subtotal 1, xlSum, Array(2)'], []),
    "outside_cells": (BASE, ['ws.Range("F2").Value = "side"', 'ws.Range("F9").Value = "low side"',
                             'ws.Range("A10").Value = "after"'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "outside_formulas": (BASE, ['ws.Range("F1").Formula = "=SUM(C2:C6)"', 'ws.Range("F2").Formula = "=C6"',
                                'ws.Names.Add "Sold", "=\'" & ws.Name & "\'!$C$2:$C$6"'],
                         ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                          'notes = notes & "name=" & ws.Names("Sold").RefersTo & " "'], []),
    "list_formulas": ("Region,Sales,Double!East,10,=B2*2!East,20,=B3*2!West,30,=B4*2", [],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(2, 3)'], []),
    # What a group is and what its label says.
    "values_kinds": ("Key,Sales!5,1!5,2!'5,3!=5,4!TRUE,5!1,6!0,7!,8!#N/A,9!#N/A,10!=1/0,11!=0/0,12!"
                     "1/2/2024,13!45293,14!east,15!East,16!EAST,17!East ,18!East,19!=\"\",20! ,21!x,22", [],
                     ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_numbers": ("Key,Sales!1234567890123,1!0.123456789012,2!1E+20,3!3.5,4!$7.00,5!50%,6!-2,7", [
                       'ws.Range("A5").NumberFormat = "0.00"'], ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_narrow": ("Key,Sales!1/2/2024,1!1/2/2024,2!123456,3!abcdefghijk,4", ['ws.Columns("A").ColumnWidth = 2'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_blanks": ("Key,Sales!East,1!East,2!,3!East,4!,5!,6!West,7", [],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_leading_blank": ("Key,Sales!,1!,2!East,3", [], ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_all_blank": ("Key,Sales!,1!,2", [], ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_formula_text": ("Key,Sales!=\"A\"&\"B\",1!AB,2!=UPPER(\"ab\"),3!ab,4", [],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_dates": ("Key,Sales!1/2/2024,1!1/2/2024 12:00,2!1/3/2024,3", ['ws.Range("A4").NumberFormat = "mmm d"'],
                     ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    # The columns totalled.
    "total_scalar": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, 3'], []),
    "total_text_column": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(2, 3)'], []),
    "total_groupby": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(1, 3)'], []),
    "total_groupby_only": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(1)'], []),
    "total_duplicate": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3, 3)'], []),
    "total_order": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3, 2)'], []),
    "total_beyond": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(4)'], []),
    "total_zero": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(0)'], []),
    "total_empty": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array()'], []),
    "total_string": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array("3")'], []),
    "total_count_text": (BASE, [], ['ws.Range("A1").Subtotal 1, xlCount, Array(2)'], []),
    "groupby_zero": (BASE, [], ['ws.Range("A1").Subtotal 0, xlSum, Array(3)'], []),
    "groupby_beyond": (BASE, [], ['ws.Range("A1").Subtotal 4, xlSum, Array(3)'], []),
    "groupby_string": (BASE, [], ['ws.Range("A1").Subtotal "1", xlSum, Array(3)'], []),
    "groupby_last": (BASE, [], ['ws.Range("A1").Subtotal 3, xlCount, Array(1)'], []),
    "function_bad": (BASE, [], ['ws.Range("A1").Subtotal 1, 1, Array(3)'], []),
    "function_zero": (BASE, [], ['ws.Range("A1").Subtotal 1, 0, Array(3)'], []),
    "summary_bad": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=5'], []),
    # Subtotals over subtotals, and replacing them.
    "nested_sum": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                              'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    "nested_count": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                'ws.Range("A1").Subtotal 2, xlCount, Array(3), Replace:=False'], []),
    "nested_same_column": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                      'ws.Range("A1").Subtotal 1, xlAverage, Array(3), Replace:=False'], []),
    "nested_above": (NEST, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove',
        'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, SummaryBelowData:=xlSummaryAbove'], []),
    "nested_mixed": (NEST, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3)',
        'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, SummaryBelowData:=xlSummaryAbove'], []),
    "nested_three": (THREE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(4)',
                                 'ws.Range("A1").Subtotal 2, xlSum, Array(4), Replace:=False',
                                 'ws.Range("A1").Subtotal 3, xlSum, Array(4), Replace:=False'], []),
    "nested_inner_cell": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                     'ws.Range("C3").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    "replace_true": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                'ws.Range("A1").Subtotal 2, xlCount, Array(3)'], []),
    "replace_above_below": (NEST, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove',
        'ws.Range("A1").Subtotal 2, xlSum, Array(3)'], []),
    "replace_false_fresh": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), Replace:=False'], []),
    "twice_same": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                              'ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "large": (large(), [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3, 4)',
                            'ws.Range("A1").Subtotal 2, xlCount, Array(4), Replace:=False'], []),
    # RemoveSubtotal.
    "remove_basic": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', 'ws.Range("A1").RemoveSubtotal'], []),
    "remove_nested": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                 'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False',
                                 'ws.Range("A1").RemoveSubtotal'], []),
    "remove_above": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove',
                                'ws.Range("A1").RemoveSubtotal'], []),
    "remove_none": (BASE, [], ['ws.Range("A1").RemoveSubtotal'], []),
    "remove_empty_cell": ("", [], ['ws.Range("H10").RemoveSubtotal'], []),
    "remove_manual_formula": (BASE, ['ws.Range("D3").Formula = "=SUBTOTAL(9,C2:C3)"'],
                              ['ws.Range("A1").RemoveSubtotal'], []),
    "remove_manual_in_list": (BASE, ['ws.Range("C3").Formula = "=SUBTOTAL(9,C2:C2)"'],
                              ['ws.Range("A1").RemoveSubtotal'], []),
    "remove_other_function": (BASE, ['ws.Range("D3").Formula = "=SUM(C2:C3)"'], ['ws.Range("A1").RemoveSubtotal'],
                              []),
    "remove_inner_cell": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', 'ws.Range("B5").RemoveSubtotal'],
                          []),
    "remove_range_rows": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                     'ws.Range("A1:C5").RemoveSubtotal'], []),
    "remove_edited_label": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                       'ws.Range("A4").Value = "changed"', 'ws.Range("C7").Value = 70',
                                       'ws.Range("A1").RemoveSubtotal'], []),
    "returns": (BASE, [], ['notes = notes & "sub=" & TypeName(ws.Range("A1").Subtotal(1, xlSum, Array(3))) & " "',
                           'notes = notes & "remove=" & TypeName(ws.Range("A1").RemoveSubtotal()) & " "'], []),
    # The outline it makes, and one there already.
    "outline_in_list": (BASE, ['ws.Rows("2:3").Group'], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "outline_elsewhere": (BASE, ['ws.Range("A12").Value = "x"', 'ws.Range("A13").Value = "y"', 'ws.Rows("12:13").Group'],
                          ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "outline_columns": (BASE, ['ws.Columns("B:C").Group'], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                                            'notes = notes & "col=" & ws.Columns(2).OutlineLevel & " "'],
                        []),
    "remove_outline_in_list": (BASE, ['ws.Rows("2:3").Group'], ['ws.Range("A1").RemoveSubtotal'], []),
    "remove_outline_elsewhere": (BASE, ['ws.Range("A12").Value = "x"', 'ws.Range("A13").Value = "y"',
                                        'ws.Rows("12:13").Group'],
                                 ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', 'ws.Range("A1").RemoveSubtotal'], []),
    "remove_outline_columns": (BASE, ['ws.Columns("B:C").Group'],
                               ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', 'ws.Range("A1").RemoveSubtotal',
                                'notes = notes & "col=" & ws.Columns(2).OutlineLevel & " "'], []),
    "summary_was_above": (BASE, ['ws.Outline.SummaryRow = xlSummaryAbove'],
                          ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "collapsed": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', 'ws.Outline.ShowLevels RowLevels:=2'], []),
    "collapsed_then_replace": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                          'ws.Outline.ShowLevels RowLevels:=2',
                                          'ws.Range("A1").Subtotal 2, xlSum, Array(3)'], []),
    "collapsed_then_remove": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                         'ws.Outline.ShowLevels RowLevels:=2', 'ws.Range("A1").RemoveSubtotal'], []),
    "hidden_row": (BASE, ['ws.Rows(3).Hidden = True'], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    # Formats the rows put in take.
    "formats_currency": ("Key,Sales,Price!a,1.5,$3.00!a,2.25,$4.00!b,3,$5.00", [],
                         ['ws.Range("A1").Subtotal 1, xlSum, Array(2, 3)'], []),
    "formats_rows": (BASE, ['ws.Range("A3:C3").Interior.Color = RGB(255, 255, 0)', 'ws.Range("A3:C3").Font.Italic = True',
                            'ws.Range("C2:C6").NumberFormat = "0.0"'],
                     ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "formats_header_above": (BASE, [HEADED],
                             ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove'], []),
    "formats_header_below": (BASE, [HEADED], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "formats_label_bold": (BASE, ['ws.Range("A2:A6").Font.Bold = True'], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'],
                           []),
    "formats_whole_row": (BASE, ['ws.Rows(3).Interior.Color = RGB(0, 255, 0)'],
                          ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    # Page breaks.
    "breaks_below": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True'], []),
    "breaks_above": (BASE, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True, SummaryBelowData:=xlSummaryAbove'], []),
    "breaks_nested": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True',
                                 'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    "breaks_nested_inner": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                       'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, PageBreaks:=True'],
                            []),
    "breaks_remove": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True',
                                 'ws.Range("A1").RemoveSubtotal'], []),
    "breaks_replace": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True',
                                  'ws.Range("A1").Subtotal 2, xlSum, Array(3)'], []),
    "breaks_manual_kept": (BASE, ['ws.HPageBreaks.Add Before:=ws.Range("A20")'],
                           ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True',
                            'ws.Range("A1").RemoveSubtotal'], []),
    # Where Excel refuses.
    "protected": (BASE, ["ws.Protect"], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], ["ws.Unprotect"]),
    "protected_remove": (BASE, ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', "ws.Protect"],
                         ['ws.Range("A1").RemoveSubtotal'], ["ws.Unprotect"]),
    "table": (BASE, ['ws.ListObjects.Add xlSrcRange, ws.Range("A1:C6"), , xlYes'],
              ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "filtered": (BASE, ['ws.Range("A1:C6").AutoFilter 1, "East"'], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                                                    'notes = notes & "filter=" & ws.FilterMode & " "'],
                 []),
    "manual_calculation": (BASE, ["Application.Calculation = xlCalculationManual"],
                           ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'],
                           ["Application.Calculation = xlCalculationAutomatic"]),
    # More of which cells make the list, and where labels go.
    "range_two_regions": ("Region,Product,Sales!East,A,10!East,B,20!West,A,30!!!North,A,5!North,B,6", [],
                          ['ws.Range("A1:C8").Subtotal 1, xlSum, Array(3)'], []),
    "single_column": ("Key!a!a!b", [], ['ws.Range("A1").Subtotal 1, xlCount, Array(1)'], []),
    "groupby_middle_total": (BASE, [], ['ws.Range("A1").Subtotal 2, xlSum, Array(2, 3)'], []),
    "groupby_total_outside": (",Region,Product,Sales!,East,A,10!,East,B,20!,West,A,30", [
        'ws.Range("B20").Value = "below"', 'ws.Range("G2").Value = "right"', 'ws.Range("A30").Value = "left"'],
        ['ws.Range("B1").Subtotal 1, xlSum, Array(1, 3)'], []),
    "groupby_total_remove": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(1, 3)',
                                        'ws.Range("B1").RemoveSubtotal'], []),
    "groupby_total_twice": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(1, 3)',
                                       'ws.Range("B1").Subtotal 2, xlCount, Array(4)'], []),
    "replace_deletes_manual": (BASE, ['ws.Range("D3").Formula = "=SUBTOTAL(9,C2:C3)"'],
                               ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "values_spaces": ("Key,Sales!East ,1!East,2! East,3!Ea st,4!East,5!  ,6!x,7", [],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(2)'], []),
    "values_all_blank_above": ("Key,Sales!,1!,2", [],
                               ['ws.Range("A1").Subtotal 1, xlSum, Array(2), SummaryBelowData:=xlSummaryAbove'], []),
    # Summary settings over an outline already there.
    "summary_two": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=2'], []),
    "summary_negative": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=-1'], []),
    "summary_text": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:="0"'], []),
    "nested_mixed_reverse": (NEST, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove',
        'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False, SummaryBelowData:=xlSummaryBelow'], []),
    "outline_in_list_above": (BASE, ['ws.Rows("2:3").Group'],
                              ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove'], []),
    "outline_in_list_keep": (BASE, ['ws.Rows("2:3").Group'],
                             ['ws.Range("A1").Subtotal 1, xlSum, Array(3), Replace:=False'], []),
    "outline_columns_keep": (BASE, ['ws.Columns("B:C").Group'],
                             ['ws.Range("A1").Subtotal 1, xlSum, Array(3), Replace:=False',
                              'notes = notes & "col=" & ws.Columns(2).OutlineLevel & " "'], []),
    "remove_hidden_manual": (BASE, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)', "ws.Rows(2).Hidden = True",
                                        'ws.Range("A1").RemoveSubtotal'], []),
    "collapsed_then_nested": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                         'ws.Outline.ShowLevels RowLevels:=2',
                                         'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    "remove_nested_function": (BASE, ['ws.Range("D3").Formula = "=ROUND(SUBTOTAL(9,C2:C3),0)"'],
                               ['ws.Range("A1").RemoveSubtotal'], []),
    "remove_text": (BASE, ['ws.Range("D3").Value = "\'=SUBTOTAL(9,C2:C3)"'], ['ws.Range("A1").RemoveSubtotal'], []),
    # Page breaks a later call takes from the list.
    "breaks_default_after_manual": (BASE, ['ws.HPageBreaks.Add Before:=ws.Range("A4")'],
                                    ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "breaks_inner_after_reset": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3), PageBreaks:=True',
                                            "ws.ResetAllPageBreaks",
                                            'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    # The number format the first label takes when the row it copies holds General.
    "register_set": (BASE, ['scratch.Range("B1").NumberFormat = "0.000"'],
                     ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_read": (BASE, ['scratch.Range("B2").NumberFormat = "0.0000"', 'scratch.Range("B2").Value = 7',
                             'scratch.Range("A1").NumberFormat = "General"', 'v = scratch.Range("B2").Value'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_read_address": (BASE, ['scratch.Range("B2").NumberFormat = "0.0000"', 'scratch.Range("B2").Value = 7',
                                     'scratch.Range("A1").NumberFormat = "General"',
                                     'v = scratch.Range("B2").Address'],
                              ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_read_empty": (BASE, ['scratch.Range("B3").NumberFormat = "0.00"', 'v = scratch.Range("Q9").Value'],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_read_first": (BASE, ['scratch.Range("C1").NumberFormat = "0.0%"', 'scratch.Range("C2").NumberFormat = "0%"',
                                   'scratch.Range("A1").NumberFormat = "General"',
                                   'v = scratch.Range("C1:C2").NumberFormat'],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_style": (BASE, ['scratch.Range("B4").NumberFormat = "0.0"', 'scratch.Range("B5").Style = "Percent"'],
                       ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_font": (BASE, ['scratch.Range("B6").NumberFormat = "0.0"', 'scratch.Range("B7").Font.Bold = True'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_fill": (BASE, ['scratch.Range("B8").NumberFormat = "0.00000"', 'scratch.Range("A1").NumberFormat = "0.0"',
                             'scratch.Range("B8").Interior.Color = RGB(1, 2, 3)'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_write_empty": (BASE, ['scratch.Range("B9").NumberFormat = "0.0"', 'scratch.Range("B10").Value = 5'],
                             ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_write_kept": (BASE, ['scratch.Range("B11").NumberFormat = "0.00000"', 'scratch.Range("B11").Value = 5',
                                   'scratch.Range("A1").NumberFormat = "0.0"', 'scratch.Range("B11").Value = 6'],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_clear_contents": (BASE, ['scratch.Range("B12").NumberFormat = "0.00"', 'scratch.Range("B12").Value = 5',
                                       'scratch.Range("A1").NumberFormat = "0.0"', 'scratch.Range("B12").ClearContents'],
                                ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_entire_row": (BASE, ['scratch.Range("A13").NumberFormat = "0.000"', 'scratch.Range("A13").Value = 5',
                                   'scratch.Range("B1").NumberFormat = "0.0"', 'v = scratch.Rows(13).OutlineLevel'],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_sort": (BASE, ['scratch.Range("A1").NumberFormat = "0.0"',
                             'ws.Range("A1:C6").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes'],
                      ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_copied": ("Region,Product,Sales!East,A,10!East,B,20!West,A,30", [
        'ws.Range("A2:A3").NumberFormat = "0.00"', 'scratch.Range("A1").NumberFormat = "0.000"'],
        ['ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "register_above": (BASE, ['scratch.Range("A1").NumberFormat = "0.000"'],
                       ['ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove'], []),
    "register_nested": (NEST, ['scratch.Range("A1").NumberFormat = "0.000"'],
                        ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                         'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
    "register_after_subtotal": ("Key,Sales!a,1!b,2!!!Key,Sales!c,3!d,4", [
        'ws.Range("A3").NumberFormat = "0.000"', 'ws.Names.Add "Second", "=\'" & ws.Name & "\'!$A$6"',
        'scratch.Range("A1").NumberFormat = "0.0"'],
        ['ws.Range("A1").Subtotal 1, xlSum, Array(2)', 'ws.Range("Second").Subtotal 1, xlSum, Array(2)'], []),
    "register_remove": (BASE, ['scratch.Range("A1").NumberFormat = "0.0"'],
                        ['ws.Range("A1").RemoveSubtotal', 'ws.Range("A1").Subtotal 1, xlSum, Array(3)'], []),
    "groupby_third_total": ("Kind,Region,Product,Sales!k,East,A,10!k,East,B,20!k,West,A,30", [],
                            ['ws.Range("A1").Subtotal 3, xlCount, Array(3, 4)'], []),
    "groupby_total_left_taken": (BASE, [], ['ws.Range("A1").Subtotal 2, xlCount, Array(1, 2, 3)'], []),
    "values_blanks_above": ("Key,Sales!East,1!East,2!,3!East,4!,5!,6!West,7", [],
                            ['ws.Range("A1").Subtotal 1, xlSum, Array(2), SummaryBelowData:=xlSummaryAbove'], []),
    "nested_above_count": (NEST, [], [
        'ws.Range("A1").Subtotal 1, xlSum, Array(3), SummaryBelowData:=xlSummaryAbove',
        'ws.Range("A1").Subtotal 2, xlCount, Array(3), Replace:=False'], []),
    "remove_formula_plus": (BASE, ['ws.Range("D3").Formula = "=SUBTOTAL(9,C2:C3)+1"'],
                            ['ws.Range("A1").RemoveSubtotal'], []),
    "groupby_total_skip_left": ("Kind,Region,Product,Sales!k,East,A,10!k,East,B,20!k,West,A,30", [],
                                ['ws.Range("A1").Subtotal 3, xlCount, Array(2, 3)'], []),
    "nested_grand_again": (NEST, [], ['ws.Range("A1").Subtotal 1, xlSum, Array(3)',
                                      'ws.Range("A1").Subtotal 2, xlAverage, Array(3), Replace:=False',
                                      'ws.Range("A1").Subtotal 2, xlSum, Array(3), Replace:=False'], []),
})

for code in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, -4100, 1000):
    CASES[f"fn_code_{code}".replace("-", "minus_")] = (BASE, [], [f'ws.Range("A1").Subtotal 1, {code}, Array(3)'], [])

HELPER = '''Private Sub Fill(ws As Object, rows As String)
    Dim r As Long, c As Long, line As Variant, part As Variant
    For Each line In Split(rows, "!")
        r = r + 1
        c = 0
        For Each part In Split(line, ",")
            c = c + 1
            If part <> "" Then ws.Cells(r, c).Formula = part
        Next
    Next
End Sub

Private Function Shown(cell As Object) As String
    Dim v As Variant
    v = cell.Value
    If IsError(v) Then
        Shown = "Error:" & cell.Text
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Dump(ws As Object) As String
    Dim cell As Object, out As String, r As Long, last As Long, right As Long, fill As String
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    right = ws.UsedRange.Column + ws.UsedRange.Columns.Count
    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(last, right)).Cells
        fill = ""
        If cell.Interior.ColorIndex <> xlColorIndexNone Then fill = CStr(cell.Interior.Color)
        If Not IsEmpty(cell.Value) Or cell.HasFormula Or cell.Font.Bold Or cell.Font.Italic Or _
                cell.NumberFormat <> "General" Or fill <> "" Then
            out = out & cell.Address(False, False) & "=" & cell.Formula & "|" & Shown(cell) & "|" & _
                IIf(cell.Font.Bold, "b", "") & IIf(cell.Font.Italic, "i", "") & "|" & cell.NumberFormat & "|" & fill & ";"
        End If
    Next
    out = out & "#"
    For r = 1 To last
        out = out & ws.Rows(r).OutlineLevel & IIf(ws.Rows(r).Hidden, "h", "") & " "
    Next
    Dump = out & "#" & ws.Outline.SummaryRow
End Function
'''

BUILD = HELPER + '''{functions}
Public Function Build(folder As String) As String
    Dim wb As Object, ws As Object, scratch As Object, out As String, name As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set scratch = wb.Worksheets(1)
    scratch.Name = "scratch"
    scratch.Range("A1").NumberFormat = "General"
{cases}
    wb.Worksheets(1).Activate
    name = "subtotals.xlsx"
    wb.SaveAs folder & "\\" & name, 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\" & name)
    out = out & "@@"
    For Each ws In wb.Worksheets
        out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    Build = out
End Function
'''

#: An application's handlers, which log the events a Subtotal and a RemoveSubtotal raise.
SINK = '''Public WithEvents App As Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "change:" & Target.Address(False, False) & "|"
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
Log = Log & "calculate:" & Sh.Name & "|"
End Sub
'''

WATCH = HELPER + '''
Public Function Watch(dependent As Boolean) As String
    Dim wb As Object, ws As Object, sink As New SubtotalSink
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Fill ws, "{base}"
    If dependent Then ws.Range("F1").Formula = "=SUM(C:C)"
    Set sink.App = Application
    Application.EnableEvents = True
    Log = ""
    ws.Range("A1").Subtotal 1, xlSum, Array(3)
    Log = Log & "then|"
    ws.Range("A1").Subtotal 2, xlCount, Array(3), Replace:=False
    Log = Log & "then|"
    ws.Range("A1").RemoveSubtotal
    Application.EnableEvents = False
    Watch = Log
    wb.Close False
End Function
'''


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def case_code(index: int, name: str, rows: str, setup: list[str], calls: list[str], after: list[str]) -> str:
    """A function of its own for one case, which a procedure's size limit keeps out of Build."""
    lines = [f"Private Function Make{index}(wb As Object, scratch As Object) As String",
             "    Dim ws As Object, notes As String, v As Variant",
             "    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
             f'    ws.Name = "{name}"', f"    Fill ws, {vba_text(rows)}"]
    lines += [f"    {statement}" for statement in setup]
    lines.append("    On Error Resume Next")
    for call in calls:
        lines += ["    Err.Clear", f"    {call}",
                  '    If Err.Number <> 0 Then notes = notes & "E" & Err.Number & ":" & Err.Description & " "']
    lines.append("    On Error GoTo 0")
    lines.append(f'    Make{index} = "{name}^" & notes & "@" & Dump(ws) & "~"')
    lines += [f"    {statement}" for statement in after]
    # Each case starts from General as the number format a first label takes in General's place, whatever the
    # reads above left it at.
    lines.append('    scratch.Range("A1").NumberFormat = "General"')
    return "\n".join([*lines, "End Function"]) + "\n"


def sheet_parts(path: Path) -> dict[str, dict[str, object]]:
    """Each sheet's sheetPr, sheetFormatPr, rows' attributes and rowBreaks as the file keeps them, by the sheet's
    name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, dict[str, object]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            out[name] = parts(package.read(part).decode("utf-8"))
        return out


def parts(text: str) -> dict[str, object]:
    """A sheet's outline settings, its rows' outline attributes and its page breaks."""
    found = re.search(r"<sheetPr\b[^>]*/>|<sheetPr\b.*?</sheetPr>", text, re.DOTALL)
    outline = re.search(r"<outlinePr\b[^>]*/>", found.group()) if found else None
    format_found = re.search(r"<sheetFormatPr\b[^>]*/>", text)
    levels = {}
    for row in re.finditer(r"<row\b([^>]*?)/?>", text):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', row.group(1)))
        kept = {key: attributes[key] for key in ("outlineLevel", "hidden", "collapsed") if key in attributes}
        if kept:
            levels[attributes["r"]] = kept
    breaks = re.search(r"<rowBreaks\b.*?</rowBreaks>|<rowBreaks\b[^>]*/>", text, re.DOTALL)
    outline_level = re.search(r'\boutlineLevelRow="(\d+)"', format_found.group()) if format_found else None
    return {"outlinePr": outline.group() if outline else "",
            "outlineLevelRow": outline_level.group(1) if outline_level else "",
            "rows": levels, "rowBreaks": breaks.group() if breaks else ""}


def install(code: str, name: str) -> str:
    """A macro adding a class module of ``code`` to the workbook the harness runs in."""
    lines = ["Function Install() As String", "Dim text As String, part As Object",
             "Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)", f'part.Name = "{name}"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
    return "\n".join([*lines, "part.CodeModule.AddFromString text", 'Install = "ok"', "End Function"])


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    functions = "".join(case_code(index, name, *case) for index, (name, case) in enumerate(CASES.items()))
    calls = "".join(f"    out = out & Make{index}(wb, scratch)\n" for index in range(len(CASES)))
    code = BUILD.replace("{functions}", functions).replace("{cases}", calls)
    watch = WATCH.replace("{base}", BASE)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", (str(OUT),), timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        excel.add_module("EventLog", "Public Log As String\n")
        installed = excel.run_vba(install(SINK, "SubtotalSink"), "Install", timeout=60.0)
        assert installed.ok, f"{installed.outcome}: {installed.message}"
        events = {}
        for dependent in (True, False):
            watched = excel.run_vba(watch, "Watch", (dependent,), timeout=300.0)
            assert watched.ok, f"{watched.outcome}: {watched.message} {watched.error}"
            events["dependent" if dependent else "plain"] = str(watched.value)
    session, reopened = str(result.value).split("@@", 1)
    path = OUT / "subtotals.xlsx"
    strip_save_path(path)
    record = {"build": code, "sink": SINK, "watch": watch,
              "cases": [{"name": name, "rows": rows, "setup": setup, "calls": calls, "after": after}
                        for name, (rows, setup, calls, after) in CASES.items()],
              "session": dict(one.split("^", 1) for one in session.split("~") if "^" in one),
              "reopened": dict(one.split("^", 1) for one in reopened.split("~") if "^" in one),
              "parts": sheet_parts(path), "events": events}
    (OUT / "subtotals.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, dump in record["session"].items():
        print(f"{name}: {dump}")
        print()
    print("events:", record["events"])


if __name__ == "__main__":
    main()
