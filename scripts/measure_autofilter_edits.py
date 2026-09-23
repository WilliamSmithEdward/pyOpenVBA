"""What a filtered sheet does to the edits a macro makes around its filter.

Every layout sits on a sheet of its own in a new workbook: a table of
names and numbers filtered to its apples, then an edit -- a write, a
format, a clear, a fill, a copy, a cut, a sort, rows or columns deleted
or inserted -- and a dump of A1:G8 (each cell's formula) with which of
rows 1 to 14 are hidden, the filter's range and criteria, the
_FilterDatabase name, what a copy landed, and what the last call
returned or the error it raised.

    python scripts/measure_autofilter_edits.py

writes tests/fixtures/autofilter_edits.json, which
tests/test_excel_autofilter_edits.py replays.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from measure_autofilter import HELPER as FILTER_HELPER  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "autofilter_edits.json"

HELPER = FILTER_HELPER.replace("Private Function Dump(ws As Object) As String", "Private Function Rows(ws As Object) As String") \
    .replace('    Dump = "hidden=" & out & " " & Filters(ws)', '    Rows = "hidden=" & out & " " & Filters(ws)') + '''
Private Function Cells8(ws As Object) As String
    Dim out As String, c As Object
    For Each c In ws.Range("A1:G8").Cells
        out = out & c.Formula & "<>"
    Next
    Cells8 = out
End Function

Private Function Dump(ws As Object) As String
    Dim landed As String, other As Object
    On Error Resume Next
    Set other = Nothing
    Set other = ws.Parent.Worksheets(ws.Name & "_to")
    If Not other Is Nothing Then landed = " landed=" & other.UsedRange.Address & ":" & Cells8(other)
    On Error GoTo 0
    Dump = Rows(ws) & " cells=" & Cells8(ws) & landed
End Function
'''

#: The table: Name and Qty over rows 2 to 7, a formula beside it, filtered to the apples.
TABLE = '''ws.Range("A1:C1").Value = Array("Name", "Qty", "Twice")
ws.Range("A2:B2").Value = Array("apple", 1)
ws.Range("A3:B3").Value = Array("pear", 2)
ws.Range("A4:B4").Value = Array("apple", 3)
ws.Range("A5:B5").Value = Array("fig", 4)
ws.Range("A6:B6").Value = Array("apple", 5)
ws.Range("A7:B7").Value = Array("kiwi", 6)
ws.Range("C2:C7").Formula = "=B2*2"
ws.Range("D1").Formula = "=SUM(B2:B7)"
ws.Range("F1:F7").Value = Application.Transpose(Array("f1", "f2", "f3", "f4", "f5", "f6", "f7"))
ws.Range("A1:C7").AutoFilter Field:=1, Criteria1:="apple"
'''

#: The same filter again, after a setup that cleared it.
APPLES = 'ws.Range("A1:C7").AutoFilter Field:=1, Criteria1:="apple"\n'

#: Kiwi instead: rows 2 to 6 hidden in one block, only row 7 shown.
KIWI = 'ws.Range("A1:C7").AutoFilter Field:=1, Criteria1:="kiwi"\n'

#: Qty over 2 instead: rows 2 and 3 hidden, rows 4 to 7 shown in one block.
QTY = 'ws.ShowAllData\nws.Range("A1:C7").AutoFilter Field:=2, Criteria1:=">2"\n'

#: Arrays whose every item says where it came from: 11 is row 1, column 1.
GRID2 = "{11,12;21,22;31,32;41,42;51,52;61,62}"
GRID3 = "{11,12,13;21,22,23;31,32,33;41,42,43;51,52,53;61,62,63}"


def target() -> str:
    """A sheet to copy to, named after the layout's own."""
    return 'Set dest = ws.Parent.Worksheets.Add(After:=ws)\ndest.Name = ws.Name & "_to"\nws.Activate\n'


#: name -> setup; each works on ws, the layout's own sheet, after the table.
LAYOUTS: dict[str, str] = {
    "visible_cells": 'v = ws.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Address',
    "visible_body": 'v = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Address',
    "copy_filtered": target() + 'v = ws.AutoFilter.Range.Copy(dest.Range("A1"))',
    "copy_part": target() + 'v = ws.Range("A2:B6").Copy(dest.Range("A1"))',
    "copy_column": target() + 'v = ws.Range("B1:B7").Copy(dest.Range("A1"))',
    "copy_outside": target() + 'v = ws.Range("D1:D3").Copy(dest.Range("A1"))',
    "copy_to_clipboard": target() + 'ws.AutoFilter.Range.Copy\ndest.Paste dest.Range("A1")',
    "copy_hidden_by_hand": 'ws.AutoFilterMode = False\nws.Rows(3).Hidden = True\n' + target() +
                           'v = ws.Range("A1:C7").Copy(dest.Range("A1"))',
    "copy_arrows_only": 'ws.ShowAllData\nws.Rows(3).Hidden = True\n' + target() +
                        'v = ws.Range("A1:C7").Copy(dest.Range("A1"))',
    "clear_filtered": 'ws.Range("A2:B7").ClearContents',
    "value_filtered": 'ws.Range("B2:B7").Value = 9',
    "delete_visible": 'ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete',
    "delete_rows_through": 'ws.Range("A3:A5").EntireRow.Delete',
    "delete_one_row": 'ws.Rows(4).Delete',
    "delete_header": 'ws.Rows(1).Delete',
    "delete_above": 'ws.Range("F1").Value = 1\nws.Rows(1).Insert\nws.Rows(1).Delete',
    "insert_inside": 'ws.Rows(3).Insert',
    "insert_above": 'ws.Rows(1).Insert',
    "insert_below": 'ws.Rows(8).Insert',
    "delete_filtered_column": 'ws.Columns(1).Delete',
    "delete_other_column": 'ws.Columns(2).Delete',
    "insert_column": 'ws.Columns(2).Insert',
    "delete_all_body": 'ws.Rows("2:7").Delete',
    "range_delete_up": 'ws.Range("A3:C3").Delete Shift:=xlUp',
    "entire_row_filtered": 'ws.Range("A2:A7").EntireRow.Hidden = False',
    "value_default_member": 'ws.Range("B2:B7") = 9',
    "value_array": 'ws.Range("B2:B7").Value = Application.Transpose(Array(11, 12, 13, 14, 15, 16))',
    "value2_filtered": 'ws.Range("B2:B7").Value2 = 9',
    "formula_filtered": 'ws.Range("B2:B7").Formula = "=ROW()"',
    "value_one_hidden_cell": 'ws.Range("B3").Value = 9',
    "value_outside_columns": 'ws.Range("E2:E7").Value = 1\nv = ws.Range("E2").Formula & "," & ws.Range("E3").Formula',
    "value_resized": 'ws.Range("B2").Resize(6, 1).Value = 9',
    "number_format_filtered": 'ws.Range("B2:B7").NumberFormat = "0.00"\n'
                              'v = ws.Range("B2").NumberFormat & "," & ws.Range("B3").NumberFormat',
    "bold_filtered": 'ws.Range("A2:A7").Font.Bold = True\nv = ws.Range("A2").Font.Bold & "," & ws.Range("A3").Font.Bold',
    "clear_all_filtered": 'ws.Range("A2:B7").Clear',
    "fill_down_filtered": 'ws.Range("B2:B7").FillDown',
    "rows_hidden_true": 'ws.Rows(4).Hidden = True',
    "rows_hidden_false_one": 'ws.Rows(3).Hidden = False',
    "delete_hidden_row": 'ws.Rows(3).Delete',
    "insert_entire_row_visible": 'ws.Range("A4").EntireRow.Insert',
    "clear_outside_filter": 'ws.ShowAllData\nws.Range("E2:E7").Value = 1\n' + APPLES +
                            'ws.Range("E2:E7").ClearContents\nv = ws.Range("E2").Formula & "," & ws.Range("E3").Formula',
    "copy_single_area_visible": target() + 'v = ws.Range("A1:C2").Copy(dest.Range("A1"))',
    "copy_multi_areas_no_formula": target() + 'v = ws.Range("A1:B7").Copy(dest.Range("A1"))',
    "value_hidden_row_cells": 'ws.Range("A3:C3").Value = 9',
    "kiwi_value_hidden_block": KIWI + 'ws.Range("B3:B5").Value = 9',
    "kiwi_value_spanning": KIWI + 'ws.Range("B5:B7").Value = 9',
    "kiwi_unhide_block": KIWI + 'ws.Rows("3:5").Hidden = False',
    "kiwi_delete_block": KIWI + 'ws.Rows("3:4").Delete',
    "kiwi_clear_block": KIWI + 'ws.Range("A3:B5").ClearContents',
    "value_hidden_by_hand": 'ws.AutoFilterMode = False\nws.Rows(3).Hidden = True\nws.Range("B2:B7").Value = 9',
    "value_filter_and_hand": 'ws.Rows(2).Hidden = True\nws.Range("B2:B7").Value = 9',
    "value_below_filter_hand": 'ws.Rows(9).Hidden = True\nws.Range("B8:B10").Value = 9\n'
                               'v = ws.Range("B8").Formula & "," & ws.Range("B9").Formula & "," & ws.Range("B10").Formula',
    "delete_below_filter_hand": 'ws.Rows(9).Hidden = True\nws.Range("A8:A10").Value = Application.Transpose(Array("x", "y", "z"))\n'
                                'ws.Rows("8:10").Delete\nv = ws.Range("A8").Formula & "," & ws.Range("A9").Formula',
    "copy_filter_and_hand": 'ws.Rows(2).Hidden = True\n' + target() + 'v = ws.Range("A1:C7").Copy(dest.Range("A1"))',
    "cut_filtered": target() + 'v = ws.Range("A1:C7").Cut(dest.Range("A1"))\nv = Rows(dest)',
    "copy_paste_formulas": target() + 'ws.AutoFilter.Range.Copy\ndest.Range("A1").PasteSpecial xlPasteFormulas',
    "copy_paste_values": target() + 'ws.AutoFilter.Range.Copy\ndest.Range("A1").PasteSpecial xlPasteValues',
    "paste_into_filtered": 'Set dest = ws.Parent.Worksheets.Add(After:=ws)\n'
                           'dest.Range("A1:A3").Value = Application.Transpose(Array(7, 8, 9))\n'
                           'v = dest.Range("A1:A3").Copy(ws.Range("B2"))',
    "read_filtered": 'v = Join(Application.Transpose(ws.Range("B2:B7").Value), ",") & ";" & '
                     'ws.Range("B2:B7").Cells.Count & ";" & Application.Sum(ws.Range("B2:B7"))',
    "delete_whole_filter_rows": 'ws.Range("A1:C7").EntireRow.Delete',
    "range_delete_filter": 'ws.AutoFilter.Range.Delete',
    "range_delete_up_multi": 'ws.Range("A2:C4").Delete Shift:=xlUp',
    "range_insert_down_multi": 'ws.Range("A2:C4").Insert Shift:=xlDown',
    "insert_rows_through": 'ws.Range("A3:A5").EntireRow.Insert',
    "autofill_filtered": 'ws.Range("B2").Value = 10\nws.Range("B2").AutoFill ws.Range("B2:B7"), xlFillSeries',
    "autofill_copy": 'ws.Range("B2").AutoFill ws.Range("B2:B7"), xlFillCopy',
    "formula_r1c1_filtered": 'ws.Range("B2:B7").FormulaR1C1 = "=ROW()"',
    "cells_clear_contents": 'ws.Cells.ClearContents',
    "rows_clear_contents": 'ws.Rows("2:7").ClearContents',
    "unhide_all_rows": 'ws.Cells.EntireRow.Hidden = False',
    "rows_hidden_false_all": 'ws.Rows.Hidden = False',
    # AutoFilter.Range read after writing around the filter.
    "ext_text_below": 'ws.Range("A8").Value = "x"',
    "ext_outside_columns": 'ws.Range("D8").Value = 1',
    "ext_far_column": 'ws.Range("E8").Value = 1',
    "ext_gap_row": 'ws.Range("B9").Value = 1',
    "ext_two_rows": 'ws.Range("B9").Value = 1\nws.Range("B8").Value = 1',
    "ext_formula": 'ws.Range("B8").Formula = "=1+1"',
    "ext_empty_string": 'ws.Range("B8").Value = ""',
    "ext_arrows_only": 'ws.ShowAllData\nws.Range("B8").Value = 9',
    "ext_format_only": 'ws.Range("B8").NumberFormat = "0.00"',
    "ext_then_clear": 'ws.Range("B8").Value = 9\nws.Range("B8").ClearContents',
    "ext_reapply": 'ws.Range("A8").Value = "apple"\nws.Range("A9").Value = "pear"\nws.AutoFilter.ApplyFilter',
    "ext_refilter": 'ws.Range("A8").Value = "pear"\nws.Range("A1").AutoFilter Field:=1, Criteria1:="apple"',
    "ext_no_filter": 'ws.AutoFilterMode = False\nws.Range("A1:C7").AutoFilter\nws.Range("B8").Value = 9',
    # Other edits in filter mode.
    "row_height_filtered": 'ws.Rows("2:7").RowHeight = 30\nws.ShowAllData\nv = ws.Rows(2).RowHeight & "," & ws.Rows(3).RowHeight',
    "replace_filtered": 'v = ws.Range("A2:A7").Replace("i", "I")',
    "sort_filtered": 'ws.Range("A1:C7").Sort Key1:=ws.Range("B1"), Order1:=xlDescending, Header:=xlYes',
    "remove_duplicates_filtered": 'ws.Range("A1:C7").RemoveDuplicates Columns:=1, Header:=xlYes',
    "special_constants": 'v = ws.Range("A1:C7").SpecialCells(xlCellTypeConstants).Address',
    "subtotal_filtered": 'ws.Range("E1").Formula = "=SUBTOTAL(9,B2:B7)&"",""&SUBTOTAL(109,B2:B7)&"",""&SUBTOTAL(3,A2:A7)"\n'
                         'v = ws.Range("E1").Value & ";" & Application.WorksheetFunction.Subtotal(9, ws.Range("B2:B7"))',
    "subtotal_by_hand": 'ws.AutoFilterMode = False\nws.Rows(3).Hidden = True\n'
                        'ws.Range("E1").Formula = "=SUBTOTAL(9,B2:B7)&"",""&SUBTOTAL(109,B2:B7)&"",""&SUBTOTAL(3,A2:A7)"\n'
                        'v = ws.Range("E1").Value',
    "interior_filtered": 'ws.Range("A2:A7").Interior.Color = 255\nv = ws.Range("A2").Interior.Color & "," & ws.Range("A3").Interior.Color',
    "alignment_filtered": 'ws.Range("A2:A7").HorizontalAlignment = xlCenter\n'
                          'v = ws.Range("A2").HorizontalAlignment & "," & ws.Range("A3").HorizontalAlignment',
    "clear_formats_filtered": 'ws.ShowAllData\nws.Range("B2:B7").NumberFormat = "0.00"\n' + APPLES +
                              'ws.Range("B2:B7").ClearFormats\nv = ws.Range("B2").NumberFormat & "," & ws.Range("B3").NumberFormat',
    "fill_right_filtered": 'ws.Range("A2:C4").FillRight',
    "fill_up_filtered": 'ws.Range("B2:B7").FillUp',
    "fill_down_first_hidden": 'ws.Range("B3:B7").FillDown',
    "fill_down_kiwi": KIWI + 'ws.Range("B2:B7").FillDown',
    "value_self_assign": 'ws.Range("A2:B7").Value = ws.Range("A2:B7").Value',
    "value_array_area": QTY + 'ws.Range("B2:B7").Value = Application.Transpose(Array(11, 12, 13, 14, 15, 16))',
    "value_row_array": 'ws.Range("A2:B7").Value = Array(7, 8)',
    "value_union_areas": 'ws.Range("B2,B3").Value = 9',
    "value_arrows_only": 'ws.ShowAllData\nws.Rows(3).Hidden = True\nws.Range("B2:B7").Value = 9',
    # Structural edits in filter mode.
    "ins_rows_2_4": 'ws.Range("A2:A4").EntireRow.Insert',
    "ins_rows_3_7": 'ws.Range("A3:A7").EntireRow.Insert',
    "ins_rows_2_7": 'ws.Rows("2:7").Insert',
    "ins_kiwi_block": KIWI + 'ws.Rows("3:4").Insert',
    "ins_cells_visible_row": 'ws.Range("A4:C4").Insert Shift:=xlDown',
    "ins_cells_hidden_row": 'ws.Range("A3:C3").Insert Shift:=xlDown',
    "ins_cells_outside": 'ws.Range("E2:E4").Insert Shift:=xlDown',
    "ins_cells_right": 'ws.Range("A2:A4").Insert Shift:=xlToRight',
    "del_cells_left": 'ws.Range("A2:A4").Delete Shift:=xlToLeft',
    "del_cells_outside": 'ws.Range("E2:F4").Delete Shift:=xlUp',
    "del_cells_part_columns": 'ws.Range("A2:B4").Delete Shift:=xlUp',
    "del_range_no_shift": 'ws.Range("A2:C4").Delete',
    "del_cells_below": 'ws.Range("A8:C9").Delete Shift:=xlUp',
    "del_filter_and_hand": 'ws.Rows(2).Hidden = True\nws.Rows("2:7").Delete',
    "field3_delete_col2": 'ws.ShowAllData\nws.Range("A1:C7").AutoFilter Field:=3, Criteria1:=">5"\nws.Columns(2).Delete',
    "delete_all_filter_columns": 'ws.Columns("A:C").Delete',
    "delete_two_columns": 'ws.Columns("B:C").Delete',
    "insert_column_before": 'ws.Columns(1).Insert',
    "insert_column_after": 'ws.Columns(4).Insert',
    # Cut, and the paste flavours of a filtered copy.
    "cut_part": target() + 'ws.Range("A2:C4").Cut dest.Range("A1")',
    "cut_same_sheet": 'ws.Range("A2:C2").Cut ws.Range("A10")\nv = ws.Range("A10").Formula & "," & ws.Range("C10").Formula',
    "cut_whole_same_sheet": 'ws.Range("A1:C7").Cut ws.Range("A11")',
    "copy_paste_all": target() + 'ws.AutoFilter.Range.Copy\ndest.Range("A1").PasteSpecial xlPasteAll',
    "copy_paste_formulas_formats": target() + 'ws.AutoFilter.Range.Copy\ndest.Range("A1").PasteSpecial xlPasteFormulasAndNumberFormats',
    "copy_same_sheet": 'ws.Range("A1:C7").Copy ws.Range("E1")',
    "paste_range_into_filtered": 'Set dest = ws.Parent.Worksheets.Add(After:=ws)\n'
                                 'dest.Range("A1:A3").Value = Application.Transpose(Array(7, 8, 9))\n'
                                 'v = dest.Range("A1:A3").Copy(ws.Range("B2:B7"))',
    "paste_special_into_filtered": 'Set dest = ws.Parent.Worksheets.Add(After:=ws)\n'
                                   'dest.Range("A1:A3").Value = Application.Transpose(Array(7, 8, 9))\n'
                                   'dest.Range("A1:A3").Copy\nws.Range("B2").PasteSpecial xlPasteValues',
    # Arrays written to several areas, filtered or not.
    "array_distinct": 'ws.Range("A2:B7").Value = Application.Evaluate("' + GRID2 + '")',
    "array_distinct_qty": QTY + 'ws.Range("A2:B7").Value = Application.Evaluate("' + GRID2 + '")',
    "array_distinct_3col": 'ws.Range("A2:C7").Value = Application.Evaluate("' + GRID3 + '")',
    "array_small": 'ws.Range("A2:B7").Value = Application.Evaluate("{11,12;21,22}")',
    "array_areas_no_filter": 'ws.AutoFilterMode = False\nws.Range("A2:B2,A4:B4,A6:B6").Value = Application.Evaluate("' + GRID2 + '")',
    "array_areas_column": 'ws.AutoFilterMode = False\nws.Range("B2,B4,B6").Value = Application.Evaluate("{11;21;31;41;51;61}")',
    "formula_array_filtered": 'ws.Range("B2:B7").Formula = Application.Evaluate("{""=1"";""=2"";""=3"";""=4"";""=5"";""=6""}")',
    # Cell deletes and inserts away from the filter.
    "del_cells_kiwi_block": KIWI + 'ws.Range("A3:C4").Delete Shift:=xlUp',
    "del_cells_below_rows": 'ws.Range("A9").Value = "a9"\nws.Range("F9").Value = "f9"\nws.Range("A8").Delete Shift:=xlUp\n'
                            'v = ws.Range("A8").Formula & "," & ws.Range("F8").Formula & "," & ws.Range("F9").Formula',
    "del_cells_far_rows": 'ws.Range("A20:A21").Value = 1\nws.Range("F21").Value = "f21"\nws.Range("A20").Delete Shift:=xlUp\n'
                          'v = ws.Range("A20").Formula & "," & ws.Range("F20").Formula & "," & ws.Range("F21").Formula',
    "del_cells_arrows_only": 'ws.ShowAllData\nws.Range("A3:C3").Delete Shift:=xlUp',
    "del_cells_left_hidden_row": 'ws.Range("A3").Delete Shift:=xlToLeft',
    "ins_cells_below": 'ws.Range("A9:C9").Insert Shift:=xlDown',
    "ins_cells_far": 'ws.Range("H20").Insert Shift:=xlDown',
    "ins_cells_arrows_only": 'ws.ShowAllData\nws.Range("A4:C4").Insert Shift:=xlDown',
    "ins_cells_right_far": 'ws.Range("H3").Insert Shift:=xlToRight',
    "ins_entire_row_below": 'ws.Range("A9").EntireRow.Insert',
    # Cut and the references around what it moves, filtered or not.
    "cut_part_no_filter": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("A2:C4").Cut dest.Range("A1")',
    "cut_bottom_no_filter": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("A5:C7").Cut dest.Range("A1")',
    "cut_middle_no_filter": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("A3:C4").Cut dest.Range("A1")',
    "cut_same_sheet_no_filter": 'ws.AutoFilterMode = False\nws.Range("A2:C4").Cut ws.Range("A10")',
    "cut_column_part_no_filter": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("B2:B4").Cut dest.Range("A1")',
    "cut_column_whole_no_filter": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("B2:B7").Cut dest.Range("A1")',
    # RemoveDuplicates, Merge and Borders around a filter.
    "remove_duplicates_arrows": 'ws.ShowAllData\nws.Range("A1:C7").RemoveDuplicates Columns:=1, Header:=xlYes',
    "remove_duplicates_hand": 'ws.AutoFilterMode = False\nws.Rows(3).Hidden = True\n'
                              'ws.Range("A1:C7").RemoveDuplicates Columns:=1, Header:=xlYes',
    "remove_duplicates_part": 'ws.Range("A2:C7").RemoveDuplicates Columns:=1, Header:=xlNo',
    "merge_filtered": 'ws.Range("E2:E4").Merge\nv = ws.Range("E2").MergeArea.Address & "," & ws.Range("E4").MergeArea.Address',
    "borders_filtered": 'ws.Range("A2:A7").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                        'v = ws.Range("A2").Borders(xlEdgeBottom).LineStyle & "," & ws.Range("A4").Borders(xlEdgeBottom).LineStyle'
                        ' & "," & ws.Range("A6").Borders(xlEdgeBottom).LineStyle & "," & ws.Range("A7").Borders(xlEdgeBottom).LineStyle',
    "delete_one_of_two_fields": 'ws.Range("A1:C7").AutoFilter Field:=2, Criteria1:=">1"\nws.Columns(1).Delete',
    # More of AutoFilter.Range growing over what is written below it.
    "ext_far_far": 'ws.Range("Z8").Value = 1',
    "ext_formula_blank": 'ws.Range("B8").Formula = "="""""',
    "ext_hidden_row": 'ws.Rows(8).Hidden = True\nws.Range("B8").Value = 1',
    # Copying several areas with no filter at all.
    "copy_areas_stacked": 'ws.AutoFilterMode = False\n' + target() + 'v = ws.Range("A1:C2,A4:C4").Copy(dest.Range("A1"))',
    "copy_areas_side": 'ws.AutoFilterMode = False\n' + target() + 'v = ws.Range("A1:A3,C1:C3").Copy(dest.Range("A1"))',
    "copy_areas_paste_all": 'ws.AutoFilterMode = False\n' + target() + 'ws.Range("A1:C2,A4:C4").Copy\n'
                            'dest.Range("A1").PasteSpecial xlPasteAll',
    "copy_areas_bad": 'ws.AutoFilterMode = False\n' + target() + 'v = ws.Range("A1:B2,C4:C5").Copy(dest.Range("A1"))',
    # Loose ends.
    "ext_refilter_explicit": 'ws.Range("A8").Value = "pear"\nws.Range("A1:C7").AutoFilter Field:=1, Criteria1:="apple"',
    "ext_show_all": 'ws.Range("A8").Value = "x"\nws.Rows(8).Hidden = True\nws.ShowAllData',
    "paste_special_range_into_filtered": 'Set dest = ws.Parent.Worksheets.Add(After:=ws)\n'
                                        'dest.Range("A1:A3").Value = Application.Transpose(Array(7, 8, 9))\n'
                                        'dest.Range("A1:A3").Copy\nws.Range("B2:B7").PasteSpecial xlPasteValues',
    "copy_formats_filtered": 'ws.ShowAllData\nws.Range("B4").NumberFormat = "0.00"\n' + APPLES + target() +
                             'v = ws.Range("A1:C7").Copy(dest.Range("A1"))\n'
                             'v = dest.Range("B3").NumberFormat & "," & dest.Range("B2").NumberFormat',
    "array_row_2d": 'ws.Range("A2:B7").Value = Application.Evaluate("{7,8}")',
    "remove_duplicates_outside": 'ws.ShowAllData\nws.Range("E1:E7").Value = Application.Transpose(Array("h", 1, 1, 2, 2, 3, 3))\n' +
                                 APPLES + 'ws.Range("E1:E7").RemoveDuplicates Columns:=1, Header:=xlYes',
    "sort_object_filtered": 'With ws.Sort\n.SortFields.Clear\n.SortFields.Add Key:=ws.Range("B2:B7"), Order:=xlDescending\n'
                            '.SetRange ws.Range("A1:C7")\n.Header = xlYes\n.Apply\nEnd With',
    "sort_hidden_block": KIWI + 'ws.Range("A2:C6").Sort Key1:=ws.Range("B2"), Order1:=xlDescending, Header:=xlNo',
    "del_narrow_no_shift": 'ws.Range("A2:A7").Delete',
    "delete_header_arrows": 'ws.ShowAllData\nws.Rows(1).Delete',
    "formula_relative_filtered": 'ws.Range("E2:E7").Formula = "=B2*10"\n'
                                 'v = ws.Range("E2").Formula & "," & ws.Range("E4").Formula & "," & ws.Range("E6").Formula',
    "formula_relative_first_hidden": 'ws.Range("E3:E7").Formula = "=B3*10"\n'
                                     'v = ws.Range("E3").Formula & "," & ws.Range("E4").Formula & "," & ws.Range("E6").Formula',
    "value_hidden_column": 'ws.Columns(5).Hidden = True\nws.Range("D2:F7").Value = 9\n'
                           'v = ws.Range("D2").Formula & "," & ws.Range("E2").Formula & "," & ws.Range("F3").Formula',
    "value_hidden_column_no_filter": 'ws.AutoFilterMode = False\nws.Columns(5).Hidden = True\nws.Range("D2:F7").Value = 9\n'
                                     'v = ws.Range("D2").Formula & "," & ws.Range("E2").Formula & "," & ws.Range("F3").Formula',
    "clear_hidden_column": 'ws.Columns(6).Hidden = True\nws.Range("E1:F7").ClearContents',
    # What clearing the header does to the filter.
    "clear_header": 'ws.Range("A1:C1").ClearContents',
    "clear_header_cell": 'ws.Range("A1").ClearContents',
    "clear_one_header_cell": 'ws.Range("B1").ClearContents',
    "clear_filter_range": 'ws.Range("A1:C7").ClearContents',
    "clear_header_arrows": 'ws.ShowAllData\nws.Range("A1:C1").ClearContents',
    "clear_header_no_criteria": 'ws.AutoFilterMode = False\nws.Range("A1:C7").AutoFilter\nws.Range("A1:C1").ClearContents',
    "value_header_empty": 'ws.Range("A1:C1").Value = ""',
    "clear_header_wider": 'ws.Range("A1:D1").ClearContents',
    "clear_rows_1_3": 'ws.Rows("1:3").ClearContents',
    "cells_clear": 'ws.Cells.Clear',
    "clear_header_formats": 'ws.Range("A1:C1").ClearFormats',
    "clear_header_then_refill": 'ws.Range("A1:C1").ClearContents\nws.Range("A1:C1").Value = Array("N", "Q", "T")',
    # Cut and copy odds and ends.
    "cut_formula_out": 'ws.Range("E2").Formula = "=D1+B2"\n' + target() + 'ws.Range("E2").Cut dest.Range("A1")\n'
                       'v = dest.Range("A1").Formula',
    "copy_areas_adjacent": 'ws.AutoFilterMode = False\n' + target() + 'v = ws.Range("A1:B3,C1:C3").Copy(dest.Range("A1"))',
}


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, dest As Object", "On Error Resume Next", "Err.Clear",
             *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        build += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", f'ws.Name = "{name[:28]}"',
                  f'out = out & Case{index}(ws) & "|"', "wb.Close False"]
        bodies.append(case_code(index, TABLE + setup))
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = str(result.value).split("|")[: len(LAYOUTS)]
    record = {"helper": HELPER, "table": TABLE,
              "layouts": [{"name": name, "setup": setup, "answers": answer}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, answer in zip(LAYOUTS, answers, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
