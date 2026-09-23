"""Row, column and sheet-wide formats as Excel keeps them, reads them and saves them.

Every case formats something on a sheet of its own in one new workbook,
and reads a list of properties back: the used range first, which also
makes Excel drop the empty row records an insert leaves until something
reads it. Excel then saves the workbook, opens it again and reads every
case once more. A second workbook holds row and column styles planted in
the XML the way another program might write them, with what Excel reads
from each and how it writes them back.

A read records the value's type and text, Null for a mixed range and
E<number> for an error, as the other probes do. Excel rounds heights to
whole pixels of the display it runs on, so this needs the 96-DPI display
the model emulates.

    python scripts/measure_row_formats.py

writes tests/fixtures/row_formats/: row_formats.xlsx and planted.xlsx as
Excel saved them, planted_source.xlsx as planted, and row_formats.json with
every case and answer. tests/test_excel_row_formats.py replays them.
"""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "row_formats"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''

B, IT, N, S = ".Font.Bold", ".Font.Italic", ".Font.Name", ".Font.Size"
T, BO = ".Borders(xlEdgeTop).LineStyle", ".Borders(xlEdgeBottom).LineStyle"
L, R = ".Borders(xlEdgeLeft).LineStyle", ".Borders(xlEdgeRight).LineStyle"
UP = ".Borders(xlDiagonalUp).LineStyle"


def at(reference: str) -> str:
    return f'ws.Range("{reference}")'


def cell_reads(reference: str) -> list[str]:
    """Every group a format has, read at one cell."""
    cell = at(reference)
    return [cell + B, cell + ".Interior.Color", cell + ".NumberFormat", cell + ".HorizontalAlignment",
            cell + ".Locked", cell + L]


def heights(*rows: int) -> list[str]:
    return [*(f"ws.Rows({row}).RowHeight" for row in rows), "ws.StandardHeight"]


C2 = at("B2")

#: name -> (setup, reads). Every setup works on ws, the case's own sheet.
CASES: dict[str, tuple[str, list[str]]] = {
    # --- a property set marks its part of the format as applied, whatever it sets ------------------
    "noop_bold_false": (f"{C2}.Font.Bold = False", cell_reads("B2")),
    "bold_unbold": (f"{C2}.Font.Bold = True\n{C2}.Font.Bold = False", cell_reads("B2")),
    "noop_name_same": (f'{C2}.Font.Name = "Aptos Narrow"', cell_reads("B2")),
    "noop_size_same": (f"{C2}.Font.Size = 11", cell_reads("B2")),
    "noop_color_auto": (f"{C2}.Font.ColorIndex = xlColorIndexAutomatic", cell_reads("B2")),
    "fill_back": (f"{C2}.Interior.Color = 255\n{C2}.Interior.ColorIndex = xlNone", cell_reads("B2")),
    "noop_fill_none": (f"{C2}.Interior.ColorIndex = xlNone", cell_reads("B2")),
    "numfmt_back": (f'{C2}.NumberFormat = "0.00"\n{C2}.NumberFormat = "General"', cell_reads("B2")),
    "noop_numfmt": (f'{C2}.NumberFormat = "General"', cell_reads("B2")),
    "align_back": (f"{C2}.HorizontalAlignment = xlCenter\n{C2}.HorizontalAlignment = xlGeneral", cell_reads("B2")),
    "noop_align": (f"{C2}.HorizontalAlignment = xlGeneral", cell_reads("B2")),
    "locked_back": (f"{C2}.Locked = False\n{C2}.Locked = True", cell_reads("B2")),
    "noop_locked": (f"{C2}.Locked = True", cell_reads("B2")),
    "border_back": (f"{C2}.Borders(xlEdgeLeft).LineStyle = xlContinuous\n{C2}.Borders(xlEdgeLeft).LineStyle = xlNone",
                    cell_reads("B2")),
    "noop_border": (f"{C2}.Borders(xlEdgeLeft).LineStyle = xlNone", cell_reads("B2")),
    "bold_clear_formats": (f"{C2}.Font.Bold = True\n{C2}.ClearFormats", cell_reads("B2")),
    "unbold_then_fill": (f"{C2}.Font.Bold = True\n{C2}.Font.Bold = False\n{C2}.Interior.Color = 65535",
                         cell_reads("B2")),
    "bold_then_fill": (f"{C2}.Font.Bold = True\n{C2}.Interior.Color = 65535", cell_reads("B2")),
    "value_noop_bold": (f"{C2}.Value = 1\n{C2}.Font.Bold = False", cell_reads("B2")),
    "value_bold_unbold": (f"{C2}.Value = 1\n{C2}.Font.Bold = True\n{C2}.Font.Bold = False", cell_reads("B2")),
    "unbold_clear_contents": (f"{C2}.Value = 1\n{C2}.Font.Bold = True\n{C2}.Font.Bold = False\n{C2}.ClearContents",
                              cell_reads("B2")),
    "range_noop_then_one": ('ws.Range("B2:C2").Font.Bold = False\nws.Range("C2").Font.Italic = True',
                            [at("C2") + IT, at("B2") + IT]),
    "copy_bold_unbold": (f'{C2}.Font.Bold = True\n{C2}.Font.Bold = False\n{C2}.Copy ws.Range("D4")', cell_reads("D4")),
    # --- whole rows and columns -------------------------------------------------------------------
    "row_bold": ('ws.Range("A3").Value = 1\nws.Range("C3").Interior.Color = RGB(255, 0, 0)\nws.Rows(3).Font.Bold = True\n'
                 'ws.Range("D3").Value = "new"',
                 [at("A3") + B, at("C3") + B, at("C3") + ".Interior.Color", at("Z3") + B, at("D3") + B,
                  "ws.Rows(3)" + B, at("A4") + B, at("XFD3") + B, "ws.Columns(26)" + B, "ws.Cells" + B]),
    "row_only": ("ws.Rows(3).Font.Bold = True", ["ws.Rows(3)" + B, 'ws.Range("A1:A5")' + B, 'ws.Range("B3:C3")' + B]),
    "column_only": ('ws.Columns("C").Font.Bold = True', ["ws.Columns(3)" + B, 'ws.Range("C1:C5")' + B,
                                                         'ws.Range("B1:C1")' + B]),
    "column_italic": ('ws.Range("C1").Value = 1\nws.Columns("C").Font.Italic = True\nws.Range("C10").Value = "new"',
                      [at("C1") + IT, at("C500") + IT, at("C10") + IT, 'ws.Columns("C")' + IT, at("D1") + IT,
                       at("C1048576") + IT]),
    "row_then_column": ('ws.Rows(5).Font.Bold = True\nws.Columns("E").Font.Italic = True',
                        [at("E5") + B, at("E5") + IT, at("F5") + IT, at("E6") + B, "ws.Rows(5)" + IT, "ws.Columns(5)" + B]),
    "column_then_row": ('ws.Columns("E").Font.Italic = True\nws.Rows(5).Font.Bold = True',
                        [at("E5") + B, at("E5") + IT, at("F5") + IT, at("E6") + B, "ws.Rows(5)" + IT, "ws.Columns(5)" + B]),
    "same_change_cross": ('ws.Columns("E").Font.Bold = True\nws.Rows(5).Font.Bold = True', [at("E5") + B]),
    "same_change_cross_rev": ('ws.Rows(5).Font.Bold = True\nws.Columns("E").Font.Bold = True', [at("E5") + B]),
    "value_in_row_and_column": ('ws.Rows(5).Font.Bold = True\nws.Columns("E").Font.Italic = True\n'
                                'ws.Range("F5").Value = "x"\nws.Range("E7").Value = "y"',
                                [at("F5") + B, at("F5") + IT, at("E7") + B, at("E7") + IT]),
    "row_noop": ("ws.Rows(3).Font.Bold = False", ["ws.Rows(3)" + B]),
    "row_bold_unbold": ("ws.Rows(3).Font.Bold = True\nws.Rows(3).Font.Bold = False", ["ws.Rows(3)" + B]),
    "column_noop": ("ws.Columns(3).Font.Bold = False", ["ws.Columns(3)" + B]),
    "column_bold_unbold": ("ws.Columns(3).Font.Bold = True\nws.Columns(3).Font.Bold = False", ["ws.Columns(3)" + B]),
    "rows_range": ('ws.Rows("3:5").Interior.Color = RGB(0, 255, 0)', [at("Q4") + ".Interior.Color",
                                                                    'ws.Rows("3:5").Interior.Color']),
    "whole_rows_by_address": ('ws.Range("A7:XFD8").Font.Bold = True', [at("Q8") + B]),
    "entire_row": ('ws.Range("B9").EntireRow.Font.Italic = True', [at("Q9") + IT]),
    "columns_range": ('ws.Columns("B:D").Font.Bold = True', [at("C9") + B, 'ws.Columns("B:D")' + B]),
    "row_number_format": ('ws.Rows(3).NumberFormat = "0.00"\nws.Range("B3").Value = 1.5',
                          [at("B3") + ".NumberFormat", at("B3") + ".Text", at("Z3") + ".NumberFormat"]),
    "column_then_value": ('ws.Columns("C").Font.Bold = True\nws.Range("C2").Value = 1', [at("C2") + B]),
    "row_style_value_far": ('ws.Rows(3).Font.Bold = True\nws.Range("XFD3").Value = 1', [at("XFD3") + B]),
    "column_hidden_style": ("ws.Columns(3).Hidden = True\nws.Columns(3).Font.Bold = True", ["ws.Columns(3).Hidden"]),
    "column_width_style": ("ws.Columns(3).ColumnWidth = 20\nws.Columns(3).Font.Bold = True",
                           ["ws.Columns(3).ColumnWidth"]),
    "style_then_width": ("ws.Columns(3).Font.Bold = True\nws.Columns(3).ColumnWidth = 20", ["ws.Columns(3).ColumnWidth"]),
    "row_height_style": ("ws.Rows(3).RowHeight = 30\nws.Rows(3).Font.Bold = True", ["ws.Rows(3).RowHeight"]),
    "row_hidden_style": ("ws.Rows(3).Hidden = True\nws.Rows(3).Font.Bold = True", ["ws.Rows(3).Hidden"]),
    "reads_mixed": ('ws.Rows(3).Font.Bold = True\nws.Columns(3).Font.Bold = True\nws.Range("E3").Font.Bold = False',
                    ['ws.Range("A1:E5")' + B, 'ws.Range("C1:C2")' + B, 'ws.Range("A3:D3")' + B, "ws.Rows(3)" + B,
                     "ws.Columns(3)" + B, "ws.Cells" + B, 'ws.Rows("3:4")' + B, 'ws.Columns("C:D")' + B]),
    "numfmt_rows_read": ('ws.Rows(3).NumberFormat = "0.00"',
                         ["ws.Rows(3).NumberFormat", 'ws.Rows("3:4").NumberFormat', "ws.Columns(1).NumberFormat",
                          "ws.Cells.NumberFormat"]),
    "interior_rows_read": ("ws.Rows(3).Interior.Color = 255",
                           ["ws.Rows(3).Interior.Color", 'ws.Rows("3:4").Interior.Color',
                            'ws.Rows("3:4").Interior.ColorIndex']),
    # --- when a cell with nothing in it goes -------------------------------------------------------------
    "cell_bold_unbold": (f"{C2}.Font.Bold = True\n{C2}.Font.Bold = False", [C2 + B]),
    "row_cell_unbold": ('ws.Rows(3).Font.Bold = True\nws.Range("C3").Font.Bold = False',
                        [at("C3") + B, "ws.Rows(3)" + B, 'ws.Range("A3:B3")' + B, 'ws.Range("B3:C3")' + B]),
    "row_cell_rebold": ('ws.Rows(3).Font.Bold = True\nws.Range("C3").Font.Bold = False\nws.Range("C3").Font.Bold = True',
                        [at("C3") + B, "ws.Rows(3)" + B]),
    "row_value_cleared": ('ws.Rows(3).Font.Bold = True\nws.Range("A3").Value = 1\nws.Range("A3").ClearContents',
                          [at("A3") + B]),
    "row_cell_clear": ('ws.Rows(3).Font.Bold = True\nws.Range("A3").Value = 1\nws.Range("A3").Clear', [at("A3") + B]),
    "row_cell_clear_formats": ('ws.Rows(3).Font.Bold = True\nws.Range("B3").ClearFormats', [at("B3") + B, "ws.Rows(3)" + B]),
    "row_clear_formats": ('ws.Range("A3").Value = 1\nws.Range("B3").Interior.Color = 255\nws.Rows(3).Font.Bold = True\n'
                          "ws.Rows(3).ClearFormats",
                          [at("A3") + B, at("Z3") + B, "ws.Rows(3)" + B, at("B3") + ".Interior.Color"]),
    "column_clear_formats": ('ws.Columns(3).Font.Bold = True\nws.Range("C2").Value = 1\nws.Columns(3).ClearFormats',
                             [at("C2") + B, "ws.Columns(3)" + B]),
    "blank_meets_row": ('ws.Rows(3).Font.Bold = True\nws.Range("C3").Font.Italic = True\nws.Rows(3).Font.Italic = True',
                        [at("C3") + B, at("C3") + IT]),
    "column_blank_clear_formats": ('ws.Columns(3).Font.Bold = True\nws.Range("C5").ClearFormats', [at("C5") + B]),
    "column_value_cleared": ('ws.Columns(3).Font.Bold = True\nws.Range("C5").Value = 1\nws.Range("C5").ClearContents',
                             [at("C5") + B]),
    "cross_cleared": ('ws.Rows(5).Font.Bold = True\nws.Columns(5).Font.Italic = True\nws.Range("E5").ClearContents',
                      [at("E5") + B, at("E5") + IT]),
    "cross_clear_formats": ('ws.Rows(5).Font.Bold = True\nws.Columns(5).Font.Italic = True\nws.Range("E5").ClearFormats',
                            [at("E5") + B, at("E5") + IT]),
    "cf_row_under_style_column": ("ws.Columns(3).Font.Bold = True\nws.Rows(3).Font.Italic = True\nws.Rows(3).ClearFormats",
                                  [at("C3") + B, at("C3") + IT]),
    "cf_plain_row_under_column": ("ws.Columns(3).Font.Bold = True\nws.Rows(3).ClearFormats", [at("C3") + B]),
    "cf_column_under_row": ("ws.Rows(3).Font.Bold = True\nws.Columns(3).ClearFormats", [at("C3") + B]),
    "cf_column_across_rows": ("ws.Rows(3).Font.Bold = True\nws.Columns(3).Font.Italic = True\nws.Columns(3).ClearFormats",
                              [at("C3") + B, at("C3") + IT]),
    "cf_cells_row_area": ('ws.Rows(3).Font.Bold = True\nws.Range("B3:D3").ClearFormats', [at("C3") + B]),
    "row_clear": ('ws.Rows(3).Font.Bold = True\nws.Range("B3").Value = 1\nws.Rows(3).Clear', [at("B3") + B]),
    "row_clear_contents": ('ws.Rows(3).Font.Bold = True\nws.Range("B3").Value = 1\nws.Rows(3).ClearContents',
                           [at("B3") + B, at("B3") + ".Value"]),
    # --- the whole sheet ---------------------------------------------------------------------------
    "whole_sheet": ('ws.Range("B2").Value = 1\nws.Cells.Font.Name = "Arial"',
                    [at("B2") + N, at("Z99") + N, "ws.Cells" + N, *heights(2)]),
    "whole_sheet_with_row": ('ws.Rows(3).Font.Bold = True\nws.Cells.Font.Name = "Arial"',
                             [at("Z3") + N, at("Z3") + B, *heights(3, 4)]),
    "whole_sheet_with_column": ("ws.Columns(3).Font.Bold = True\nws.Cells.Font.Italic = True",
                                [at("C3") + B, at("C3") + IT, at("D3") + IT]),
    "rows_all": ('ws.Rows("1:1048576").Font.Italic = True', [at("Z3") + IT]),
    "sheet_then_column_clear": ("ws.Cells.Font.Bold = True\nws.Columns(3).ClearFormats", [at("C4") + B, at("D4") + B]),
    "sheet_then_row": ("ws.Cells.Font.Bold = True\nws.Rows(4).Font.Italic = True", [at("C4") + B, at("C4") + IT]),
    "sheet_then_column": ("ws.Cells.Font.Bold = True\nws.Columns(4).Font.Italic = True", [at("D9") + B, at("D9") + IT]),
    "sheet_clear_formats": ("ws.Columns(3).Font.Italic = True\nws.Cells.Font.Bold = True\nws.Cells.ClearFormats",
                            [at("C3") + IT, at("C3") + B]),
    "sheet_twice": ('ws.Cells.Font.Bold = True\nws.Cells.Font.Italic = True\nws.Range("B2").Value = 1',
                    [at("B2") + B, at("B2") + IT]),
    "sheet_then_value_far": ('ws.Cells.Font.Bold = True\nws.Range("XFD1048576").Value = 1', [at("XFD1048576") + B]),
    "sheet_with_width": ("ws.Columns(3).ColumnWidth = 20\nws.Cells.Font.Bold = True", ["ws.Columns(3).ColumnWidth"]),
    "sheet_with_hidden": ("ws.Columns(3).Hidden = True\nws.Cells.Font.Bold = True", ["ws.Columns(3).Hidden"]),
    "columns_all_but_one": ('ws.Columns("A:XFC").Font.Bold = True', [at("XFD1") + B]),
    "base_all_but_last_then_row": ('ws.Columns("A:XFC").Font.Bold = True\nws.Rows(4).Font.Italic = True',
                                   [at("A4") + B, at("XFD4") + B, at("XFD4") + IT]),
    # --- rows and columns moving ---------------------------------------------------------------------
    "insert_below_formatted": ("ws.Rows(3).Font.Bold = True\nws.Rows(4).Insert",
                               [at("Z4") + B, at("Z3") + B, at("Z5") + B]),
    "insert_at_formatted": ("ws.Rows(3).Font.Bold = True\nws.Rows(3).Insert", [at("Z3") + B, at("Z4") + B]),
    "ins_at_bold_above_italic": ("ws.Rows(3).Font.Bold = True\nws.Rows(2).Font.Italic = True\nws.Rows(3).Insert",
                                 [at("Z3") + IT, at("Z3") + B]),
    "ins_far_below": ("ws.Rows(6).Font.Bold = True\nws.Rows(3).Insert", [at("Z7") + B]),
    "ins_two_at_bold": ('ws.Rows(3).Font.Bold = True\nws.Rows("3:4").Insert', [at("Z5") + B]),
    "ins_first_row": ("ws.Rows(1).Font.Bold = True\nws.Rows(1).Insert", [at("Z1") + B, at("Z2") + B]),
    "ins_below_bold_cells": ('ws.Rows(3).Font.Bold = True\nws.Range("B3").Interior.Color = 255\n'
                             'ws.Range("C3").Value = 5\nws.Rows(4).Insert', [at("B4") + ".Interior.Color", at("C4") + B]),
    "ins_below_cell_format": ('ws.Range("B3").Font.Italic = True\nws.Range("C3").Value = 5\nws.Rows(4).Insert',
                              [at("B4") + IT, at("C4") + ".Value"]),
    "ins_below_row_under_column": ('ws.Columns(2).Font.Italic = True\nws.Rows(3).Font.Bold = True\nws.Rows(4).Insert',
                                   [at("B4") + IT, at("B4") + B]),
    "ins_col_at_italic": ("ws.Columns(3).Font.Italic = True\nws.Columns(3).Insert", [at("C5") + IT, at("D5") + IT]),
    "ins_col_right_cells": ('ws.Range("C2").Font.Italic = True\nws.Columns(4).Insert', [at("D2") + IT]),
    "insert_column_right": ('ws.Columns("C").Font.Italic = True\nws.Columns("D").Insert', [at("D5") + IT, at("E5") + IT]),
    "delete_formatted_row": ("ws.Rows(3).Font.Bold = True\nws.Rows(5).Font.Italic = True\nws.Rows(3).Delete",
                             [at("Z3") + B, at("Z4") + IT, at("Z5") + IT]),
    "del_italic_below": ("ws.Rows(5).Font.Italic = True\nws.Rows(3).Delete", [at("Z4") + IT]),
    "del_bold_itself": ("ws.Rows(3).Font.Bold = True\nws.Rows(3).Delete", [at("Z3") + B]),
    "del_above_bold": ("ws.Rows(3).Font.Bold = True\nws.Rows(2).Delete", [at("Z2") + B]),
    "del_col_italic_right": ("ws.Columns(5).Font.Italic = True\nws.Columns(3).Delete", [at("D5") + IT]),
    "copy_row": ('ws.Rows(3).Font.Bold = True\nws.Range("B3").Value = 1\nws.Rows(3).Copy ws.Rows(8)',
                 [at("Z8") + B, at("B8") + B]),
    "copy_row_onto_row": ('ws.Rows(3).Font.Bold = True\nws.Range("B5").Value = 1\nws.Rows(3).Copy ws.Rows(5)',
                          [at("B5") + ".Value", at("Z5") + B]),
    "copy_plain_row_onto_bold": ("ws.Rows(5).Font.Bold = True\nws.Rows(3).Copy ws.Rows(5)", [at("Z5") + B]),
    "copy_column": ("ws.Columns(3).Font.Italic = True\nws.Columns(3).Copy ws.Columns(6)", [at("F9") + IT]),
    "copy_cells_from_bold_row": ('ws.Rows(3).Font.Bold = True\nws.Range("A3:B3").Copy ws.Range("D7")',
                                 [at("D7") + B, at("F7") + B]),
    # --- borders on whole rows, columns and the sheet ---------------------------------------------
    "b_row_top_italic_above": ("ws.Rows(2).Font.Italic = True\nws.Rows(3).Borders(xlEdgeTop).LineStyle = xlContinuous",
                               [at("Z2") + BO, at("Z3") + T]),
    "b_row_top_over_bottom": ("ws.Rows(2).Borders(xlEdgeBottom).LineStyle = xlContinuous\n"
                              "ws.Rows(3).Borders(xlEdgeTop).LineStyle = xlDash",
                              [at("Z2") + BO, at("Z3") + T]),
    "b_row_top_neighbour_cell": ('ws.Range("B2").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                                 "ws.Rows(3).Borders(xlEdgeTop).LineStyle = xlDash", [at("B2") + BO, at("B3") + T]),
    "b_row_bottom": ("ws.Rows(3).Borders(xlEdgeBottom).LineStyle = xlContinuous", [at("Z4") + T]),
    "b_row_left": ("ws.Rows(3).Borders(xlEdgeLeft).LineStyle = xlContinuous", [at("A3") + L, at("Z3") + L]),
    "b_row_right": ("ws.Rows(3).Borders(xlEdgeRight).LineStyle = xlContinuous", [at("XFD3") + R, at("Z3") + R]),
    "b_rows_left_two": ('ws.Rows("3:4").Borders(xlEdgeLeft).LineStyle = xlContinuous', [at("A3") + L, at("A4") + L]),
    "b_row_left_cells": ('ws.Range("B3").Value = 1\nws.Range("A3").Value = 2\n'
                         "ws.Rows(3).Borders(xlEdgeLeft).LineStyle = xlContinuous", [at("A3") + L]),
    "b_rows_inside_h": ('ws.Rows("3:4").Borders(xlInsideHorizontal).LineStyle = xlContinuous',
                        [at("Z3") + BO, at("Z4") + T, at("Z4") + BO]),
    "b_row_inside_v": ("ws.Rows(3).Borders(xlInsideVertical).LineStyle = xlContinuous",
                       [at("Z3") + L, at("Z3") + R, at("A3") + L]),
    "b_row_all": ("ws.Rows(3).Borders.LineStyle = xlContinuous",
                  [at("Z3") + L, at("Z3") + T, at("Z3") + BO, "ws.Rows(3).Borders.LineStyle"]),
    "b_row_around": ("ws.Rows(3).BorderAround xlContinuous", [at("Z3") + T, at("Z3") + L, at("A3") + L]),
    "b_row_diag": ("ws.Rows(3).Borders(xlDiagonalUp).LineStyle = xlContinuous", [at("Z3") + UP]),
    "b_row_top_first": ("ws.Rows(1).Borders(xlEdgeTop).LineStyle = xlContinuous", [at("Z1") + T]),
    "b_row_bottom_last": ("ws.Rows(1048576).Borders(xlEdgeBottom).LineStyle = xlContinuous", [at("Z1048576") + BO]),
    "b_row_top_cells": ('ws.Range("A3").Value = 1\nws.Range("C2").Value = 2\n'
                        "ws.Rows(3).Borders(xlEdgeTop).LineStyle = xlContinuous", [at("C2") + BO, at("A3") + T]),
    "b_row_top_under_column": ("ws.Columns(3).Font.Italic = True\nws.Rows(3).Borders(xlEdgeTop).LineStyle = xlContinuous",
                               [at("C3") + T, at("C3") + IT]),
    "b_row_read": ("ws.Rows(3).Borders(xlEdgeTop).LineStyle = xlContinuous",
                   ["ws.Rows(3)" + T, "ws.Rows(3).Borders.LineStyle", "ws.Rows(3)" + BO]),
    "b_col_left": ("ws.Columns(3).Borders(xlEdgeLeft).LineStyle = xlContinuous", [at("B9") + R]),
    "b_col_right": ("ws.Columns(3).Borders(xlEdgeRight).LineStyle = xlContinuous", [at("D9") + L, at("C9") + R]),
    "b_col_top": ("ws.Columns(3).Borders(xlEdgeTop).LineStyle = xlContinuous", [at("C1") + T, at("C9") + T]),
    "b_col_bottom": ("ws.Columns(3).Borders(xlEdgeBottom).LineStyle = xlContinuous",
                     [at("C1048576") + BO, at("C9") + BO]),
    "b_col_top_cells": ('ws.Range("C1").Value = 1\nws.Columns(3).Borders(xlEdgeTop).LineStyle = xlContinuous',
                        [at("C1") + T, at("C9") + T]),
    "b_cols_inside_v": ('ws.Columns("C:D").Borders(xlInsideVertical).LineStyle = xlContinuous',
                        [at("C9") + R, at("D9") + L, at("D9") + R]),
    "b_col_inside_h": ("ws.Columns(3).Borders(xlInsideHorizontal).LineStyle = xlContinuous", [at("C9") + T, at("C9") + BO]),
    "b_col_all": ("ws.Columns(3).Borders.LineStyle = xlContinuous",
                  [at("C9") + L, at("C9") + T, at("C9") + R, at("C9") + BO, "ws.Columns(3).Borders.LineStyle"]),
    "b_col_around": ("ws.Columns(3).BorderAround xlContinuous", [at("C9") + L, at("C9") + T, at("C1") + T]),
    "b_sheet_top": ("ws.Cells.Borders(xlEdgeTop).LineStyle = xlContinuous", [at("Z1") + T, at("C9") + T]),
    "b_sheet_left": ("ws.Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous", [at("A9") + L, at("C9") + L]),
    "b_sheet_inside_h": ("ws.Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous", [at("C9") + T, at("C9") + BO]),
    # --- how tall rows are in the fonts rows and columns carry --------------------------------------
    "row_font_size": ("ws.Rows(2).Font.Size = 20", [at("Q2") + S, *heights(2)]),
    "row_small_font": ("ws.Rows(2).Font.Size = 8", heights(2)),
    "row_arial": ('ws.Rows(2).Font.Name = "Arial"', [*heights(2), "ws.Rows(2).UseStandardHeight"]),
    "row_arial_normal_cell": ('ws.Rows(2).Font.Name = "Arial"\nws.Range("B2").Font.Name = "Aptos Narrow"', heights(2)),
    "row_arial_value": ('ws.Rows(2).Font.Name = "Arial"\nws.Range("B2").Value = 5', heights(2)),
    "column_font_size": ('ws.Columns("B").Font.Size = 20\nws.Range("B7").Value = "x"', [at("B1") + S, *heights(1, 7)]),
    "column_big_cell_small": ('ws.Columns("B").Font.Size = 20\nws.Range("B7").Font.Size = 8', heights(1, 7)),
    "column_arial": ('ws.Columns(2).Font.Name = "Arial"', heights(2)),
    "row_arial_column_arial": ('ws.Rows(2).Font.Name = "Arial"\nws.Columns(3).Font.Name = "Arial"', heights(2, 3)),
    "sheet_size": ('ws.Cells.Font.Size = 20\nws.Range("C3").Font.Size = 11', heights(1, 3)),
    "row_style_rows_all_height": ('ws.Rows("1:1048576").RowHeight = 20\nws.Rows(3).Font.Bold = True', heights(3)),
    "unmeasured_row_column_mix": ('ws.Rows(2).Font.Name = "Arial"\nws.Columns(3).Font.Name = "Times New Roman"',
                                  heights(2)),
}

#: Planted sheets: name -> (cols, sheetData body, a macro run once the file is open, reads).
#: {B} and {I} stand for a bold and an italic xf the base workbook holds.
PLANTED: dict[str, tuple[str, str, str, list[str]]] = {
    "load_all_cols": ('<col min="1" max="16384" width="9.140625" style="{B}"/>', "", "", [at("Z9") + B]),
    "load_split": ('<col min="1" max="2" width="9.140625" style="{B}"/><col min="3" max="3" width="9.140625" style="{I}"/>'
                   '<col min="4" max="16384" width="9.140625" style="{B}"/>', "", "", [at("C5") + IT, at("D5") + B]),
    "load_all_but_last": ('<col min="1" max="16383" width="9.140625" style="{B}"/>'
                          '<col min="16384" max="16384" width="9.140625" style="{I}"/>', "", "", [at("XFD5") + IT]),
    "load_half": ('<col min="1" max="8192" width="9.140625" style="{B}"/>'
                  '<col min="8193" max="16384" width="9.140625" style="{I}"/>', "", "", [at("LCC5") + IT]),
    "load_xfd_only": ('<col min="16384" max="16384" width="9.140625" style="{B}"/>', "", "", [at("A5") + B]),
    "load_all_then_row": ('<col min="1" max="16384" width="9.140625" style="{B}"/>', "",
                          "ws.Rows(4).Font.Italic = True", [at("C4") + B, at("C4") + IT]),
    "load_split_then_row": ('<col min="1" max="2" width="9.140625" style="{B}"/><col min="3" max="3" width="9.140625" '
                            'style="{I}"/><col min="4" max="16384" width="9.140625" style="{B}"/>', "",
                            "ws.Rows(4).Font.Underline = xlUnderlineStyleSingle", [at("C4") + IT, at("D4") + B]),
    "load_all_then_column_clear": ('<col min="1" max="16384" width="9.140625" style="{B}"/>', "",
                                   "ws.Columns(3).ClearFormats", [at("C4") + B]),
    "load_one_col_no_width": ('<col min="3" max="3" style="{B}"/>', "", "",
                              ["ws.Columns(3).ColumnWidth", "ws.Columns(3).Hidden", at("C4") + B]),
    "load_row_cf_no_s": ('<col min="3" max="3" width="9.140625" style="{B}"/>', '<row r="3" customFormat="1"/>', "",
                         [at("C3") + B, at("C4") + B]),
    "load_row_s_no_cf": ("", '<row r="3" s="{B}"/>', "", [at("Z3") + B]),
    "load_row_blank_cell": ("", '<row r="3" s="{B}" customFormat="1"><c r="C3"/></row>', "", [at("C3") + B, at("D3") + B]),
    "load_blank_same_as_row": ("", '<row r="3" s="{B}" customFormat="1"><c r="C3" s="{B}"/></row>', "", [at("C3") + B]),
    "load_blank_default_cell": ("", '<row r="3"><c r="C3"/></row>', "", [at("C3") + B]),
    "load_blank_same_as_column": ('<col min="3" max="3" width="9.140625" style="{B}"/>',
                                  '<row r="5"><c r="C5" s="{B}"/></row>', "", [at("C5") + B]),
    "load_row_style_cells_same": ("", '<row r="3" s="{B}" customFormat="1"><c r="C3" s="{B}"><v>1</v></c></row>',
                                  'ws.Range("C3").ClearContents', [at("C3") + B]),
}


def reads_function(name: str, reads: list[str]) -> list[str]:
    """Reads the used range first, then every read, into one line."""
    lines = [f"Private Function {name}(ws As Object) As String", "Dim out As String, v As Variant",
             "On Error Resume Next"]
    for expression in ["ws.UsedRange.Address", *reads]:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return [*lines, "On Error GoTo 0", f"{name} = out", "End Function"]


def case_code(index: int, setup: str, reads: list[str]) -> str:
    """Case<n>(ws) sets the case up on ws and returns Reads<n>(ws); a setup error is recorded as S<n>."""
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String", "On Error Resume Next",
             "Err.Clear", *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f"Case{index} = failed & Reads{index}(ws)", "End Function",
             *reads_function(f"Reads{index}", reads)]
    return "\n".join(lines) + "\n"


def sheet_parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        names = sorted((n for n in package.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
                       key=lambda n: int(re.search(r"(\d+)\.xml", n).group(1)))  # type: ignore[union-attr]
        return [package.read(name).decode("utf-8") for name in names]


def run(excel: ExcelSession, code: str, procedure: str, timeout: float = 900.0) -> str:
    result = excel.run_vba(code, procedure, timeout=timeout)
    assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    return str(result.value)


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    authored, base = folder / "row_formats.xlsx", folder / "planted_base.xlsx"
    planted, resaved = folder / "planted_source.xlsx", folder / "planted.xlsx"
    names = list(CASES)
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, (setup, reads)) in enumerate(CASES.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup, reads))
    build += [f'wb.SaveAs Filename:="{authored}", FileFormat:=51', "wb.Close False", "Build = out", "End Function"]
    reopen = [HELPER, "Public Function Reopen() As String", "Dim wb As Object, out As String",
              "Application.DisplayAlerts = False", f'Set wb = Workbooks.Open("{authored}")']
    reopen += [f'out = out & Reads{index}(wb.Worksheets({index + 1})) & "|"' for index in range(len(CASES))]
    reopen += ["wb.Close False", "Reopen = out", "End Function"]
    reopen += [line for index, (_, reads) in enumerate(CASES.values()) for line in reads_function(f"Reads{index}", reads)]
    make_base = ["Public Function MakeBase() As String", "Dim wb As Object, i As Long",
                 "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
                 'wb.Worksheets(1).Name = "styles"', 'wb.Worksheets(1).Range("A1").Font.Bold = True',
                 'wb.Worksheets(1).Range("A2").Font.Italic = True', f"For i = 1 To {len(PLANTED)}",
                 "wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)", "Next i",
                 f'wb.SaveAs Filename:="{base}", FileFormat:=51', "wb.Close False", 'MakeBase = "ok"', "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        standard = run(excel, "Public Function S() As String\nS = Workbooks.Add(xlWBATWorksheet).Worksheets(1)"
                              ".StandardHeight\nActiveWorkbook.Close False\nEnd Function\n", "S", 60.0)
        assert standard == "15", f"needs a 96-DPI display (StandardHeight {standard})"
        before = run(excel, "\n".join(build) + "\n" + "".join(bodies), "Build").split("|")[: len(CASES)]
        after = run(excel, "\n".join(reopen) + "\n", "Reopen").split("|")[: len(CASES)]
        print("built and reopened", len(CASES), "cases", flush=True)
        run(excel, "\n".join(make_base) + "\n", "MakeBase", 120.0)
        with zipfile.ZipFile(base) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        styles_sheet = sheet_parts(base)[0]
        bold = re.search(r'<c r="A1" s="(\d+)"', styles_sheet).group(1)  # type: ignore[union-attr]
        italic = re.search(r'<c r="A2" s="(\d+)"', styles_sheet).group(1)  # type: ignore[union-attr]
        sheet_names = sorted((n for n in parts if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
                             key=lambda n: int(re.search(r"(\d+)\.xml", n).group(1)))  # type: ignore[union-attr]
        for name, (cols, rows, _, _) in zip(sheet_names[1:], PLANTED.values(), strict=True):
            xml = parts[name].decode("utf-8")
            cols_xml = f"<cols>{cols.format(B=bold, I=italic)}</cols>" if cols else ""
            rows_xml = f"<sheetData>{rows.format(B=bold, I=italic)}</sheetData>" if rows else "<sheetData/>"
            planted_xml = cols_xml + rows_xml
            xml = re.sub(r"<sheetData/>|<sheetData>.*?</sheetData>", lambda _, put=planted_xml: put, xml, count=1,
                         flags=re.DOTALL)
            parts[name] = xml.encode("utf-8")
        with zipfile.ZipFile(planted, "w", zipfile.ZIP_DEFLATED) as package:
            for name, data in parts.items():
                package.writestr(name, data)
        loaded_code = [HELPER, "Public Function Loaded() As String", "Dim wb As Object, ws As Object, out As String",
                       "Application.DisplayAlerts = False", f'Set wb = Workbooks.Open("{planted}")']
        functions: list[str] = []
        for index, (_, (_, _, macro, reads)) in enumerate(PLANTED.items()):
            loaded_code += [f"Set ws = wb.Worksheets({index + 2})", "On Error Resume Next", *macro.splitlines(),
                            "On Error GoTo 0", f'out = out & Planted{index}(ws) & "|"']
            functions += reads_function(f"Planted{index}", reads)
        loaded_code += [f'wb.SaveAs Filename:="{resaved}", FileFormat:=51', "wb.Close False", "Loaded = out",
                        "End Function"]
        loaded = run(excel, "\n".join(loaded_code + functions) + "\n", "Loaded").split("|")[: len(PLANTED)]
    OUT.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(authored, OUT / "row_formats.xlsx")
    shutil.copyfile(planted, OUT / "planted_source.xlsx")
    shutil.copyfile(resaved, OUT / "planted.xlsx")
    record = {
        "helper": HELPER,
        "cases": [{"name": name, "setup": setup, "reads": reads, "before": first, "after": second}
                  for name, (setup, reads), first, second in zip(names, CASES.values(), before, after, strict=True)],
        "planted": [{"name": name, "cols": cols, "rows": rows, "macro": macro, "reads": reads, "answers": answer}
                    for (name, (cols, rows, macro, reads)), answer in zip(PLANTED.items(), loaded, strict=True)],
        "bold": int(bold), "italic": int(italic),
    }
    (OUT / "row_formats.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, first, second in zip(names, before, after, strict=True):
        print(f"{name}: {first}" + ("" if first == second else f"  REOPENED {second}"))
    for name, answer in zip(PLANTED, loaded, strict=True):
        print(f"{name}: {answer}")


if __name__ == "__main__":
    main()
