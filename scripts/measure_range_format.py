"""Cell formatting as Excel's object model reports it: fonts, fills, borders, alignment, protection.

Every probe starts from cleared cells on the same sheet, sets something
and reads a list of properties back. Each read records the value's type
as well as its text, reads a mixed range as Null, and records an error
as E<number>, so the replay in tests/test_excel_format.py holds the
model to all three.

    python scripts/measure_range_format.py

writes tests/fixtures/range_format.json.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent

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

FONT = ("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color", "ColorIndex",
        "ThemeColor", "TintAndShade", "Superscript", "Subscript", "FontStyle")
FILL = ("Color", "ColorIndex", "Pattern", "PatternColor", "PatternColorIndex", "ThemeColor",
        "TintAndShade", "PatternThemeColor", "PatternTintAndShade")
ALIGN = ("HorizontalAlignment", "VerticalAlignment", "WrapText", "IndentLevel", "Orientation",
         "ShrinkToFit", "AddIndent", "ReadingOrder")
PROTECT = ("Locked", "FormulaHidden")
BORDER = ("LineStyle", "Weight", "Color", "ColorIndex", "ThemeColor", "TintAndShade")
EDGES = {"left": 7, "top": 8, "bottom": 9, "right": 10}


def font(target: str) -> list[str]:
    return [f"{target}.Font.{name}" for name in FONT]


def fill(target: str) -> list[str]:
    return [f"{target}.Interior.{name}" for name in FILL]


def align(target: str) -> list[str]:
    return [f"{target}.{name}" for name in ALIGN]


def protect(target: str) -> list[str]:
    return [f"{target}.{name}" for name in PROTECT]


def border(target: str, index: int, names: tuple[str, ...] = BORDER) -> list[str]:
    return [f"{target}.Borders({index}).{name}" for name in names]


def styles(target: str) -> list[str]:
    return [f"{target}.Borders({index}).LineStyle" for index in (5, 6, 7, 8, 9, 10, 11, 12)]


A1 = 'Range("A1")'
AB = 'Range("A1:B1")'
BC = 'Range("B2:C3")'

PROBES: dict[str, tuple[str, list[str]]] = {
    # --- a cell nobody has touched -------------------------------------------------
    "default_font": ("", font(A1)),
    "default_fill": ("", fill(A1)),
    "default_alignment": ("", [*align(A1), *protect(A1), f"{A1}.NumberFormat"]),
    "default_borders": ("", [*[one for index in (5, 6, 7, 8, 9, 10) for one in border(A1, index)],
                             f"{A1}.Borders.LineStyle", f"{A1}.Borders.Weight", f"{A1}.Borders.Count"]),
    "default_inside": ("", [*border(BC, 11), *border(BC, 12)]),
    # --- fonts -----------------------------------------------------------------------
    "font_name": (f'{A1}.Font.Name = "Arial"', font(A1)),
    "font_size": (f"{A1}.Font.Size = 14", font(A1)),
    "font_size_half": (f"{A1}.Font.Size = 10.5", [f"{A1}.Font.Size"]),
    "font_size_odd": (f"{A1}.Font.Size = 10.3", [f"{A1}.Font.Size"]),
    "font_size_zero": (f"{A1}.Font.Size = 0", [f"{A1}.Font.Size"]),
    "font_bold": (f"{A1}.Font.Bold = True", font(A1)),
    "font_italic": (f"{A1}.Font.Italic = True", font(A1)),
    "font_bold_italic": (f"{A1}.Font.Bold = True\n{A1}.Font.Italic = True", font(A1)),
    "font_style_text": (f'{A1}.Font.FontStyle = "Bold Italic"', font(A1)),
    "font_style_regular": (f'{A1}.Font.Bold = True\n{A1}.Font.FontStyle = "Regular"', font(A1)),
    "font_underline_single": (f"{A1}.Font.Underline = xlUnderlineStyleSingle", font(A1)),
    "font_underline_double": (f"{A1}.Font.Underline = xlUnderlineStyleDouble", [f"{A1}.Font.Underline"]),
    "font_underline_accounting": (f"{A1}.Font.Underline = xlUnderlineStyleSingleAccounting",
                                  [f"{A1}.Font.Underline"]),
    "font_underline_double_accounting": (f"{A1}.Font.Underline = xlUnderlineStyleDoubleAccounting",
                                         [f"{A1}.Font.Underline"]),
    "font_underline_true": (f"{A1}.Font.Underline = True", [f"{A1}.Font.Underline"]),
    "font_underline_off": (f"{A1}.Font.Underline = xlUnderlineStyleSingle\n{A1}.Font.Underline = xlUnderlineStyleNone",
                           [f"{A1}.Font.Underline"]),
    "font_strike": (f"{A1}.Font.Strikethrough = True", font(A1)),
    "font_superscript": (f"{A1}.Font.Superscript = True", font(A1)),
    "font_subscript_after_superscript": (f"{A1}.Font.Superscript = True\n{A1}.Font.Subscript = True", font(A1)),
    "font_color": (f"{A1}.Font.Color = RGB(255, 0, 0)", font(A1)),
    "font_color_unlisted": (f"{A1}.Font.Color = RGB(18, 52, 86)", font(A1)),
    "font_color_index": (f"{A1}.Font.ColorIndex = 5", font(A1)),
    "font_color_index_automatic": (f"{A1}.Font.ColorIndex = 5\n{A1}.Font.ColorIndex = xlColorIndexAutomatic",
                                   font(A1)),
    "font_color_index_none": (f"{A1}.Font.ColorIndex = xlColorIndexNone", font(A1)),
    "font_color_index_bad": (f"{A1}.Font.ColorIndex = 57", font(A1)),
    "font_theme": (f"{A1}.Font.ThemeColor = xlThemeColorAccent1", font(A1)),
    "font_theme_dark1": (f"{A1}.Font.ThemeColor = xlThemeColorDark1", font(A1)),
    "font_theme_light1": (f"{A1}.Font.ThemeColor = xlThemeColorLight1", font(A1)),
    "font_theme_tint": (f"{A1}.Font.ThemeColor = xlThemeColorAccent1\n{A1}.Font.TintAndShade = 0.4", font(A1)),
    "font_theme_shade": (f"{A1}.Font.ThemeColor = xlThemeColorAccent2\n{A1}.Font.TintAndShade = -0.25", font(A1)),
    "font_tint_on_rgb": (f"{A1}.Font.Color = RGB(200, 100, 50)\n{A1}.Font.TintAndShade = 0.5", font(A1)),
    "font_theme_bad": (f"{A1}.Font.ThemeColor = 13", font(A1)),
    "font_mixed_bold": (f"{A1}.Font.Bold = True", font(AB)),
    "font_mixed_name": (f'{A1}.Font.Name = "Arial"\n{A1}.Font.Size = 14', font(AB)),
    "font_range_set": (f'{AB}.Font.Name = "Arial"\n{AB}.Font.Color = RGB(0, 128, 0)', font(AB)),
    # --- fills -----------------------------------------------------------------------
    "fill_color": (f"{A1}.Interior.Color = RGB(255, 255, 0)", fill(A1)),
    "fill_color_unlisted": (f"{A1}.Interior.Color = RGB(18, 52, 86)", fill(A1)),
    "fill_color_index": (f"{A1}.Interior.ColorIndex = 6", fill(A1)),
    "fill_color_index_none": (f"{A1}.Interior.Color = RGB(255, 255, 0)\n{A1}.Interior.ColorIndex = xlColorIndexNone",
                              fill(A1)),
    "fill_color_index_automatic": (f"{A1}.Interior.ColorIndex = xlColorIndexAutomatic", fill(A1)),
    "fill_pattern": (f"{A1}.Interior.Pattern = xlGray50", fill(A1)),
    "fill_pattern_color": (f"{A1}.Interior.Pattern = xlChecker\n{A1}.Interior.PatternColor = RGB(0, 0, 255)", fill(A1)),
    "fill_pattern_none": (f"{A1}.Interior.Color = RGB(255, 255, 0)\n{A1}.Interior.Pattern = xlNone", fill(A1)),
    "fill_theme": (f"{A1}.Interior.ThemeColor = xlThemeColorAccent2", fill(A1)),
    "fill_theme_tint": (f"{A1}.Interior.ThemeColor = xlThemeColorAccent2\n{A1}.Interior.TintAndShade = 0.6", fill(A1)),
    "fill_mixed": (f"{A1}.Interior.Color = RGB(255, 255, 0)", fill(AB)),
    # --- alignment and protection ----------------------------------------------------
    "align_center": (f"{A1}.HorizontalAlignment = xlCenter", align(A1)),
    "align_vertical_top": (f"{A1}.VerticalAlignment = xlTop", align(A1)),
    "align_wrap": (f"{A1}.WrapText = True", align(A1)),
    "align_indent": (f"{A1}.IndentLevel = 2", align(A1)),
    "align_indent_right": (f"{A1}.HorizontalAlignment = xlRight\n{A1}.IndentLevel = 3", align(A1)),
    "align_indent_center": (f"{A1}.HorizontalAlignment = xlCenter\n{A1}.IndentLevel = 1", align(A1)),
    "align_indent_bad": (f"{A1}.IndentLevel = 251", align(A1)),
    "align_orientation": (f"{A1}.Orientation = 45", align(A1)),
    "align_orientation_upward": (f"{A1}.Orientation = xlUpward", align(A1)),
    "align_orientation_vertical": (f"{A1}.Orientation = xlVertical", align(A1)),
    "align_orientation_bad": (f"{A1}.Orientation = 91", align(A1)),
    "align_shrink": (f"{A1}.ShrinkToFit = True", align(A1)),
    "align_shrink_and_wrap": (f"{A1}.ShrinkToFit = True\n{A1}.WrapText = True", align(A1)),
    "align_bad": (f"{A1}.HorizontalAlignment = 99", align(A1)),
    "align_mixed": (f"{A1}.HorizontalAlignment = xlCenter\n{A1}.WrapText = True", align(AB)),
    "protect_unlocked": (f"{A1}.Locked = False\n{A1}.FormulaHidden = True", protect(A1)),
    "protect_mixed": (f"{A1}.Locked = False", protect(AB)),
    # --- borders ---------------------------------------------------------------------
    "border_left": (f"{A1}.Borders(xlEdgeLeft).LineStyle = xlContinuous", border(A1, 7)),
    "border_weight_only": (f"{A1}.Borders(xlEdgeLeft).Weight = xlMedium", border(A1, 7)),
    "border_color_only": (f"{A1}.Borders(xlEdgeLeft).Color = RGB(255, 0, 0)", border(A1, 7)),
    "border_color_index": (f"{A1}.Borders(xlEdgeLeft).ColorIndex = 3", border(A1, 7)),
    "border_none": (f"{A1}.Borders(xlEdgeLeft).LineStyle = xlContinuous\n{A1}.Borders(xlEdgeLeft).LineStyle = xlNone",
                    border(A1, 7)),
    "border_range_top": (f"{BC}.Borders(xlEdgeTop).LineStyle = xlContinuous",
                         ['Range("B2").Borders(8).LineStyle', 'Range("C2").Borders(8).LineStyle',
                          'Range("B3").Borders(8).LineStyle', 'Range("B1").Borders(9).LineStyle',
                          f"{BC}.Borders(8).LineStyle", f"{BC}.Borders(9).LineStyle", f"{BC}.Borders(12).LineStyle",
                          f"{BC}.Borders.LineStyle"]),
    "border_shared_edge": ('Range("B2").Borders(xlEdgeBottom).LineStyle = xlContinuous',
                           ['Range("B2").Borders(9).LineStyle', 'Range("B3").Borders(8).LineStyle',
                            'Range("B2:B3").Borders(12).LineStyle']),
    "border_shared_right": ('Range("B2").Borders(xlEdgeRight).LineStyle = xlDouble',
                            ['Range("B2").Borders(10).LineStyle', 'Range("C2").Borders(7).LineStyle',
                             'Range("C2").Borders(7).Weight']),
    "border_inside": (f"{BC}.Borders(xlInsideHorizontal).LineStyle = xlContinuous",
                      ['Range("B2").Borders(9).LineStyle', 'Range("B3").Borders(8).LineStyle',
                       'Range("C2").Borders(9).LineStyle', 'Range("B3").Borders(9).LineStyle',
                       f"{BC}.Borders(12).LineStyle", f"{BC}.Borders(11).LineStyle", f"{BC}.Borders(8).LineStyle"]),
    "border_inside_single_cell": (f"{A1}.Borders(xlInsideHorizontal).LineStyle = xlContinuous",
                                  [*styles(A1)]),
    "border_around": (f"{BC}.BorderAround xlContinuous, xlMedium",
                      ['Range("B2").Borders(7).Weight', 'Range("B2").Borders(8).Weight',
                       'Range("C3").Borders(10).Weight', 'Range("C3").Borders(9).Weight',
                       'Range("B2").Borders(9).LineStyle', f"{BC}.Borders(12).LineStyle",
                       f"{BC}.Borders(8).LineStyle", f"{BC}.Borders.LineStyle"]),
    "border_around_color": (f"{BC}.BorderAround LineStyle:=xlDash, Weight:=xlThin, Color:=RGB(0, 0, 255)",
                            [*border(BC, 8), *border(BC, 7)]),
    "border_all": (f"{BC}.Borders.LineStyle = xlContinuous",
                   [*[f"{BC}.Borders({index}).LineStyle" for index in (5, 6, 7, 8, 9, 10, 11, 12)],
                    f"{BC}.Borders.LineStyle", f"{BC}.Borders.Weight"]),
    "border_all_weight": (f"{BC}.Borders.Weight = xlThick",
                          [*[f"{BC}.Borders({index}).LineStyle" for index in (7, 8, 11, 12)],
                           *[f"{BC}.Borders({index}).Weight" for index in (7, 8, 11, 12)]]),
    "border_diagonal": (f"{A1}.Borders(xlDiagonalUp).LineStyle = xlContinuous", styles(A1)),
    "border_diagonal_both": (f"{A1}.Borders(xlDiagonalUp).LineStyle = xlDot\n"
                             f"{A1}.Borders(xlDiagonalDown).LineStyle = xlContinuous",
                             [*border(A1, 5, ("LineStyle", "Weight")), *border(A1, 6, ("LineStyle", "Weight"))]),
    "border_bad_index": (f"{A1}.Borders(4).LineStyle = xlContinuous", styles(A1)),
    "border_legacy_left": (f"{A1}.Borders(1).LineStyle = xlContinuous", styles(A1)),
    "border_legacy_right": (f"{A1}.Borders(2).LineStyle = xlContinuous", styles(A1)),
    "border_legacy_top": (f"{A1}.Borders(3).LineStyle = xlContinuous", styles(A1)),
    "border_index_zero": (f"{A1}.Borders(0).LineStyle = xlContinuous", styles(A1)),
    "border_index_thirteen": (f"{A1}.Borders(13).LineStyle = xlContinuous", styles(A1)),
    "border_value": (f"{A1}.Borders(xlEdgeTop).Value = xlContinuous", border(A1, 8, ("LineStyle", "Weight"))),
    "borders_value": (f"{A1}.Borders.Value = xlContinuous", [*styles(A1), f"{A1}.Borders.Value"]),
    "border_mixed_style": ('Range("B2").Borders(xlEdgeTop).LineStyle = xlContinuous\n'
                           'Range("C2").Borders(xlEdgeTop).LineStyle = xlDash',
                           [*border('Range("B2:C2")', 8), 'Range("B2:C2").Borders.LineStyle']),
    "border_first_cell": ('Range("B2").Borders(xlEdgeTop).LineStyle = xlDash\n'
                          'Range("C2").Borders(xlEdgeTop).LineStyle = xlContinuous',
                          [*border('Range("B2:C2")', 8), 'Range("B2:C2").Borders.LineStyle']),
    "border_last_cell_only": ('Range("C2").Borders(xlEdgeTop).LineStyle = xlContinuous',
                              [*border('Range("B2:C2")', 8), 'Range("B2:C2").Borders.LineStyle']),
    "border_left_edge_column": ('Range("B3").Borders(xlEdgeLeft).LineStyle = xlDouble',
                                [*border('Range("B2:B3")', 7, ("LineStyle", "Weight"))]),
    "border_inside_partial": ('Range("C2").Borders(xlEdgeBottom).LineStyle = xlContinuous',
                              [*border('Range("B2:C3")', 12, ("LineStyle", "Weight"))]),
    "borders_count_range": ("", ['Range("B2:C3").Borders.Count', 'Range("B2:B3").Borders.Count',
                                 'Range("A:A").Borders.Count']),
    "border_mixed_color": ('Range("B2:C2").Borders(xlEdgeTop).LineStyle = xlContinuous\n'
                           'Range("C2").Borders(xlEdgeTop).Color = RGB(255, 0, 0)',
                           [*border('Range("B2:C2")', 8), 'Range("B2:C2").Borders.Color']),
    "border_theme": (f"{A1}.Borders(xlEdgeLeft).ThemeColor = xlThemeColorAccent3", border(A1, 7)),
    "border_theme_tint": (f"{A1}.Borders(xlEdgeLeft).ThemeColor = xlThemeColorAccent3\n"
                          f"{A1}.Borders(xlEdgeLeft).TintAndShade = -0.5", border(A1, 7)),
    "font_size_bounds": (f"{A1}.Font.Size = 1\nReport = Report & {A1}.Font.Size & \";\"\n"
                         f"{A1}.Font.Size = 409\nReport = Report & {A1}.Font.Size & \";\"\n"
                         f"{A1}.Font.Size = 409.5\nReport = Report & Err.Number & \";\"\nErr.Clear\n"
                         f"{A1}.Font.Size = 0.5\nReport = Report & Err.Number & \";\"\nErr.Clear",
                         [f"{A1}.Font.Size"]),
    "font_name_empty": (f'{A1}.Font.Name = ""', [f"{A1}.Font.Name"]),
    "font_style_then_italic": (f'{A1}.Font.FontStyle = "Bold"\n{A1}.Font.Italic = True', font(A1)),
    "font_style_unknown": (f'{A1}.Font.FontStyle = "Heavy"', font(A1)),
    "font_mixed_color": (f"{A1}.Font.Color = RGB(255, 0, 0)", font(AB)),
    "font_mixed_theme": (f"{A1}.Font.ThemeColor = xlThemeColorAccent1", font(AB)),
    "font_mixed_underline": (f"{A1}.Font.Underline = xlUnderlineStyleDouble\n"
                             f'Range("B1").Font.Underline = xlUnderlineStyleSingle', font(AB)),
    "fill_mixed_two_colors": (f'{A1}.Interior.Color = RGB(255, 255, 0)\nRange("B1").Interior.Color = RGB(255, 0, 0)',
                              fill(AB)),
    "fill_mixed_same": (f'{AB}.Interior.Color = RGB(255, 255, 0)', fill(AB)),
    "fill_mixed_pattern": (f"{A1}.Interior.Pattern = xlGray50", fill(AB)),
    "fill_mixed_theme": (f"{A1}.Interior.ThemeColor = xlThemeColorAccent1\n"
                         f'Range("B1").Interior.ThemeColor = xlThemeColorAccent2', fill(AB)),
    "fill_solid_pattern": (f"{A1}.Interior.Pattern = xlSolid", fill(A1)),
    "fill_tint_on_rgb": (f"{A1}.Interior.Color = RGB(200, 100, 50)\n{A1}.Interior.TintAndShade = -0.5", fill(A1)),
    "fill_pattern_color_index_none": (f"{A1}.Interior.Pattern = xlChecker\n{A1}.Interior.PatternColorIndex = xlNone",
                                      fill(A1)),
    "fill_color_on_pattern": (f"{A1}.Interior.Pattern = xlChecker\n{A1}.Interior.Color = RGB(0, 255, 0)", fill(A1)),
    "fill_pattern_theme": (f"{A1}.Interior.Pattern = xlChecker\n{A1}.Interior.PatternThemeColor = xlThemeColorAccent4\n"
                           f"{A1}.Interior.PatternTintAndShade = 0.25", fill(A1)),
    "align_orientation_right_angles": (f"{A1}.Orientation = 90\nReport = Report & {A1}.Orientation & \";\"\n"
                                       f"{A1}.Orientation = -90\nReport = Report & {A1}.Orientation & \";\"\n"
                                       f"{A1}.Orientation = xlDownward\nReport = Report & {A1}.Orientation & \";\"\n"
                                       f"{A1}.Orientation = 0", [f"{A1}.Orientation"]),
    "align_horizontal_all": ("For Each k In Array(1, -4131, -4108, -4152, 5, -4130, 7, -4117)\n"
                             f"{A1}.HorizontalAlignment = k\nReport = Report & {A1}.HorizontalAlignment & \",\"\n"
                             "Next k", []),
    "align_vertical_all": ("For Each k In Array(-4160, -4108, -4107, -4130, -4117)\n"
                           f"{A1}.VerticalAlignment = k\nReport = Report & {A1}.VerticalAlignment & \",\"\n"
                           "Next k", []),
    "align_vertical_bad": (f"{A1}.VerticalAlignment = xlLeft", [f"{A1}.VerticalAlignment"]),
    "align_indent_distributed": (f"{A1}.HorizontalAlignment = xlDistributed\n{A1}.IndentLevel = 2", align(A1)),
    "align_indent_then_center": (f"{A1}.IndentLevel = 2\n{A1}.HorizontalAlignment = xlCenter", align(A1)),
    "align_reading_order": (f"{A1}.ReadingOrder = xlRTL", align(A1)),
    "align_add_indent": (f"{A1}.HorizontalAlignment = xlDistributed\n{A1}.AddIndent = True", align(A1)),
    # --- clearing and copying --------------------------------------------------------
    "clear_formats": (f'{A1}.Value = 7\n{A1}.Font.Bold = True\n{A1}.Interior.Color = RGB(255, 255, 0)\n'
                      f"{A1}.Borders(xlEdgeLeft).LineStyle = xlContinuous\n{A1}.HorizontalAlignment = xlCenter\n"
                      f"{A1}.ClearFormats",
                      [f"{A1}.Value", f"{A1}.Font.Bold", f"{A1}.Interior.ColorIndex",
                       f"{A1}.Borders(7).LineStyle", f"{A1}.HorizontalAlignment"]),
    "clear_contents_keeps_format": (f'{A1}.Value = 7\n{A1}.Font.Bold = True\n{A1}.ClearContents',
                                    [f"{A1}.Value", f"{A1}.Font.Bold"]),
    # --- values Excel refuses or treats specially ---------------------------------
    "bad_line_style": (f"{A1}.Borders(xlEdgeLeft).LineStyle = 99", border(A1, 7, ("LineStyle", "Weight"))),
    "bad_weight": (f"{A1}.Borders(xlEdgeLeft).Weight = 3", border(A1, 7, ("LineStyle", "Weight"))),
    "bad_font_color_index_zero": (f"{A1}.Font.ColorIndex = 0", [f"{A1}.Font.ColorIndex"]),
    "bad_fill_color_index_zero": (f"{A1}.Interior.ColorIndex = 0", fill(A1)),
    "bad_fill_color_index": (f"{A1}.Interior.ColorIndex = 57", fill(A1)),
    "bad_underline": (f"{A1}.Font.Underline = 3", [f"{A1}.Font.Underline"]),
    "bad_tint": (f"{A1}.Font.ThemeColor = xlThemeColorAccent1\n{A1}.Font.TintAndShade = 1.5", font(A1)),
    "bad_orientation": (f"{A1}.Orientation = -91", [f"{A1}.Orientation"]),
    "bad_reading_order": (f"{A1}.ReadingOrder = 5", [f"{A1}.ReadingOrder"]),
    "big_font_size": (f"{A1}.Font.Size = 409.6", [f"{A1}.Font.Size"]),
    "fill_pattern_automatic": (f"{A1}.Interior.Pattern = xlPatternAutomatic", fill(A1)),
    "fill_pattern_color_on_none": (f"{A1}.Interior.PatternColor = RGB(0, 0, 255)", fill(A1)),
    "fill_pattern_color_index_on_none": (f"{A1}.Interior.PatternColorIndex = 5", fill(A1)),
    "fill_tint_on_none": (f"{A1}.Interior.TintAndShade = 0.5", fill(A1)),
    "fill_tint_on_index": (f"{A1}.Interior.ColorIndex = 6\n{A1}.Interior.TintAndShade = 0.5", fill(A1)),
    "font_tint_on_automatic": (f"{A1}.Font.ColorIndex = xlColorIndexAutomatic\n{A1}.Font.TintAndShade = 0.5", font(A1)),
    "font_tint_on_index": (f"{A1}.Font.ColorIndex = 3\n{A1}.Font.TintAndShade = 0.5", font(A1)),
    "border_color_index_none": (f"{A1}.Borders(xlEdgeLeft).LineStyle = xlContinuous\n"
                                f"{A1}.Borders(xlEdgeLeft).ColorIndex = xlColorIndexNone", border(A1, 7)),
    "border_tint_on_automatic": (f"{A1}.Borders(xlEdgeLeft).LineStyle = xlContinuous\n"
                                 f"{A1}.Borders(xlEdgeLeft).TintAndShade = 0.5", border(A1, 7)),
    "border_tint_on_none": (f"{A1}.Borders(xlEdgeLeft).TintAndShade = 0.5", border(A1, 7)),
    "number_format_mixed": (f'{A1}.NumberFormat = "0.00"', [f"{AB}.NumberFormat", f"{A1}.NumberFormat"]),
    "number_format_same": (f'{AB}.NumberFormat = "0.00"', [f"{AB}.NumberFormat"]),
    "clear_formats_merged": ('Range("A1:B2").Merge\nRange("A1:B2").ClearFormats',
                             ['Range("A1:B2").MergeCells', 'Range("B2").MergeArea.Address']),
    "font_name_back": (f'{A1}.Font.Name = "Arial"\n{A1}.Font.Name = "Aptos Narrow"', font(A1)),
    "copy_format": (f'{A1}.Value = 7\n{A1}.Font.Italic = True\n{A1}.Interior.Color = RGB(0, 255, 0)\n'
                    f'{A1}.Borders(xlEdgeTop).LineStyle = xlContinuous\n{A1}.WrapText = True\n'
                    f'{A1}.Copy Range("C3")',
                    ['Range("C3").Value', 'Range("C3").Font.Italic', 'Range("C3").Interior.Color',
                     'Range("C3").Borders(8).LineStyle', 'Range("C3").WrapText']),
}

# The 56-colour palette, read both ways, and the nearest-palette answer for colours not on it.
PALETTE = {
    "palette_font": ('For i = 1 To 56\nRange("A1").Font.ColorIndex = i\n'
                     'Report = Report & CStr(Range("A1").Font.Color) & ","\nNext i', []),
    "palette_fill": ('For i = 1 To 56\nRange("A1").Interior.ColorIndex = i\n'
                     'Report = Report & CStr(Range("A1").Interior.Color) & ","\nNext i', []),
    "palette_nearest": ('For Each c In Array(RGB(250, 5, 5), RGB(10, 10, 10), RGB(128, 128, 0), '
                        'RGB(200, 200, 255), RGB(255, 128, 64), RGB(1, 2, 3), RGB(100, 150, 200), '
                        'RGB(255, 255, 254))\nRange("A1").Font.Color = c\n'
                        'Report = Report & CStr(Range("A1").Font.ColorIndex) & ","\nNext c', []),
    # Colours where Euclidean, Manhattan and Chebyshev distance each pick a different palette entry.
    "palette_metric": ('For Each c In Array(RGB(66, 126, 203), RGB(116, 240, 100), RGB(100, 182, 163), '
                       'RGB(106, 254, 112), RGB(196, 211, 107), RGB(229, 127, 55))\nRange("A1").Font.Color = c\n'
                       'Report = Report & CStr(Range("A1").Font.ColorIndex) & ","\nNext c', []),
    "border_style_after_weight": ('On Error Resume Next\nFor Each k In Array(1, -4138, 4)\n'
                                  'For Each c In Array(1, -4115, 4, 5, -4118, -4119, 13)\nErr.Clear\n'
                                  'Range("A1").Borders(xlEdgeLeft).LineStyle = xlNone\n'
                                  'Range("A1").Borders(xlEdgeLeft).Weight = k\n'
                                  'Range("A1").Borders(xlEdgeLeft).LineStyle = c\n'
                                  'Report = Report & Err.Number & ":" & Range("A1").Borders(xlEdgeLeft).LineStyle '
                                  '& ":" & Range("A1").Borders(xlEdgeLeft).Weight & ","\nNext c\nNext k', []),
    "border_line_styles": ('On Error Resume Next\nFor Each k In Array(1, -4115, 4, 5, -4118, -4119, 13, -4142)\n'
                           'Err.Clear\nRange("A1").Borders(xlEdgeLeft).LineStyle = k\n'
                           'Report = Report & Err.Number & ":" & Range("A1").Borders(xlEdgeLeft).LineStyle & ":" '
                           '& Range("A1").Borders(xlEdgeLeft).Weight & ","\nNext k', []),
    "border_weights": ('On Error Resume Next\nFor Each k In Array(1, 2, -4138, 4)\n'
                       'Err.Clear\nRange("A1").Borders(xlEdgeLeft).LineStyle = xlContinuous\n'
                       'Range("A1").Borders(xlEdgeLeft).Weight = k\n'
                       'Report = Report & Err.Number & ":" & Range("A1").Borders(xlEdgeLeft).LineStyle & ":" '
                       '& Range("A1").Borders(xlEdgeLeft).Weight & ","\nNext k', []),
    "border_style_weight_pairs": ('On Error Resume Next\nFor Each c In Array(1, -4115, 4, 5, -4118, -4119, 13)\n'
                                  'For Each k In Array(1, 2, -4138, 4)\nErr.Clear\n'
                                  'Range("A1").Borders(xlEdgeLeft).LineStyle = xlNone\n'
                                  'Range("A1").Borders(xlEdgeLeft).LineStyle = c\n'
                                  'Range("A1").Borders(xlEdgeLeft).Weight = k\n'
                                  'Report = Report & Err.Number & ":" & Range("A1").Borders(xlEdgeLeft).LineStyle '
                                  '& ":" & Range("A1").Borders(xlEdgeLeft).Weight & ","\nNext k\nNext c', []),
    "tints": ('For Each t In Array(-0.9, -0.5, -0.25, -0.1, 0, 0.1, 0.25, 0.4, 0.6, 0.8, 0.95)\n'
              'For Each k In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)\n'
              'Range("A1").Font.ThemeColor = k\nRange("A1").Font.TintAndShade = t\n'
              'Report = Report & CStr(Range("A1").Font.Color) & ":" & CStr(Range("A1").Font.TintAndShade) & ","\n'
              "Next k\nNext t", []),
}


def body(setup: str, reads: list[str]) -> str:
    # The setup runs under Resume Next too, so a property Excel refuses
    # records its error number as S<n> instead of ending the probe.
    lines = ['Application.DisplayAlerts = False', 'Range("A1:D6").Clear', "On Error Resume Next", setup,
             'If Err.Number <> 0 Then Report = "S" & Err.Number & ";"']
    for expression in reads:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then Report = Report & "E" & Err.Number & ";" '
                  'Else Report = Report & Show(v) & ";"']
    return "\n".join(line for line in lines if line) + "\n"


def module(text: str) -> str:
    return (HELPER + "Public Function Report() As String\n"
            "Dim v As Variant, i As Long, c As Variant, t As Variant, k As Variant\n" + text + "End Function\n")


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        cases = {name: module(body(setup, reads)) for name, (setup, reads) in PROBES.items()}
        cases.update({name: module(body("", []) .replace("On Error Resume Next\n", "") + loop + "\n")
                      for name, (loop, _) in PALETTE.items()})
        for name, code in cases.items():
            result = excel.run_vba(code, "Report", timeout=120.0)
            reported = str(result.value) if result.ok else f"ERROR {result.outcome}: {result.message}"
            records.append({"name": name, "code": code, "reported": reported})
            print(name, reported[:200], flush=True)
    (ROOT / "tests/fixtures/range_format.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
