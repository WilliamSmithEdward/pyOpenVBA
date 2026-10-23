"""Data bars and icon sets, as FormatConditions.AddDatabar and AddIconSetCondition make them.

Every layout sits on a sheet of its own in one new workbook: the numbers 0
to 100 in tens down A1:A11, A12 blank, a data bar or an icon set over
A1:A12 made one way or another. Excel draws a bar or an icon and keeps no
trace of it in the object model, so the dump reads the rules: every rule
on the sheet with its TypeName, Type, Priority and range; for a data bar
its MinPoint and MaxPoint, the colours of the bar, its border, its axis
and its negative bars as FormatColor's Color, ThemeColor and TintAndShade,
the fill type, direction, lengths, ShowValue, StopIfTrue, Formula and
ScopeType; for an icon set its IconSet's ID and Count, ReverseOrder,
ShowIconOnly, PercentileValues and each criterion's Index, Type, Value,
Operator and Icon. A1's and A11's DisplayFormat fill go first, and what
the last call returned or the error it raised. Every property read that
fails reads E and its error number.

The layouts ask for the defaults, each kind of point and criterion, each
setting of each property, each of Excel's twenty icon sets, custom icons,
the errors of values out of range, and how bars and icons meet other
rules.

Then the workbook is saved and opened again: each sheet's
conditionalFormatting elements and its extLst go beside the dumps Excel
reads back. Last, a few of the opened sheets are changed and the workbook
saved once more, to show how Excel writes rules it read from a file.

    python scripts/measure_data_bars_icons.py

writes tests/fixtures/data_bars_icons/: data_bars_icons.xlsx as Excel
saved it, data_bars_icons_resaved.xlsx as it saved it again, and
data_bars_icons.json, which tests/test_excel_data_bars_icons.py replays.
"""

from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_colour_scales import _targets  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "data_bars_icons"

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(CLng(v))
    ElseIf IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "Object:" & TypeName(v)
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function P(o As Object, name As String) As String
    Dim v As Variant
    If o Is Nothing Then P = "None": Exit Function
    On Error Resume Next
    Err.Clear
    v = CallByName(o, name, VbGet)
    If Err.Number <> 0 Then P = "E" & Err.Number Else P = Show(v)
End Function

Private Function Child(o As Object, name As String) As Object
    Set Child = Nothing
    If o Is Nothing Then Exit Function
    On Error Resume Next
    Set Child = CallByName(o, name, VbGet)
End Function

Private Function Tint(c As Object) As String
    Tint = P(c, "Color") & "/" & P(c, "ThemeColor") & "/" & P(c, "TintAndShade")
End Function

Private Function Bar(f As Object) As String
    Dim out As String, n As Object
    out = " min=" & P(Child(f, "MinPoint"), "Type") & "/" & P(Child(f, "MinPoint"), "Value")
    out = out & " max=" & P(Child(f, "MaxPoint"), "Type") & "/" & P(Child(f, "MaxPoint"), "Value")
    out = out & " bar=" & Tint(Child(f, "BarColor")) & " fill=" & P(f, "BarFillType")
    out = out & " border=" & P(Child(f, "BarBorder"), "Type") & "/" & Tint(Child(Child(f, "BarBorder"), "Color"))
    out = out & " dir=" & P(f, "Direction") & " show=" & P(f, "ShowValue")
    out = out & " len=" & P(f, "PercentMin") & "/" & P(f, "PercentMax")
    out = out & " axis=" & P(f, "AxisPosition") & "/" & Tint(Child(f, "AxisColor"))
    Set n = Child(f, "NegativeBarFormat")
    out = out & " neg=" & P(n, "ColorType") & "/" & Tint(Child(n, "Color")) & "/" & P(n, "BorderColorType") & _
        "/" & Tint(Child(n, "BorderColor"))
    Bar = out
End Function

Private Function Icons(f As Object) As String
    Dim out As String, c As Object, s As Object, i As Long, n As Long
    Set s = Child(f, "IconSet")
    out = " set=" & P(s, "ID") & "/" & P(s, "Count")
    out = out & " rev=" & P(f, "ReverseOrder") & " only=" & P(f, "ShowIconOnly") & " pv=" & P(f, "PercentileValues")
    On Error Resume Next
    n = -1
    n = f.IconCriteria.Count
    out = out & " n=" & n
    For i = 1 To n
        Set c = Nothing
        Set c = f.IconCriteria(i)
        out = out & " c" & P(c, "Index") & "=" & P(c, "Type") & "/" & P(c, "Value") & "/" & P(c, "Operator") & _
            "/" & P(c, "Icon")
    Next
    Icons = out
End Function

Private Function Points(f As Object) As String
    Dim out As String, c As Object, i As Long, n As Long
    On Error Resume Next
    n = -1
    n = f.ColorScaleCriteria.Count
    For i = 1 To n
        Set c = Nothing
        Set c = f.ColorScaleCriteria(i)
        out = out & " c" & i & "=" & P(c, "Type") & "/" & P(c, "Value")
    Next
    Points = out
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, f As Object, i As Long, n As Long
    On Error Resume Next
    out = "fill=" & ws.Range("A1").DisplayFormat.Interior.Color & "," & ws.Range("A11").DisplayFormat.Interior.Color
    n = ws.Cells.FormatConditions.Count
    out = out & " n=" & n
    For i = 1 To n
        Set f = Nothing
        Set f = ws.Cells.FormatConditions(i)
        out = out & " [" & TypeName(f) & ":" & P(f, "Type") & ":" & P(f, "Priority") & ":" & f.AppliesTo.Address
        out = out & " stop=" & P(f, "StopIfTrue") & " formula=" & P(f, "Formula") & " scope=" & P(f, "ScopeType")
        If TypeName(f) = "Databar" Then out = out & Bar(f)
        If TypeName(f) = "IconSetCondition" Then out = out & Icons(f)
        If TypeName(f) = "ColorScale" Then out = out & Points(f)
        out = out & "]"
    Next
    Dump = out
End Function
'''

#: The numbers every layout works on: 0 to 100 in tens down A1:A11, A12 blank.
LIST = "For i = 1 To 11: ws.Cells(i, 1).Value = (i - 1) * 10: Next\n"
BAR = 'Set f = ws.Range("A1:A12").FormatConditions.AddDatabar\n'
ICONS = 'Set f = ws.Range("A1:A12").FormatConditions.AddIconSetCondition\n'
FILL_RULE = ('Set g = ws.Range("A1:A12").FormatConditions.Add(xlCellValue, xlGreater, "50")\n'
             "g.Interior.Color = RGB(0, 0, 255)\n")
#: Text, a blank, an error and a logical among the numbers, and each one's Text read back with A2's.
ODD_CELLS = ('ws.Range("A3").Value = "text"\nws.Range("A5").ClearContents\nws.Range("A7").Formula = "=1/0"\n'
             'ws.Range("A9").Value = True\n')
TEXTS = 'v = ws.Range("A2").Text & "~" & ws.Range("A3").Text & "~" & ws.Range("A5").Text & "~" & ' \
        'ws.Range("A7").Text & "~" & ws.Range("A9").Text\n'
SCALE3 = 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
#: A formula naming the first sheet.
OTHER = '"=bar_default!$A$2"'
#: XlIconSet, 1 to 20.
ICON_SETS = ("xl3Arrows", "xl3ArrowsGray", "xl3Flags", "xl3TrafficLights1", "xl3TrafficLights2", "xl3Signs",
             "xl3Symbols", "xl3Symbols2", "xl4Arrows", "xl4ArrowsGray", "xl4RedToBlack", "xl4CRV",
             "xl4TrafficLights", "xl5Arrows", "xl5ArrowsGray", "xl5CRV", "xl5Quarters", "xl3Stars", "xl3Triangles",
             "xl5Boxes")


def point(which: str, kind: str, value: object | None = None) -> str:
    return f"f.{which}.Modify {kind}" + (f", {value}" if value is not None else "") + "\n"


#: name -> setup; each works on ws, the layout's own sheet, its rule in f, keeping the last call's result in v.
LAYOUTS: dict[str, str] = {
    "bar_default": BAR,
    "bar_numbers": BAR + point("MinPoint", "xlConditionValueNumber", 20) +
                   point("MaxPoint", "xlConditionValueNumber", 80),
    "bar_percent": BAR + point("MinPoint", "xlConditionValuePercent", 10) +
                   point("MaxPoint", "xlConditionValuePercent", 90),
    "bar_percentile": BAR + point("MinPoint", "xlConditionValuePercentile", 25) +
                      point("MaxPoint", "xlConditionValuePercentile", 75),
    "bar_formula": 'ws.Range("C1").Value = 30\n' + BAR + point("MinPoint", "xlConditionValueFormula", '"=$C$1"'),
    "bar_lowest": BAR + point("MinPoint", "xlConditionValueLowestValue") +
                  point("MaxPoint", "xlConditionValueHighestValue"),
    "bar_min_highest": BAR + point("MinPoint", "xlConditionValueHighestValue"),
    "bar_max_lowest": BAR + point("MaxPoint", "xlConditionValueLowestValue"),
    "bar_min_automax": BAR + point("MinPoint", "xlConditionValueAutomaticMax"),
    "bar_max_automin": BAR + point("MaxPoint", "xlConditionValueAutomaticMin"),
    "bar_modify_bare": BAR + point("MinPoint", "xlConditionValueNumber"),
    "bar_modify_none": BAR + "f.MinPoint.Modify\n",
    "bar_value_set": BAR + "f.MinPoint.Value = 5\n",
    "bar_type_set": BAR + "f.MinPoint.Type = xlConditionValueNumber\n",
    "bar_number_text": BAR + point("MinPoint", "xlConditionValueNumber", '"abc"'),
    "bar_number_string": BAR + point("MinPoint", "xlConditionValueNumber", '"25"'),
    "bar_percent_past": BAR + point("MinPoint", "xlConditionValuePercent", 150),
    "bar_colour": BAR + "f.BarColor.Color = RGB(255, 0, 0)\n",
    "bar_theme": BAR + "f.BarColor.ThemeColor = xlThemeColorAccent2\nf.BarColor.TintAndShade = -0.25\n",
    "bar_solid": BAR + "f.BarFillType = xlDataBarFillSolid\n",
    "bar_fill_bad": BAR + "f.BarFillType = 2\n",
    "bar_border": BAR + "f.BarBorder.Type = xlDataBarBorderSolid\nf.BarBorder.Color.Color = RGB(0, 0, 255)\n",
    "bar_no_border": BAR + "f.BarBorder.Type = xlDataBarBorderNone\n",
    "bar_rtl": BAR + "f.Direction = xlRTL\n",
    "bar_ltr": BAR + "f.Direction = xlLTR\n",
    "bar_direction_bad": BAR + "f.Direction = 5\n",
    "bar_hide_value": BAR + "f.ShowValue = False\n",
    "bar_lengths": BAR + "f.PercentMin = 10\nf.PercentMax = 90\n",
    "bar_lengths_crossed": BAR + "f.PercentMin = 80\nf.PercentMax = 20\n",
    "bar_length_past": BAR + "f.PercentMax = 101\n",
    "bar_length_negative": BAR + "f.PercentMin = -1\n",
    "bar_axis_mid": BAR + "f.AxisPosition = xlDataBarAxisMidpoint\nf.AxisColor.Color = RGB(255, 0, 255)\n",
    "bar_axis_none": BAR + "f.AxisPosition = xlDataBarAxisNone\n",
    "bar_negative_same": BAR + "f.NegativeBarFormat.ColorType = xlDataBarSameAsPositive\n",
    "bar_negative_colours": BAR + "f.NegativeBarFormat.Color.Color = RGB(0, 255, 0)\n"
                                  "f.NegativeBarFormat.BorderColorType = xlDataBarSameAsPositive\n",
    "bar_negative_border": BAR + "f.NegativeBarFormat.BorderColor.Color = RGB(0, 128, 0)\n",
    "bar_negatives": "For i = 1 To 11: ws.Cells(i, 1).Value = (i - 6) * 10: Next\n" + BAR,
    "bar_type_names": BAR + 'v = TypeName(f) & ":" & TypeName(f.MinPoint) & ":" & TypeName(f.BarColor) & ":" & '
                            'TypeName(f.BarBorder) & ":" & TypeName(f.NegativeBarFormat) & ":" & '
                            'TypeName(f.AxisColor) & ":" & TypeName(f.BarBorder.Color)\n',
    "bar_stop": BAR + "v = f.StopIfTrue\nf.StopIfTrue = True\n",
    "bar_formula_set": BAR + 'f.Formula = "=$A1>20"\n',
    "bar_moved": BAR + 'f.ModifyAppliesToRange ws.Range("A1:A6")\n',
    "bar_deleted": BAR + "f.Delete\n",
    "bar_two_areas": 'Set f = ws.Range("A1:A6,C1:C6").FormatConditions.AddDatabar\n',
    "bar_other_sheet": BAR + point("MinPoint", "xlConditionValueFormula", '"=bar_default!$A$2"'),
    "bar_then_fill": BAR + FILL_RULE + "f.SetFirstPriority\n",
    "icons_default": ICONS,
    **{"icons_" + name.removeprefix("xl"): ICONS + f"f.IconSet = ws.Parent.IconSets({name})\n" for name in ICON_SETS},
    "icons_set_by_set": ICONS + "Set f.IconSet = ws.Parent.IconSets(xl3Flags)\n",
    "icons_criteria": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 30\n"
                              "f.IconCriteria(3).Operator = xlGreater\n",
    "icons_percentile": ICONS + "f.IconCriteria(2).Type = xlConditionValuePercentile\n"
                                "f.IconCriteria(3).Type = xlConditionValuePercentile\n",
    "icons_formula": 'ws.Range("C1").Value = 30\n' + ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                                             'f.IconCriteria(2).Value = "=$C$1"\n',
    "icons_first_type": ICONS + "f.IconCriteria(1).Type = xlConditionValueNumber\n",
    "icons_first_value": ICONS + "f.IconCriteria(1).Value = 5\n",
    "icons_first_operator": ICONS + "f.IconCriteria(1).Operator = xlGreater\n",
    "icons_operator_bad": ICONS + "f.IconCriteria(2).Operator = xlEqual\n",
    "icons_value_text": ICONS + 'f.IconCriteria(2).Value = "abc"\n',
    "icons_value_string": ICONS + 'f.IconCriteria(2).Value = "40"\n',
    "icons_type_lowest": ICONS + "f.IconCriteria(2).Type = xlConditionValueLowestValue\n",
    "icons_reverse": ICONS + "f.ReverseOrder = True\n",
    "icons_only": ICONS + "f.ShowIconOnly = True\n",
    "icons_custom": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\n",
    "icons_no_icon": ICONS + "f.IconCriteria(2).Icon = xlIconNoCellIcon\n",
    "icons_custom_then_set": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\n"
                                     "f.IconSet = ws.Parent.IconSets(xl3Flags)\n",
    "icons_criteria_then_set": ICONS + "f.IconCriteria(2).Value = 50\nf.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_type_names": ICONS + 'v = TypeName(f) & ":" & TypeName(f.IconCriteria) & ":" & '
                                'TypeName(f.IconCriteria(1)) & ":" & TypeName(f.IconSet) & ":" & '
                                'TypeName(ws.Parent.IconSets) & ":" & TypeName(f.IconSet.Item(1))\n',
    "icons_sets": ICONS + 'v = ws.Parent.IconSets.Count & ":" & ws.Parent.IconSets(1).ID & ":" & '
                          'ws.Parent.IconSets(xl5Boxes).Count & ":" & f.IconSet.Item(2).Index\n',
    "icons_criterion_zero": ICONS + "v = f.IconCriteria(0).Type\n",
    "icons_criterion_past": ICONS + "v = f.IconCriteria(4).Type\n",
    "icons_set_zero": ICONS + "f.IconSet = ws.Parent.IconSets(0)\n",
    "icons_stop": ICONS + "v = f.StopIfTrue\nf.StopIfTrue = True\n",
    "icons_moved": ICONS + 'f.ModifyAppliesToRange ws.Range("A1:A6")\n',
    "icons_deleted": ICONS + "f.Delete\n",
    "icons_percentile_set": ICONS + "f.PercentileValues = True\n",
    "mixed": 'Set k = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n' + BAR +
             'Set h = ws.Range("A1:A12").FormatConditions.AddIconSetCondition\n' + FILL_RULE + "f.SetFirstPriority\n",
    # Which lengths the plain rule keeps, and what puts a bar in the extension.
    "bar_len_0_90": BAR + "f.PercentMax = 90\n",
    "bar_len_10_100": BAR + "f.PercentMin = 10\n",
    "bar_len_5_95": BAR + "f.PercentMin = 5\nf.PercentMax = 95\n",
    "bar_len_50_50": BAR + "f.PercentMin = 50\nf.PercentMax = 50\n",
    "bar_len_0_50": BAR + "f.PercentMax = 50\n",
    "bar_old_solid": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.BarFillType = xlDataBarFillSolid\n",
    "bar_old_border": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.BarBorder.Type = xlDataBarBorderSolid\n",
    "bar_old_rtl": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.Direction = xlRTL\n",
    "bar_old_ltr": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.Direction = xlLTR\n",
    "bar_old_axis": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.AxisPosition = xlDataBarAxisAutomatic\n",
    "bar_old_negative": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.NegativeBarFormat.ColorType = xlDataBarColor\n",
    "bar_old_auto": BAR + "f.PercentMin = 10\nf.PercentMax = 90\n" + point("MinPoint", "xlConditionValueAutomaticMin"),
    "bar_old_hidden": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.ShowValue = False\n",
    "bar_old_theme": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.BarColor.ThemeColor = xlThemeColorAccent4\n",
    "bar_old_axis_colour": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.AxisColor.Color = RGB(1, 2, 3)\n",
    "bar_old_border_colour": BAR + "f.PercentMin = 10\nf.PercentMax = 90\nf.BarBorder.Color.Color = RGB(1, 2, 3)\n",
    "bar_old_formula": BAR + "f.PercentMin = 10\nf.PercentMax = 90\n" +
                       point("MinPoint", "xlConditionValueFormula", '"=bar_default!$A$2"'),
    # Points.
    "bar_auto": BAR + point("MinPoint", "xlConditionValueAutomaticMin") +
                point("MaxPoint", "xlConditionValueAutomaticMax"),
    "bar_auto_value": BAR + point("MinPoint", "xlConditionValueAutomaticMin", 5),
    "bar_lowest_value": BAR + point("MinPoint", "xlConditionValueLowestValue", 5),
    "bar_percent_negative": BAR + point("MinPoint", "xlConditionValuePercent", -1),
    "bar_percentile_past": BAR + point("MaxPoint", "xlConditionValuePercentile", 101),
    "bar_percent_formula": 'ws.Range("C1").Value = 30\n' + BAR + point("MinPoint", "xlConditionValuePercent", '"=$C$1"'),
    "bar_number_negative": BAR + point("MinPoint", "xlConditionValueNumber", -20),
    "bar_number_fraction": BAR + point("MinPoint", "xlConditionValueNumber", 0.1),
    "bar_type_bad": BAR + point("MinPoint", "8", 5),
    "bar_back_from_sheet": BAR + point("MinPoint", "xlConditionValueFormula", '"=bar_default!$A$2"') +
                           point("MinPoint", "xlConditionValueNumber", 5),
    # Negative bars, the axis and the border.
    "bar_neg_colour_type": BAR + "f.NegativeBarFormat.ColorType = xlDataBarColor\n",
    "bar_neg_colour_set": BAR + "f.NegativeBarFormat.ColorType = xlDataBarColor\n"
                                "f.NegativeBarFormat.Color.Color = RGB(0, 255, 0)\n",
    "bar_neg_theme": BAR + "f.NegativeBarFormat.ColorType = xlDataBarColor\n"
                           "f.NegativeBarFormat.Color.ThemeColor = xlThemeColorAccent6\n",
    "bar_neg_border_type": BAR + "f.NegativeBarFormat.BorderColorType = xlDataBarColor\n",
    "bar_neg_border_set": BAR + "f.BarBorder.Type = xlDataBarBorderSolid\n"
                                "f.NegativeBarFormat.BorderColorType = xlDataBarColor\n"
                                "f.NegativeBarFormat.BorderColor.Color = RGB(0, 128, 0)\n",
    "bar_neg_bad_type": BAR + "f.NegativeBarFormat.ColorType = 2\n",
    "bar_axis_auto": BAR + "f.AxisPosition = xlDataBarAxisAutomatic\n",
    "bar_axis_colour_none": BAR + "f.AxisColor.Color = RGB(1, 2, 3)\n",
    "bar_axis_bad": BAR + "f.AxisPosition = 3\n",
    "bar_axis_theme": BAR + "f.AxisPosition = xlDataBarAxisMidpoint\nf.AxisColor.ThemeColor = xlThemeColorDark1\n",
    "bar_border_colour_none": BAR + "f.BarBorder.Color.Color = RGB(0, 0, 255)\n",
    "bar_border_only_type": BAR + "f.BarBorder.Type = xlDataBarBorderSolid\n",
    "bar_border_bad": BAR + "f.BarBorder.Type = 2\n",
    "bar_solid_border": BAR + "f.BarFillType = xlDataBarFillSolid\nf.BarBorder.Type = xlDataBarBorderSolid\n",
    "bar_rgb_tint": BAR + "f.BarColor.TintAndShade = 0.5\n",
    "bar_everything": BAR + point("MinPoint", "xlConditionValueAutomaticMin") +
                      point("MaxPoint", "xlConditionValueAutomaticMax") +
                      "f.BarFillType = xlDataBarFillSolid\nf.BarBorder.Type = xlDataBarBorderSolid\n"
                      "f.Direction = xlRTL\nf.AxisPosition = xlDataBarAxisMidpoint\n"
                      "f.NegativeBarFormat.ColorType = xlDataBarColor\n"
                      "f.NegativeBarFormat.BorderColorType = xlDataBarColor\n"
                      "f.PercentMin = 5\nf.PercentMax = 95\nf.ShowValue = False\n"
                      "f.BarColor.ThemeColor = xlThemeColorAccent3\n",
    "bar_formula_cleared": BAR + 'f.Formula = "=$A1>20"\nf.Formula = ""\n',
    "bar_formula_bad": BAR + 'f.Formula = "=("\n',
    # Several rules in the extension.
    "bars_two_same": BAR + 'Set g = ws.Range("A1:A12").FormatConditions.AddDatabar\n',
    "bars_two_ranges": 'Set f = ws.Range("A1:A6").FormatConditions.AddDatabar\n'
                       'Set g = ws.Range("A7:A12").FormatConditions.AddDatabar\n',
    "bar_and_stars": BAR + 'Set h = ws.Range("A1:A12").FormatConditions.AddIconSetCondition\n'
                           "h.IconSet = ws.Parent.IconSets(xl3Stars)\n",
    "stars_then_bar": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\n"
                              'Set g = ws.Range("A1:A12").FormatConditions.AddDatabar\n',
    "stars_then_bar_first": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\n"
                                    'Set g = ws.Range("A1:A12").FormatConditions.AddDatabar\ng.SetLastPriority\n',
    # Changing an icon set keeps or resets the criteria.
    "icons_c2_then_5": ICONS + "f.IconCriteria(2).Value = 50\nf.IconSet = ws.Parent.IconSets(xl5Arrows)\n",
    "icons_type_then_4": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\n"
                                 "f.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_5_then_3": ICONS + "f.IconSet = ws.Parent.IconSets(xl5Arrows)\nf.IconCriteria(2).Value = 10\n"
                              "f.IconSet = ws.Parent.IconSets(xl3Arrows)\n",
    "icons_4_then_3": ICONS + "f.IconSet = ws.Parent.IconSets(xl4Arrows)\nf.IconSet = ws.Parent.IconSets(xl3Arrows)\n",
    "icons_operator_then_4": ICONS + "f.IconCriteria(2).Operator = xlGreater\n"
                                     "f.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_same_then_4": ICONS + "f.IconCriteria(2).Value = 33\nf.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_reverse_then_set": ICONS + "f.ReverseOrder = True\nf.ShowIconOnly = True\n"
                                      "f.IconSet = ws.Parent.IconSets(xl3Flags)\n",
    "icons_pct_then_set": ICONS + "f.PercentileValues = True\nf.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_pct_false": ICONS + "f.PercentileValues = True\nf.PercentileValues = False\n",
    "icons_pct_mixed_false": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.PercentileValues = False\n",
    "icons_value_negative": ICONS + "f.IconCriteria(2).Value = -5\n",
    "icons_value_past": ICONS + "f.IconCriteria(2).Value = 150\n",
    "icons_value_order": ICONS + "f.IconCriteria(2).Value = 80\n",
    "icons_number_values": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 30\n"
                                   "f.IconCriteria(3).Type = xlConditionValueNumber\nf.IconCriteria(3).Value = 70\n",
    "icons_type_bad": ICONS + "f.IconCriteria(2).Type = xlConditionValueAutomaticMax\n",
    "icons_icon_bad": ICONS + "f.IconCriteria(1).Icon = 99\n",
    "icons_custom_all": ICONS + "f.IconCriteria(1).Icon = xlIconRedCross\nf.IconCriteria(2).Icon = xlIconGoldStar\n"
                                "f.IconCriteria(3).Icon = xlIcon2FilledBoxes\n",
    "icons_custom_reverse": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\nf.ReverseOrder = True\n",
    "icons_custom_only": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\nf.ShowIconOnly = True\n",
    "icons_custom_back": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\n"
                                 "f.IconCriteria(1).Icon = xlIconRedCircleWithBorder\n",
    "icons_custom_criteria": ICONS + "f.IconCriteria(1).Icon = xlIconGreenCheck\n"
                                     "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 30\n"
                                     "f.IconCriteria(2).Operator = xlGreater\n",
    "icons_two_areas": 'Set f = ws.Range("A1:A6,C1:C6").FormatConditions.AddIconSetCondition\n',
    "icons_stars_criteria": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.IconCriteria(2).Value = 50\n"
                                    "f.IconCriteria(3).Operator = xlGreater\n",
    "icons_stars_reverse": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.ReverseOrder = True\n"
                                   "f.ShowIconOnly = True\n",
    "icons_other_sheet": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                 'f.IconCriteria(2).Value = "=bar_default!$A$2"\n',
    "icons_formula_set": ICONS + 'f.Formula = "=$A1>20"\n',
    "icons_stars_moved": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\n"
                                 'f.ModifyAppliesToRange ws.Range("A2:A8")\n',
    "icons_stars_deleted": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.Delete\n",
    "icons_stars_then_3": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.IconSet = ws.Parent.IconSets(xl3Arrows)\n",
    # A colour scale's values as strings.
    "scale_value_string": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                          'f.ColorScaleCriteria(2).Value = "40"\n',
    "scale_value_text": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                        'f.ColorScaleCriteria(2).Value = "abc"\n',
    "scale_type_number": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                         "f.ColorScaleCriteria(2).Type = xlConditionValueNumber\n",
    # What a criterion's value becomes when its type changes.
    "icons_t_pctile_pct": ICONS + "f.IconCriteria(2).Type = xlConditionValuePercentile\n"
                                  "f.IconCriteria(2).Type = xlConditionValuePercent\n",
    "icons_t_number_pct": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 40\n"
                                  "f.IconCriteria(2).Type = xlConditionValuePercent\n",
    "icons_t_formula": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n",
    "icons_t_formula_number": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                      'f.IconCriteria(2).Value = "=$C$1"\nf.IconCriteria(2).Type = xlConditionValueNumber\n',
    "icons_t_same": ICONS + "f.IconCriteria(2).Type = xlConditionValuePercent\n",
    "icons_formula_number": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\nf.IconCriteria(2).Value = 30\n",
    "icons_t_pct40_pctile": ICONS + "f.IconCriteria(2).Value = 40\nf.IconCriteria(2).Type = xlConditionValuePercentile\n",
    "icons_t_pct40_formula": ICONS + "f.IconCriteria(2).Value = 40\nf.IconCriteria(2).Type = xlConditionValueFormula\n",
    "icons_t_pctile40_pct": ICONS + "f.IconCriteria(2).Type = xlConditionValuePercentile\nf.IconCriteria(2).Value = 40\n"
                                    "f.IconCriteria(2).Type = xlConditionValuePercent\n",
    "icons_t_formula_pct": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                   'f.IconCriteria(2).Value = "=$C$1"\nf.IconCriteria(2).Type = xlConditionValuePercent\n',
    "icons_t_number_pctile": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 40\n"
                                     "f.IconCriteria(2).Type = xlConditionValuePercentile\n",
    "icons_t_number_formula": ICONS + "f.IconCriteria(2).Type = xlConditionValueNumber\nf.IconCriteria(2).Value = 40\n"
                                      "f.IconCriteria(2).Type = xlConditionValueFormula\n",
    "icons_t_text_pctile": ICONS + 'f.IconCriteria(2).Value = "40"\nf.IconCriteria(2).Type = xlConditionValuePercentile\n',
    "icons_t_c3_number_pct": ICONS + "f.IconCriteria(3).Type = xlConditionValueNumber\nf.IconCriteria(3).Value = 40\n"
                                     "f.IconCriteria(3).Type = xlConditionValuePercent\n",
    "icons_t_4_number_pct": ICONS + "f.IconSet = ws.Parent.IconSets(xl4Arrows)\n"
                                    "f.IconCriteria(3).Type = xlConditionValueNumber\nf.IconCriteria(3).Value = 40\n"
                                    "f.IconCriteria(3).Type = xlConditionValuePercent\n",
    "scale_t_pctile40_pct": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                            "f.ColorScaleCriteria(2).Value = 40\nf.ColorScaleCriteria(2).Type = xlConditionValuePercent\n",
    "scale_t_num_pct": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                       "f.ColorScaleCriteria(2).Type = xlConditionValueNumber\nf.ColorScaleCriteria(2).Value = 40\n"
                       "f.ColorScaleCriteria(2).Type = xlConditionValuePercent\n",
    "scale_t_num_pctile": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                          "f.ColorScaleCriteria(2).Type = xlConditionValueNumber\nf.ColorScaleCriteria(2).Value = 40\n"
                          "f.ColorScaleCriteria(2).Type = xlConditionValuePercentile\n",
    "scale_t_min_pct": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                       "f.ColorScaleCriteria(1).Type = xlConditionValuePercent\n",
    "scale_t_max_number": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                          "f.ColorScaleCriteria(3).Type = xlConditionValueNumber\n",
    "scale_t_num_min": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                       "f.ColorScaleCriteria(2).Type = xlConditionValueNumber\nf.ColorScaleCriteria(2).Value = 40\n"
                       "f.ColorScaleCriteria(2).Type = xlConditionValueLowestValue\n",
    "scale_t_percent": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                       "f.ColorScaleCriteria(2).Type = xlConditionValuePercent\n",
    "scale_t_formula": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                       "f.ColorScaleCriteria(2).Type = xlConditionValueFormula\n",
    "scale_t_same": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(3)\n'
                    "f.ColorScaleCriteria(2).Type = xlConditionValuePercentile\n",
    "bar_formula_number": BAR + point("MinPoint", "xlConditionValueFormula", 30),
    # Whether hiding the value changes what a cell shows.
    "bar_text": BAR + 'v = ws.Range("A2").Text & "~" & ws.Range("A2").DisplayFormat.NumberFormat\n',
    "bar_hidden_text": BAR + 'f.ShowValue = False\nv = ws.Range("A2").Text & "~" & '
                             'ws.Range("A2").DisplayFormat.NumberFormat\n',
    "icons_only_text": ICONS + 'f.ShowIconOnly = True\nv = ws.Range("A2").Text & "~" & '
                               'ws.Range("A2").DisplayFormat.NumberFormat\n',
    # More icon-set switching, negative bars and custom icons.
    "icons_c3op_then_5": ICONS + "f.IconCriteria(2).Value = 50\nf.IconCriteria(3).Operator = xlGreater\n"
                                 "f.IconSet = ws.Parent.IconSets(xl5Arrows)\n",
    "icons_c1op_then_4": ICONS + "f.IconCriteria(1).Operator = xlGreater\nf.IconSet = ws.Parent.IconSets(xl4Arrows)\n",
    "icons_reverse_custom": ICONS + "f.ReverseOrder = True\nf.IconCriteria(1).Icon = xlIconGreenCheck\n",
    "icons_sets_past": ICONS + "v = ws.Parent.IconSets(21).ID\n",
    "bar_old_neg_border": BAR + "f.PercentMin = 10\nf.PercentMax = 90\n"
                                "f.NegativeBarFormat.BorderColorType = xlDataBarColor\n",
    "bar_neg_back": BAR + "f.NegativeBarFormat.ColorType = xlDataBarColor\n"
                          "f.NegativeBarFormat.Color.Color = RGB(0, 255, 0)\n"
                          "f.NegativeBarFormat.ColorType = xlDataBarSameAsPositive\n",
    "bar_neg_again": BAR + "f.NegativeBarFormat.ColorType = xlDataBarColor\n"
                           "f.NegativeBarFormat.Color.Color = RGB(0, 255, 0)\n"
                           "f.NegativeBarFormat.ColorType = xlDataBarSameAsPositive\n"
                           "f.NegativeBarFormat.ColorType = xlDataBarColor\n",
    "bar_neg_colour_first": BAR + "f.NegativeBarFormat.Color.TintAndShade = 0.5\n",
    # A formula naming another sheet, or its own.
    "scale_other_sheet": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(2)\n'
                         "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                         'f.ColorScaleCriteria(1).Value = "=bar_default!$A$2"\n',
    "scale_own_sheet": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(2)\n'
                       "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                       'f.ColorScaleCriteria(1).Value = "=scale_own_sheet!$A$2"\n',
    "bar_own_sheet": BAR + point("MinPoint", "xlConditionValueFormula", '"=bar_own_sheet!$A$2"'),
    # What a hidden value leaves of each kind of cell, and of one a stopping rule covers first.
    "bar_hidden_kinds": ODD_CELLS + BAR + "f.ShowValue = False\n" + TEXTS,
    "icons_hidden_kinds": ODD_CELLS + ICONS + "f.ShowIconOnly = True\n" + TEXTS,
    "bar_hidden_stopped": FILL_RULE + BAR + 'f.ShowValue = False\nv = ws.Range("A2").Text & "~" & ws.Range("A11").Text\n',
    "icons_hidden_stopped": FILL_RULE + ICONS + 'f.ShowIconOnly = True\nv = ws.Range("A2").Text & "~" & '
                                                'ws.Range("A11").Text\n',
    # Changing a point once another holds a formula naming another sheet.
    "bar_sheet_then_max": BAR + point("MinPoint", "xlConditionValueFormula", OTHER) +
                          point("MaxPoint", "xlConditionValueNumber", 80),
    "bar_sheet_then_formula": BAR + point("MinPoint", "xlConditionValueFormula", OTHER) +
                              point("MinPoint", "xlConditionValueFormula", '"=$C$1"'),
    "bar_sheet_then_percent": BAR + point("MinPoint", "xlConditionValueFormula", OTHER) +
                              point("MinPoint", "xlConditionValuePercent", 10),
    "bar_sheet_then_lowest": BAR + point("MinPoint", "xlConditionValueFormula", OTHER) +
                             point("MinPoint", "xlConditionValueLowestValue"),
    "bar_max_sheet_then_min": BAR + point("MaxPoint", "xlConditionValueFormula", OTHER) +
                              point("MinPoint", "xlConditionValueNumber", 5),
    "icons_sheet_then_value": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                      f"f.IconCriteria(2).Value = {OTHER}\nf.IconCriteria(2).Value = 30\n",
    "icons_sheet_then_type": ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                     f"f.IconCriteria(2).Value = {OTHER}\n"
                                     "f.IconCriteria(2).Type = xlConditionValueNumber\n",
    "scale_sheet_then_value": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(2)\n'
                              "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                              f"f.ColorScaleCriteria(1).Value = {OTHER}\n"
                              'f.ColorScaleCriteria(1).Value = "=$A$3"\n',
    # A colour scale's ends made other types.
    "scale_t_c3_percent": SCALE3 + "f.ColorScaleCriteria(3).Type = xlConditionValuePercent\n",
    "scale_t_c3_percentile": SCALE3 + "f.ColorScaleCriteria(3).Type = xlConditionValuePercentile\n",
    "scale_t_c1_percentile": SCALE3 + "f.ColorScaleCriteria(1).Type = xlConditionValuePercentile\n",
    "scale_t_c1_highest": SCALE3 + "f.ColorScaleCriteria(1).Type = xlConditionValueHighestValue\n",
    "scale_t_c3_lowest": SCALE3 + "f.ColorScaleCriteria(3).Type = xlConditionValueLowestValue\n",
    "scale_t_c1_formula": SCALE3 + "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n",
    "scale2_t_c2_percent": 'Set f = ws.Range("A1:A12").FormatConditions.AddColorScale(2)\n'
                           "f.ColorScaleCriteria(2).Type = xlConditionValuePercent\n",
    "scale_t_eight": SCALE3 + "f.ColorScaleCriteria(2).Type = 8\n",
    # A custom icon on another set, criterion types past the measured ones, and a set given as a number.
    "icons_custom_arrows": ICONS + "f.IconSet = ws.Parent.IconSets(xl3Arrows)\nf.IconCriteria(1).Icon = xlIconGreenCheck\n",
    "icons_custom_five": ICONS + "f.IconSet = ws.Parent.IconSets(xl5Arrows)\nf.IconCriteria(3).Icon = xlIconGoldStar\n",
    "icons_first_same_type": ICONS + "f.IconCriteria(1).Type = xlConditionValuePercent\n",
    "icons_type_eight": ICONS + "f.IconCriteria(2).Type = 8\n",
    "icons_set_number": ICONS + "f.IconSet = 3\n",
    # A formula given without "=".
    "bar_formula_bare": 'ws.Range("C1").Value = 30\n' + BAR + point("MinPoint", "xlConditionValueFormula", '"$C$1"'),
    "icons_formula_bare": 'ws.Range("C1").Value = 30\n' + ICONS + "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                                                                  'f.IconCriteria(2).Value = "$C$1"\n',
    "scale_formula_bare": 'ws.Range("C1").Value = 30\n' + SCALE3 + "f.ColorScaleCriteria(2).Type = xlConditionValueFormula\n"
                                                                   'f.ColorScaleCriteria(2).Value = "$C$1"\n',
    # Rules Excel 2007 could read, for a resave to carry into the extension.
    "bars_two_old": 'Set f = ws.Range("A1:A6").FormatConditions.AddDatabar\nf.PercentMin = 10\nf.PercentMax = 90\n'
                    'Set g = ws.Range("A7:A12").FormatConditions.AddDatabar\ng.PercentMin = 10\ng.PercentMax = 90\n',
    "icons_plain_two": 'Set f = ws.Range("A1:A6").FormatConditions.AddIconSetCondition\n'
                       'Set g = ws.Range("A7:A12").FormatConditions.AddIconSetCondition\n',
    "bar_old_then_new": 'Set f = ws.Range("A1:A6").FormatConditions.AddDatabar\nf.PercentMin = 10\nf.PercentMax = 90\n',
    "bars_old_reordered": 'Set f = ws.Range("A1:A6").FormatConditions.AddDatabar\nf.PercentMin = 10\nf.PercentMax = 90\n'
                          'Set g = ws.Range("A7:A12").FormatConditions.AddDatabar\ng.PercentMin = 10\ng.PercentMax = 90\n'
                          "g.SetFirstPriority\n",
    "bars_old_after_new": BAR + 'Set g = ws.Range("A7:A12").FormatConditions.AddDatabar\ng.PercentMin = 10\n'
                                "g.PercentMax = 90\n",
    "bar_negative_type_names": BAR + 'v = TypeName(f.NegativeBarFormat.Color) & ":" & '
                                     "TypeName(f.NegativeBarFormat.BorderColor)\n",
}

#: sheet -> a change made to it once Excel has opened its own file, before it saves the file again.
RESAVE: dict[str, str] = {
    "bar_default": "",
    "bar_lengths": "ws.Cells.FormatConditions(1).PercentMin = 0\n",
    "bar_and_stars": 'ws.Range("A1:A12").FormatConditions.Add(xlCellValue, xlLess, "5").Interior.Color = '
                     "RGB(255, 0, 0)\n",
    "mixed": "ws.Cells.FormatConditions(2).Delete\n",
    "bars_two_ranges": "ws.Cells.FormatConditions(2).BarColor.Color = RGB(0, 128, 0)\n",
    "icons_3Stars": 'ws.Range("A1:A12").FormatConditions.AddDatabar\n',
    "stars_then_bar": "ws.Cells.FormatConditions(2).SetFirstPriority\n",
    "icons_custom": "ws.Cells.FormatConditions(1).ShowIconOnly = True\n",
    "bar_other_sheet": "ws.Cells.FormatConditions(1).BarFillType = xlDataBarFillSolid\n",
    "bars_two_old": "ws.Cells.FormatConditions(2).PercentMin = 0\nws.Cells.FormatConditions(1).PercentMin = 0\n",
    "icons_plain_two": "ws.Cells.FormatConditions(2).IconSet = ws.Parent.IconSets(xl3Stars)\n",
    "bar_old_then_new": 'ws.Range("A7:A12").FormatConditions.AddDatabar\nws.Cells.FormatConditions(1).PercentMin = 0\n',
    "bars_old_reordered": "ws.Cells.FormatConditions(1).PercentMin = 0\nws.Cells.FormatConditions(2).PercentMin = 0\n",
    "bars_old_after_new": "ws.Cells.FormatConditions(2).PercentMin = 0\n",
}


def file_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """Each sheet's conditionalFormatting elements, then the worksheet's own extLst, by sheet name."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = _targets(package, "xl/_rels/workbook.xml.rels")
    found: dict[str, str] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        part = posixpath.normpath(posixpath.join("xl", targets[rid].lstrip("/").removeprefix("xl/")))
        text = package.read(part).decode("utf-8")
        groups = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", text, re.DOTALL)
        rest = re.sub(r"<conditionalFormatting\b.*?</conditionalFormatting>", "", text, flags=re.DOTALL)
        found[name] = "".join(groups) + "".join(re.findall(r"<extLst>.*?</extLst>", rest, re.DOTALL))
    return found


def case_code(index: int, setup: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, v As Variant, i As Long, f As Object, g As Object, h As Object, k As Object",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def resave_code(index: int, step: str) -> str:
    lines = [f"Private Function Resave{index}(ws As Object) As String", "Dim failed As String", "On Error Resume Next",
             "Err.Clear", *step.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f'Resave{index} = failed & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "data_bars_icons.xlsx"
    resaved = OUT / "data_bars_icons_resaved.xlsx"
    for one in (path, resaved):
        if one.exists():
            one.unlink()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, setup) in enumerate(LAYOUTS.items()):
        assert len(name) <= 31, name
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, LIST + setup))
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "<<PHASE>>"', f'Set wb = Workbooks.Open("{path}")',
              "For Each ws In wb.Worksheets", '    out = out & ws.Name & "^" & Dump(ws) & "|"', "Next",
              'out = out & "<<PHASE>>"']
    for index, (name, step) in enumerate(RESAVE.items()):
        build += [f'Set ws = wb.Worksheets("{name}")', f'out = out & ws.Name & "^" & Resave{index}(ws) & "|"']
        bodies.append(resave_code(index, step))
    build += [f'wb.SaveAs "{resaved}", 51', "wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    strip_save_path(resaved)
    before, after, changed = str(result.value).split("<<PHASE>>")
    answers = before.split("|")[: len(LAYOUTS)]
    # An answer holding the separator would shift every answer after it onto the wrong layout.
    assert before.count("|") == len(LAYOUTS), "an answer holds the layout separator"
    reopened = dict(one.split("^", 1) for one in after.split("|") if one)
    resave_answers = dict(one.split("^", 1) for one in changed.split("|") if one)
    with zipfile.ZipFile(path) as package:
        parts = file_parts(package)
    with zipfile.ZipFile(resaved) as package:
        resaved_parts = file_parts(package)
    record = {"helper": HELPER, "list": LIST,
              "layouts": [{"name": name, "setup": setup, "answers": answer, "reopened": reopened.get(name, ""),
                           "file": parts.get(name, "")}
                          for (name, setup), answer in zip(LAYOUTS.items(), answers, strict=True)],
              "resaved": [{"name": name, "step": step, "answers": resave_answers.get(name, ""),
                           "file": resaved_parts.get(name, "")} for name, step in RESAVE.items()]}
    (OUT / "data_bars_icons.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']}: {one['answers'][:900]}")
        if one["reopened"] not in one["answers"]:
            print(f"    reopened: {one['reopened'][:900]}")
        print(f"    {one['file'][:1500]}")
    for one in record["resaved"]:
        print(f"resaved {one['name']}: {one['answers'][:900]}")
        print(f"    {one['file'][:1500]}")


if __name__ == "__main__":
    main()
