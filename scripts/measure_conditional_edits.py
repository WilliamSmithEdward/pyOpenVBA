"""What inserting and deleting rows, columns and cells does to conditional formats.

Every layout sits on a sheet of its own in one new workbook: the numbers
r * 10 + c in A1:J20, rules over a block of them, and one structural
edit: whole rows or columns inserted or deleted, or cells shifted down,
up, right or left. The dump lists every rule on the sheet with its
TypeName, Priority and range, a FormatCondition's Type and formulas, a
colour scale's, data bar's or icon set's point values; then which cells
of B1:B16 and of A2:J2 show a red fill, as the rules over them paint it.

Some layouts make no edit, to show the order Excel reads a range's areas
back in and which rules it joins as it reads them; others follow an edit
with a top 10 that a column delete makes anew, whose priority shows how
many numbers the edit used up.

Then the workbook is saved and opened again, each sheet's
conditionalFormatting elements and extLst going beside the dumps Excel
reads back. Last, some sheets get their edits once Excel has opened the
file, and it is saved again, to show how Excel moves and numbers rules
it read from a file.

    python scripts/measure_conditional_edits.py

writes tests/fixtures/conditional_edits/: conditional_edits.xlsx and
conditional_edits_resaved.xlsx as Excel saved them, and
conditional_edits.json, which tests/test_excel_conditional_edits.py
replays.
"""

from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_data_bars_icons import file_parts  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_edits"

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

Private Function Rules(ws As Object) As String
    Dim out As String, f As Object, i As Long, n As Long, j As Long
    On Error Resume Next
    n = ws.Cells.FormatConditions.Count
    out = "n=" & n
    For i = 1 To n
        Set f = Nothing
        Set f = ws.Cells.FormatConditions(i)
        out = out & " [" & TypeName(f) & ":" & P(f, "Priority") & ":" & f.AppliesTo.Address(False, False)
        Select Case TypeName(f)
        Case "FormatCondition"
            out = out & ":" & P(f, "Type") & ":" & P(f, "Formula1") & ":" & P(f, "Formula2")
        Case "ColorScale"
            For j = 1 To f.ColorScaleCriteria.Count
                out = out & ":" & P(f.ColorScaleCriteria(j), "Value")
            Next
        Case "Databar"
            out = out & ":" & P(f.MinPoint, "Value") & ":" & P(f.MaxPoint, "Value")
        Case "IconSetCondition"
            For j = 1 To f.IconCriteria.Count
                out = out & ":" & P(f.IconCriteria(j), "Value")
            Next
        End Select
        out = out & "]"
    Next
    Rules = out
End Function

Private Function Painted(ws As Object) As String
    Dim out As String, r As Long, c As Long
    On Error Resume Next
    out = " B="
    For r = 1 To 16
        If ws.Cells(r, 2).DisplayFormat.Interior.Color = RGB(255, 0, 0) Then out = out & "R" Else out = out & "."
    Next
    out = out & " row2="
    For c = 1 To 10
        If ws.Cells(2, c).DisplayFormat.Interior.Color = RGB(255, 0, 0) Then out = out & "R" Else out = out & "."
    Next
    Painted = out
End Function

Private Function Dump(ws As Object) As String
    Dump = Rules(ws) & Painted(ws)
End Function
'''

#: The numbers every layout works on: r * 10 + c in A1:J20.
GRID = "For r = 1 To 20: For c = 1 To 10: ws.Cells(r, c).Value = r * 10 + c: Next: Next\n"


def rule(where: str, kind: str, operator: str = "", first: str = "", second: str = "") -> str:
    """A rule over ``where`` painting red, kept in f; ``kind`` xlCellValue or xlExpression."""
    arguments = ", ".join(one for one in (kind, operator, first, second) if one) if operator else \
        f"{kind}, , {first}"
    return f'Set f = ws.Range("{where}").FormatConditions.Add({arguments})\nf.Interior.Color = RGB(255, 0, 0)\n'


ROWS = rule("B5:B10", "xlCellValue", "xlGreater", '"70"')
COLUMNS = rule("B2:E2", "xlCellValue", "xlGreater", '"22"')
BLOCK = rule("B5:D10", "xlCellValue", "xlGreater", '"70"')
WIDE = rule("B2:E3", "xlCellValue", "xlGreater", '"22"')


def helper_sheet(values: str = "For r = 1 To 5: w2.Cells(r, 1).Value = r * 100: Next\n") -> str:
    """A sheet after the layout's own, named after it with _r, in w2."""
    return ('Set w2 = ws.Parent.Worksheets.Add(After:=ws)\nw2.Name = ws.Name & "_r"\n' + values +
            "ws.Activate\n")


#: name -> (setup, edit); each works on ws, the layout's own sheet, a rule in f, another in g.
LAYOUTS: dict[str, tuple[str, str]] = {
    # Whole rows, a rule over B5:B10.
    "ins_above": (ROWS, "ws.Rows(2).Insert"),
    "ins_top": (ROWS, "ws.Rows(5).Insert"),
    "ins_inside": (ROWS, "ws.Rows(7).Insert"),
    "ins_last": (ROWS, "ws.Rows(10).Insert"),
    "ins_below": (ROWS, "ws.Rows(11).Insert"),
    "ins_far": (ROWS, "ws.Rows(15).Insert"),
    "ins_many_inside": (ROWS, 'ws.Rows("6:8").Insert'),
    "ins_below_shift": (ROWS, "ws.Rows(11).Insert CopyOrigin:=xlFormatFromRightOrBelow"),
    "ins_top_below": (ROWS, "ws.Rows(5).Insert CopyOrigin:=xlFormatFromRightOrBelow"),
    "del_above": (ROWS, "ws.Rows(2).Delete"),
    "del_top": (ROWS, "ws.Rows(5).Delete"),
    "del_inside": (ROWS, "ws.Rows(7).Delete"),
    "del_bottom": (ROWS, "ws.Rows(10).Delete"),
    "del_overlap_top": (ROWS, 'ws.Rows("3:6").Delete'),
    "del_overlap_bottom": (ROWS, 'ws.Rows("9:12").Delete'),
    "del_all": (ROWS, 'ws.Rows("5:10").Delete'),
    "del_cover": (ROWS, 'ws.Rows("4:11").Delete'),
    "del_below": (ROWS, "ws.Rows(11).Delete"),
    # Whole columns, a rule over B2:E2.
    "col_ins_left": (COLUMNS, 'ws.Columns("A").Insert'),
    "col_ins_first": (COLUMNS, 'ws.Columns("B").Insert'),
    "col_ins_inside": (COLUMNS, 'ws.Columns("C").Insert'),
    "col_ins_after": (COLUMNS, 'ws.Columns("F").Insert'),
    "col_del_inside": (COLUMNS, 'ws.Columns("C").Delete'),
    "col_del_first": (COLUMNS, 'ws.Columns("B").Delete'),
    "col_del_all": (COLUMNS, 'ws.Columns("B:E").Delete'),
    "col_del_left": (COLUMNS, 'ws.Columns("A").Delete'),
    # Formulas the rules hold.
    "expr_rel_above": (rule("B5:B10", "xlExpression", first='"=$A5>65"'), "ws.Rows(2).Insert"),
    "expr_rel_inside": (rule("B5:B10", "xlExpression", first='"=$A5>65"'), "ws.Rows(7).Insert"),
    "expr_rel_between": (rule("B5:B10", "xlExpression", first='"=$A4>55"'), "ws.Rows(5).Insert"),
    "expr_rel_delete": (rule("B5:B10", "xlExpression", first='"=$A4>55"'), "ws.Rows(4).Delete"),
    "expr_abs_insert": (rule("B5:B10", "xlExpression", first='"=B5>$D$1"'), "ws.Rows(1).Insert"),
    "expr_abs_deleted": (rule("B5:B10", "xlExpression", first='"=B5>$D$3"'), "ws.Rows(3).Delete"),
    "expr_column_deleted": (rule("B5:B10", "xlExpression", first='"=C5>70"'), 'ws.Columns("C").Delete'),
    "expr_column_inserted": (rule("B5:B10", "xlExpression", first='"=C5>70"'), 'ws.Columns("C").Insert'),
    "cell_between_refs": (rule("B5:B10", "xlCellValue", "xlBetween", '"=$D$1"', '"=$D$2"'), "ws.Rows(1).Insert"),
    "text_rule": ('Set f = ws.Range("B5:B10").FormatConditions.Add(xlTextString, String:="8", '
                  "TextOperator:=xlContains)\nf.Interior.Color = RGB(255, 0, 0)\n", "ws.Rows(2).Insert"),
    "top10_rule": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\nf.Interior.Color = RGB(255, 0, 0)\n',
                   "ws.Rows(7).Delete"),
    "scale_point": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddColorScale(2)\n'
                    "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                    'f.ColorScaleCriteria(1).Value = "=$D$1"\n', "ws.Rows(1).Insert"),
    "bar_point": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddDatabar\n'
                  'f.MinPoint.Modify xlConditionValueFormula, "=$D$1"\n', "ws.Rows(1).Insert"),
    "icons_point": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddIconSetCondition\n'
                    "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                    'f.IconCriteria(2).Value = "=$D$1"\n', "ws.Rows(1).Insert"),
    "bar_inside": ('Set f = ws.Range("B5:B10").FormatConditions.AddDatabar\n', "ws.Rows(7).Insert"),
    "stars_inside": ('Set f = ws.Range("B5:B10").FormatConditions.AddIconSetCondition\n'
                     "f.IconSet = ws.Parent.IconSets(xl3Stars)\n", "ws.Rows(7).Delete"),
    "x14_other_insert": (helper_sheet() + 'Set f = ws.Range("B5:B10").FormatConditions.AddColorScale(2)\n'
                         "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                         'f.ColorScaleCriteria(1).Value = "=x14_other_insert_r!$A$2"\n', "w2.Rows(1).Insert"),
    "x14_other_delete": (helper_sheet() + 'Set f = ws.Range("B5:B10").FormatConditions.AddColorScale(2)\n'
                         "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                         'f.ColorScaleCriteria(1).Value = "=x14_other_delete_r!$A$2"\n', "w2.Rows(2).Delete"),
    # Several areas, groups and rules.
    "multi_area": (rule("B2:B3,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(5).Insert"),
    "multi_area_delete": (rule("B2:B3,B8:B9", "xlCellValue", "xlGreater", '"20"'), 'ws.Rows("2:3").Delete'),
    "two_groups": (ROWS + 'Set g = ws.Range("B12:B14").FormatConditions.Add(xlCellValue, xlLess, "200")\n'
                   "g.Interior.Color = RGB(255, 0, 0)\n", "ws.Rows(11).Delete"),
    "two_rules_one_range": (ROWS + 'Set g = ws.Range("B5:B10").FormatConditions.Add(xlCellValue, xlLess, "60")\n'
                            "g.Interior.Color = RGB(255, 0, 0)\n", "ws.Rows(7).Delete"),
    "groups_meet": (rule("B5:B6", "xlCellValue", "xlGreater", '"50"') +
                    'Set g = ws.Range("B8:B9").FormatConditions.Add(xlCellValue, xlGreater, "50")\n'
                    "g.Interior.Color = RGB(255, 0, 0)\n", "ws.Rows(7).Delete"),
    "same_rule_meet": (rule("B5:B6,B8:B9", "xlCellValue", "xlGreater", '"50"'), "ws.Rows(7).Delete"),
    # Cells shifted.
    "shift_down_above": (ROWS, 'ws.Range("B2").Insert Shift:=xlDown'),
    "shift_down_inside": (ROWS, 'ws.Range("B7").Insert Shift:=xlDown'),
    "shift_down_outside": (ROWS, 'ws.Range("D7").Insert Shift:=xlDown'),
    "shift_down_part": (BLOCK, 'ws.Range("C7").Insert Shift:=xlDown'),
    "shift_down_band": (BLOCK, 'ws.Range("B7:D7").Insert Shift:=xlDown'),
    "shift_down_wider": (BLOCK, 'ws.Range("A7:E7").Insert Shift:=xlDown'),
    "shift_up_inside": (ROWS, 'ws.Range("B7").Delete Shift:=xlUp'),
    "shift_up_part": (BLOCK, 'ws.Range("C7").Delete Shift:=xlUp'),
    "shift_up_all": (ROWS, 'ws.Range("B5:B10").Delete Shift:=xlUp'),
    "shift_up_part_all": (BLOCK, 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "shift_up_above": (ROWS, 'ws.Range("B2").Delete Shift:=xlUp'),
    "shift_right_inside": (COLUMNS, 'ws.Range("C2").Insert Shift:=xlToRight'),
    "shift_right_part": (WIDE, 'ws.Range("C2").Insert Shift:=xlToRight'),
    "shift_left_inside": (COLUMNS, 'ws.Range("C2").Delete Shift:=xlToLeft'),
    "shift_left_part": (WIDE, 'ws.Range("C2").Delete Shift:=xlToLeft'),
    "shift_left_all": (COLUMNS, 'ws.Range("B2:E2").Delete Shift:=xlToLeft'),
    "shift_expr_part": (rule("B5:D10", "xlExpression", first='"=$A5>65"'), 'ws.Range("C7").Insert Shift:=xlDown'),
    # Which kinds of rule each edit keeps in place and which it makes anew, at a new priority (KINDS x EDITS).
    **{f"k_{kind}_{edit}": (setup.replace("RANGE", where), action)
       for kind, (setup, block) in {
           "cell": (rule("RANGE", "xlCellValue", "xlGreater", '"70"'), False),
           "abs": (rule("RANGE", "xlExpression", first='"=B5>$J$1"'), False),
           "top": ('Set f = ws.Range("RANGE").FormatConditions.AddTop10\nf.Interior.Color = RGB(255, 0, 0)\n', False),
           "avg": ('Set f = ws.Range("RANGE").FormatConditions.AddAboveAverage\nf.Interior.Color = RGB(255, 0, 0)\n',
                   False),
           "uniq": ('Set f = ws.Range("RANGE").FormatConditions.AddUniqueValues\nf.Interior.Color = RGB(255, 0, 0)\n',
                    False),
           "text": ('Set f = ws.Range("RANGE").FormatConditions.Add(xlTextString, String:="8", '
                    "TextOperator:=xlContains)\nf.Interior.Color = RGB(255, 0, 0)\n", False),
           "scale": ('Set f = ws.Range("RANGE").FormatConditions.AddColorScale(3)\n', False),
           "bar": ('Set f = ws.Range("RANGE").FormatConditions.AddDatabar\n', False),
           "icons": ('Set f = ws.Range("RANGE").FormatConditions.AddIconSetCondition\n', False),
           "stars": ('Set f = ws.Range("RANGE").FormatConditions.AddIconSetCondition\n'
                     "f.IconSet = ws.Parent.IconSets(xl3Stars)\n", False),
       }.items()
       for edit, (where, action) in {
           "ins": ("B5:B10", "ws.Rows(7).Insert"), "del": ("B5:B10", "ws.Rows(7).Delete"),
           "up": ("B5:B10", "ws.Rows(2).Delete"), "below": ("B5:B10", "ws.Rows(11).Insert"),
           "part": ("B5:D10", 'ws.Range("C7").Insert Shift:=xlDown'),
       }.items() if not block},
    # Whether a rule made anew takes the highest priority in the workbook, plus one.
    "high_priority": (rule("B5:B10", "xlCellValue", "xlGreater", '"70"') + "f.Priority = 50\n", ""),
    "after_high": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\nf.Interior.Color = RGB(255, 0, 0)\n',
                   "ws.Rows(7).Delete"),
    "two_kinds": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\nf.Interior.Color = RGB(255, 0, 0)\n'
                  'Set g = ws.Range("B5:B10").FormatConditions.Add(xlCellValue, xlGreater, "70")\n'
                  "g.Interior.Color = RGB(0, 0, 255)\n", "ws.Rows(7).Delete"),
    # Where a shift splits a block, and where it moves it whole.
    "p_above_ins": (BLOCK, 'ws.Range("C2").Insert Shift:=xlDown'),
    "p_above_del": (BLOCK, 'ws.Range("C2").Delete Shift:=xlUp'),
    "p_top_ins": (BLOCK, 'ws.Range("C5").Insert Shift:=xlDown'),
    "p_top_del": (BLOCK, 'ws.Range("C5").Delete Shift:=xlUp'),
    "p_full_ins": (BLOCK, 'ws.Range("C5:C10").Insert Shift:=xlDown'),
    "p_edge_full_del": (BLOCK, 'ws.Range("B5:B10").Delete Shift:=xlUp'),
    "p_edge_full_ins": (BLOCK, 'ws.Range("B5:B10").Insert Shift:=xlDown'),
    "p_edge_ins": (BLOCK, 'ws.Range("B7").Insert Shift:=xlDown'),
    "p_edge_del": (BLOCK, 'ws.Range("D7").Delete Shift:=xlUp'),
    "p_two_cols": (BLOCK, 'ws.Range("C7:D7").Insert Shift:=xlDown'),
    "p_below_area": (BLOCK, 'ws.Range("C12").Insert Shift:=xlDown'),
    "p_last_row": (BLOCK, 'ws.Range("C10").Insert Shift:=xlDown'),
    "p_after_last": (BLOCK, 'ws.Range("C11").Insert Shift:=xlDown'),
    "p_many": (BLOCK, 'ws.Range("C7:C8").Insert Shift:=xlDown'),
    "p_multi": (rule("B5:D10,F5:F10", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("C7:F7").Insert Shift:=xlDown'),
    "p_two_areas": (rule("B5:D6,B8:D10", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("C7").Insert Shift:=xlDown'),
    "h_before_ins": (WIDE, 'ws.Range("A2").Insert Shift:=xlToRight'),
    "h_left_ins": (WIDE, 'ws.Range("B2").Insert Shift:=xlToRight'),
    "h_full_del": (WIDE, 'ws.Range("B2:E2").Delete Shift:=xlToLeft'),
    "h_bottom_ins": (WIDE, 'ws.Range("C3").Insert Shift:=xlToRight'),
    "h_middle_ins": (rule("B2:E4", "xlCellValue", "xlGreater", '"22"'), 'ws.Range("C3").Insert Shift:=xlToRight'),
    "h_middle_del": (rule("B2:E4", "xlCellValue", "xlGreater", '"22"'), 'ws.Range("C3").Delete Shift:=xlToLeft'),
    # What inserted lines take from the lines they copy their formats from.
    "inherit_block": (BLOCK, "ws.Rows(11).Insert"),
    "inherit_part": (rule("B5:B10,D8:D10", "xlCellValue", "xlGreater", '"70"'), "ws.Rows(11).Insert"),
    "inherit_columns": (WIDE, 'ws.Columns("F").Insert'),
    "inherit_two_rules": (ROWS + rule("C9:C10", "xlCellValue", "xlGreater", '"70"'), "ws.Rows(11).Insert"),
    "inherit_many": (ROWS, 'ws.Rows("11:13").Insert'),
    # Which deletes make a statistical or drawn rule anew: at its edges, in part of its band, or a whole column.
    "r_top_first": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\n', "ws.Rows(5).Delete"),
    "r_top_last": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\n', "ws.Rows(10).Delete"),
    "r_top_over": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\n', 'ws.Rows("3:6").Delete'),
    "r_top_part": ('Set f = ws.Range("B5:D10").FormatConditions.AddTop10\n', 'ws.Range("C7").Delete Shift:=xlUp'),
    "r_top_column": ('Set f = ws.Range("B5:D10").FormatConditions.AddTop10\n', 'ws.Columns("C").Delete'),
    "r_top_band": ('Set f = ws.Range("B5:D10").FormatConditions.AddTop10\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "r_bar_first": ('Set f = ws.Range("B5:B10").FormatConditions.AddDatabar\n', "ws.Rows(5).Delete"),
    "r_scale_part": ('Set f = ws.Range("B5:D10").FormatConditions.AddColorScale(3)\n',
                     'ws.Range("C7").Delete Shift:=xlUp'),
    "r_two_groups": ('Set f = ws.Range("B5:B10").FormatConditions.AddTop10\n'
                     'Set g = ws.Range("C5:C10").FormatConditions.Add(xlCellValue, xlGreater, "70")\n'
                     'Set h = ws.Range("D5:D10").FormatConditions.AddTop10\n', "ws.Rows(7).Delete"),
    # The order Excel reads areas back in.
    "order_tall_first": (rule("A1:A10,B2:B3", "xlCellValue", "xlGreater", '"0"'), ""),
    "order_wide_first": (rule("C1:C2,A1:A10", "xlCellValue", "xlGreater", '"0"'), ""),
    "order_same_top": (rule("D3:D9,B3:B4,C3:C6", "xlCellValue", "xlGreater", '"0"'), ""),
    "order_rows": (rule("A8:C8,A2:C2,A5:C5", "xlCellValue", "xlGreater", '"0"'), ""),
    # Rules a resave edits once Excel has read them from the file.
    "loaded_rows": (ROWS + rule("C5:C10", "xlExpression", first='"=$A5>65"') +
                    'ws.Range("D5:D10").FormatConditions.AddDatabar\n'
                    'Set h = ws.Range("E5:E10").FormatConditions.AddIconSetCondition\n'
                    "h.IconSet = ws.Parent.IconSets(xl3Stars)\n", ""),
    "loaded_columns": (COLUMNS + 'ws.Range("B3:E3").FormatConditions.AddDatabar\n', ""),
    "loaded_shift": (BLOCK, ""),
    "loaded_delete_all": (ROWS + 'ws.Range("B5:B10").FormatConditions.AddDatabar\n', ""),
    # Whole columns and rows, through edits along them.
    "full_column_rows": (rule("B:B", "xlCellValue", "xlGreater", '"70"'), "ws.Rows(7).Delete"),
    "full_column_insert": (rule("B:B", "xlCellValue", "xlGreater", '"70"'), "ws.Rows(7).Insert"),
    "full_row_columns": (rule("2:2", "xlCellValue", "xlGreater", '"22"'), 'ws.Columns("C").Delete'),
    "full_column_up": (rule("B:B", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("B7").Delete Shift:=xlUp'),
    "full_column_down": (rule("B:B", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("B7").Insert Shift:=xlDown'),
    "full_column_part": (rule("B:C", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("B7").Delete Shift:=xlUp'),
    # Cells inserted just after a range, and CopyOrigin on cell shifts.
    "shift_after_last": (ROWS, 'ws.Range("B11").Insert Shift:=xlDown'),
    "shift_origin_after": (ROWS, 'ws.Range("B11").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow'),
    "shift_origin_top": (BLOCK, 'ws.Range("C5").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow'),
    "shift_origin_inside": (ROWS, 'ws.Range("B5").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow'),
    "h_origin_left": (WIDE, 'ws.Range("B2").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow'),
    # A drawn rule's point formula made #REF!.
    "scale_point_lost": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddColorScale(2)\n'
                         "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                         'f.ColorScaleCriteria(1).Value = "=$D$1"\n', "ws.Rows(1).Delete"),
    "bar_point_lost": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddDatabar\n'
                       'f.MinPoint.Modify xlConditionValueFormula, "=$D$1"\n', "ws.Rows(1).Delete"),
    "icons_point_lost": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddIconSetCondition\n'
                         "f.IconCriteria(2).Type = xlConditionValueFormula\n"
                         'f.IconCriteria(2).Value = "=$D$1"\n', "ws.Rows(1).Delete"),
    # An absolute reference through a cell shift, and a statistical rule losing a whole area.
    "shift_abs_ref": (rule("B5:B10", "xlExpression", first='"=B5>$J$3"'), 'ws.Range("J2").Insert Shift:=xlDown'),
    "shift_abs_lost": (rule("B5:B10", "xlExpression", first='"=B5>$J$3"'), 'ws.Range("J3").Delete Shift:=xlUp'),
    "r_top_area": ('Set f = ws.Range("B2:B3,B8:B9").FormatConditions.AddTop10\n', 'ws.Rows("2:3").Delete'),
    "r_top_area_part": ('Set f = ws.Range("B2:B3,B8:B9").FormatConditions.AddTop10\n', "ws.Rows(3).Delete"),
    "r_avg_area": ('Set f = ws.Range("B2:B3,B8:B9").FormatConditions.AddAboveAverage\n', 'ws.Rows("2:3").Delete'),
    "r_top_full_column": ('Set f = ws.Range("B:B").FormatConditions.AddTop10\n', "ws.Rows(7).Delete"),
    "r_top_two_part": ('Set f = ws.Range("B2:B4,B8:B9").FormatConditions.AddTop10\n', "ws.Rows(3).Delete"),
    "r_top_second_part": ('Set f = ws.Range("B2:B3,B8:B10").FormatConditions.AddTop10\n', "ws.Rows(9).Delete"),
    "r_top_small": ('Set f = ws.Range("B5:B6").FormatConditions.AddTop10\n', "ws.Rows(6).Delete"),
    "r_top_block_two": ('Set f = ws.Range("B5:D10,F5:F10").FormatConditions.AddTop10\n', 'ws.Columns("C").Delete'),
    "r_scale_two": ('Set f = ws.Range("B2:B3,B8:B9").FormatConditions.AddColorScale(3)\n', "ws.Rows(3).Delete"),
    # Whether a rule a delete would make anew stays when its first area moves whole.
    "u_ref_strip": (rule("B5:D10", "xlExpression", first='"=B5>$C$7"'), 'ws.Range("C7").Delete Shift:=xlUp'),
    "u_ref_strip_moving": (rule("B5:D10", "xlExpression", first='"=B5>$C$2"'), 'ws.Range("C2").Delete Shift:=xlUp'),
    "u_ref_first_moved": (rule("B8:B9,B2:B4", "xlExpression", first='"=$J$3>0"'), "ws.Rows(3).Delete"),
    "u_ref_first_still": (rule("B2:B3,B8:B9", "xlExpression", first='"=$J$5>0"'), "ws.Rows(5).Delete"),
    "u_top_first_moved": ('Set f = ws.Range("B8:B9,B2:B4").FormatConditions.AddTop10\n', "ws.Rows(3).Delete"),
    "u_top_untouched_cut_moved": ('Set f = ws.Range("B2:B3,B5:B6,B8:B9").FormatConditions.AddTop10\n',
                                  "ws.Rows(6).Delete"),
    "u_ref_two_moved": (rule("B8:B9,B2:B3,B11:B12", "xlExpression", first='"=$J$5>0"'), "ws.Rows(5).Delete"),
    "u_shift_moved_first": (rule("B2:B3,B8:B9", "xlExpression", first='"=$B$5>0"'), 'ws.Range("B5").Delete Shift:=xlUp'),
    "r_top_removed_last": ('Set f = ws.Range("B2:B3,B8:B9").FormatConditions.AddTop10\n', 'ws.Rows("8:9").Delete'),
    # Areas a delete brings together, and areas a cell shift cuts.
    "merge_order": (rule("B5:B6,B2:B3,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(7).Delete"),
    "merge_columns": (rule("B2:C2,E2:F2", "xlCellValue", "xlGreater", '"20"'), 'ws.Columns("D").Delete'),
    "merge_cut": (rule("B2:B4,B6:B7", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(5).Delete"),
    "shift_multi_del": (rule("B5:D10,F5:F10", "xlCellValue", "xlGreater", '"70"'), 'ws.Range("C7:F7").Delete Shift:=xlUp'),
    "shift_multi_del_rev": (rule("F5:F10,B5:D10", "xlCellValue", "xlGreater", '"70"'),
                            'ws.Range("C7:F7").Delete Shift:=xlUp'),
    # The order of a range's areas once a delete cuts or an insert stretches some of them.
    "multi_shrink": (rule("B2:B4,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(3).Delete"),
    "multi_shrink_second": (rule("B2:B3,B8:B10", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(9).Delete"),
    "multi_three": (rule("B2:B3,B5:B6,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(5).Delete"),
    "multi_shrink_cols": (rule("B2:C2,E2:F2", "xlCellValue", "xlGreater", '"20"'), 'ws.Columns("B").Delete'),
    "multi_grow": (rule("B2:B3,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(3).Insert"),
    "multi_grow_after": (rule("B2:B3,B8:B9", "xlCellValue", "xlGreater", '"20"'), "ws.Rows(4).Insert"),
    # When a reference made #REF! makes its rule anew: whether the rule moves, the reference's kind, the edit's.
    "lost_rel_moving": (rule("D5:D10", "xlExpression", first='"=C5>70"'), 'ws.Columns("C").Delete'),
    "lost_abs_still_row": (rule("B5:B10", "xlExpression", first='"=B5>$J$20"'), "ws.Rows(20).Delete"),
    "lost_abs_still_col": (rule("B5:B10", "xlExpression", first='"=B5>$J$3"'), 'ws.Columns("J").Delete'),
    "lost_abs_moving_shift": (rule("B5:B10", "xlExpression", first='"=B5>$B$3"'),
                              'ws.Range("B3").Delete Shift:=xlUp'),
    "lost_inside": (rule("B5:B10", "xlExpression", first='"=B5>$J$7"'), "ws.Rows(7).Delete"),
    "lost_cell_still": (rule("B5:B10", "xlCellValue", "xlGreater", '"=$J$20"'), "ws.Rows(20).Delete"),
    "lost_scale_still": ('ws.Range("D20").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddColorScale(2)\n'
                         "f.ColorScaleCriteria(1).Type = xlConditionValueFormula\n"
                         'f.ColorScaleCriteria(1).Value = "=$D$20"\n', "ws.Rows(20).Delete"),
    "lost_stars_moving": ('ws.Range("D1").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddIconSetCondition\n'
                          "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.IconCriteria(2).Type = xlConditionValueFormula\n"
                          'f.IconCriteria(2).Value = "=$D$1"\n', "ws.Rows(1).Delete"),
    "lost_stars_still": ('ws.Range("D20").Value = 60\nSet f = ws.Range("B5:B10").FormatConditions.AddIconSetCondition\n'
                         "f.IconSet = ws.Parent.IconSets(xl3Stars)\nf.IconCriteria(2).Type = xlConditionValueFormula\n"
                         'f.IconCriteria(2).Value = "=$D$20"\n', "ws.Rows(20).Delete"),
    # A delete taking the middle of a block's band, for each kind of rule.
    "band_avg": ('ws.Range("B5:D10").FormatConditions.AddAboveAverage\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_scale": ('ws.Range("B5:D10").FormatConditions.AddColorScale 3\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_bar": ('ws.Range("B5:D10").FormatConditions.AddDatabar\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_icons": ('ws.Range("B5:D10").FormatConditions.AddIconSetCondition\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_text": ('Set f = ws.Range("B5:D10").FormatConditions.Add(xlTextString, String:="8", '
                  "TextOperator:=xlContains)\nf.Interior.Color = RGB(255, 0, 0)\n", 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_expr": (rule("B5:D10", "xlExpression", first='"=B5>70"'), 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_unique": ('ws.Range("B5:D10").FormatConditions.AddUniqueValues\n', 'ws.Range("C5:C10").Delete Shift:=xlUp'),
    "band_top_left": ('ws.Range("B5:D10").FormatConditions.AddTop10\n', 'ws.Range("C5:C10").Delete Shift:=xlToLeft'),
    # The order Excel reads back the areas of rules drawn from their range.
    "order_scale": ('ws.Range("B5:B6,B2:B3").FormatConditions.AddColorScale 3\n', ""),
    "order_bar": ('ws.Range("B5:B6,B2:B3").FormatConditions.AddDatabar\n', ""),
    "order_icons": ('ws.Range("B5:B6,B2:B3").FormatConditions.AddIconSetCondition\n', ""),
    "order_text": ('Set f = ws.Range("B5:B6,B2:B3").FormatConditions.Add(xlTextString, String:="8", '
                   "TextOperator:=xlContains)\nf.Interior.Color = RGB(255, 0, 0)\n", ""),
    "order_expr": (rule("B5:B6,B2:B3", "xlExpression", first='"=B2>0"'), ""),
    "order_cell": (rule("B5:B6,B2:B3", "xlCellValue", "xlGreater", '"0"'), ""),
    "order_avg": ('ws.Range("B5:B6,B2:B3").FormatConditions.AddAboveAverage\n', ""),
    # Which rules Excel joins as it reads a file: alike but for their ranges, beside each other or not.
    "same_apart": (rule("B2:B3", "xlCellValue", "xlGreater", '"22"') + rule("B8:B9", "xlCellValue", "xlGreater", '"22"'),
                   ""),
    "same_side": (rule("B2:B3", "xlCellValue", "xlGreater", '"22"') + rule("C2:C3", "xlCellValue", "xlGreater", '"22"'),
                  ""),
    "same_touch": (rule("B2:B3", "xlCellValue", "xlGreater", '"22"') + rule("C3:C4", "xlCellValue", "xlGreater", '"22"'),
                   ""),
    "same_three": (rule("B2", "xlCellValue", "xlGreater", '"22"') + rule("B3", "xlCellValue", "xlGreater", '"22"') +
                   rule("B4", "xlCellValue", "xlGreater", '"22"'), ""),
    "same_overlap": (rule("B2:B4", "xlCellValue", "xlGreater", '"22"') + rule("B3:B5", "xlCellValue", "xlGreater", '"22"'),
                     ""),
    "same_other_fill": (rule("B2:B3", "xlCellValue", "xlGreater", '"22"') + 'Set g = ws.Range("B4:B5").FormatConditions'
                        '.Add(xlCellValue, xlGreater, "22")\ng.Interior.Color = RGB(0, 0, 255)\n', ""),
    "same_between": (rule("B2:B3", "xlCellValue", "xlGreater", '"22"') + rule("D2", "xlCellValue", "xlLess", '"5"') +
                     rule("B4:B5", "xlCellValue", "xlGreater", '"22"'), ""),
    "same_relative": (rule("B2:B3", "xlExpression", first='"=B2>22"') + rule("B4:B5", "xlExpression", first='"=B4>22"'),
                      ""),
    "same_top": ('ws.Range("B2:B3").FormatConditions.AddTop10\nws.Range("B4:B5").FormatConditions.AddTop10\n', ""),
    "same_reversed": (rule("B4:B5", "xlCellValue", "xlGreater", '"22"') + rule("B2:B3", "xlCellValue", "xlGreater", '"22"'),
                      ""),
    # A rule whose own formula names another sheet.
    "cell_other_insert": (helper_sheet() + rule("B5:B10", "xlCellValue", "xlGreater", '"=cell_other_insert_r!$A$2"'),
                          "w2.Rows(1).Insert"),
    "cell_other_delete": (helper_sheet() + rule("B5:B10", "xlCellValue", "xlGreater", '"=cell_other_delete_r!$A$2"'),
                          "w2.Rows(2).Delete"),
}

#: Rules beside one a delete makes anew, each layout followed by a probe whose top 10 a column delete makes anew.
TOP_ROWS = 'ws.Range("B5:B10").FormatConditions.AddTop10\n'
CUT_C = rule("C5:C10", "xlCellValue", "xlGreater", '"70"')
for _name, _setup in {
    "rx_cut_before": CUT_C + TOP_ROWS,
    "rx_cut_after": TOP_ROWS + CUT_C,
    "rx_moved_before": rule("B12:E12", "xlCellValue", "xlGreater", '"70"') + TOP_ROWS,
    "rx_moved_after": TOP_ROWS + rule("B12:E12", "xlCellValue", "xlGreater", '"70"'),
    "rx_untouched": rule("B2:E2", "xlCellValue", "xlGreater", '"22"') + TOP_ROWS,
    "rx_two_cut_before": CUT_C + rule("D5:D10", "xlCellValue", "xlGreater", '"70"') + TOP_ROWS,
    "rx_cut_bar_before": 'ws.Range("C5:C10").FormatConditions.AddDatabar\n' + TOP_ROWS,
    "rx_cut_top_before": 'ws.Range("C5:C10").FormatConditions.AddTop10\n' + TOP_ROWS,
}.items():
    LAYOUTS[_name] = (_setup, "ws.Rows(7).Delete")
    LAYOUTS[_name + "_after"] = ('Set h = ws.Range("H5:I10").FormatConditions.AddTop10\n', 'ws.Columns("I").Delete')

#: Whether an edit uses up a priority: each runs, then a top 10 made over H5:I10 is made anew by a column delete.
LAYOUTS.update({f"n_{name}": (LAYOUTS[name][0] + LAYOUTS[name][1] + "\n" +
                              'Set h = ws.Range("H5:I10").FormatConditions.AddTop10\n', 'ws.Columns("I").Delete')
                for name in ("del_all", "multi_area", "multi_area_delete", "two_groups", "two_rules_one_range",
                             "groups_meet", "same_rule_meet", "shift_up_all", "shift_left_all", "shift_expr_part")})
LAYOUTS["n_none"] = ('Set h = ws.Range("H5:I10").FormatConditions.AddTop10\n', 'ws.Columns("I").Delete')

#: sheet -> the edit made to it once Excel has opened its own file, before it saves the file again.
RESAVE: dict[str, str] = {
    "k_top_ins": "ws.Rows(9).Delete",
    "k_cell_ins": "ws.Rows(9).Delete",
    "k_avg_ins": "ws.Rows(9).Delete",
    "loaded_rows": "ws.Rows(7).Insert",
    "loaded_columns": 'ws.Columns("C").Delete',
    "loaded_shift": 'ws.Range("C7").Insert Shift:=xlDown',
    "loaded_delete_all": 'ws.Rows("5:10").Delete',
    "high_priority": 'ws.Range("B5:B10").FormatConditions.AddTop10\nws.Rows(7).Delete',
}
#: Rules a resave moves once Excel has read them from the file, each sheet with a top 10 over H2:I3 that a column
#: delete then makes anew, to show whether moving them used up priorities.
PROBE = 'Set h = ws.Range("H2:I3").FormatConditions.AddTop10\n'
BAR = 'ws.Range("D5:D10").FormatConditions.AddDatabar\n'
LOADED_MOVES: dict[str, tuple[str, str]] = {
    "lx_none": ("", "ws.Rows(7).Insert"),
    "lx_cell": (rule("B5:B10", "xlCellValue", "xlGreater", '"70"'), "ws.Rows(7).Insert"),
    "lx_bar": (BAR, "ws.Rows(7).Insert"),
    "lx_stars": ('Set g = ws.Range("E5:E10").FormatConditions.AddIconSetCondition\n'
                 "g.IconSet = ws.Parent.IconSets(xl3Stars)\n", "ws.Rows(7).Insert"),
    "lx_two_bars": (BAR + BAR, "ws.Rows(7).Insert"),
    "lx_bar_scale": (BAR + 'ws.Range("F5:F10").FormatConditions.AddColorScale 3\n', "ws.Rows(7).Insert"),
    "lx_top": ('ws.Range("B5:B10").FormatConditions.AddTop10\n', "ws.Rows(7).Insert"),
    "lx_icons": ('ws.Range("E5:E10").FormatConditions.AddIconSetCondition\n', "ws.Rows(7).Insert"),
    "lx_bar_shift": (BAR, 'ws.Range("D2").Insert Shift:=xlDown'),
    "lx_bar_delete": (BAR, "ws.Rows(4).Delete"),
    "lx_bar_still": (BAR, "ws.Rows(20).Insert"),
    "lx_expr": (rule("C5:C10", "xlExpression", first='"=$A5>65"'), "ws.Rows(7).Insert"),
    "lx_rows_like": (ROWS + rule("C5:C10", "xlExpression", first='"=$A5>65"') + BAR +
                     'Set g = ws.Range("E5:E10").FormatConditions.AddIconSetCondition\n'
                     "g.IconSet = ws.Parent.IconSets(xl3Stars)\n", "ws.Rows(7).Insert"),
}
LAYOUTS.update({name: (setup + PROBE, "") for name, (setup, _) in LOADED_MOVES.items()})
RESAVE.update({name: edit + '\nws.Columns("I").Delete' for name, (_, edit) in LOADED_MOVES.items()})
#: Rules read from the file that a resave cuts or makes anew, each followed by a probe on a sheet of its own showing
#: the numbers it used.
CELLS_ROW2 = rule("B2:E2", "xlCellValue", "xlGreater", '"22"')
LOADED_COUNTED: dict[str, tuple[str, str]] = {
    "lx_bar_remade": ('ws.Range("B3:E3").FormatConditions.AddDatabar\n', 'ws.Columns("C").Delete'),
    "lx_scale_remade": ('ws.Range("B3:E3").FormatConditions.AddColorScale 3\n', 'ws.Columns("C").Delete'),
    "lx_stars_remade": ('Set g = ws.Range("B3:E3").FormatConditions.AddIconSetCondition\n'
                        "g.IconSet = ws.Parent.IconSets(xl3Stars)\n", 'ws.Columns("C").Delete'),
    "lx_top_remade": ('ws.Range("B3:E3").FormatConditions.AddTop10\n', 'ws.Columns("C").Delete'),
    "lx_cell_cols": (CELLS_ROW2, 'ws.Columns("C").Delete'),
    "lx_cell_rows": (ROWS, "ws.Rows(7).Delete"),
    "lx_cell_cols_bar": (CELLS_ROW2 + 'ws.Range("B3:E3").FormatConditions.AddDatabar\n', 'ws.Columns("C").Delete'),
    "lx_cell_cols_top": (CELLS_ROW2 + 'ws.Range("B3:E3").FormatConditions.AddTop10\n', 'ws.Columns("C").Delete'),
    "lx_two_cut": (CELLS_ROW2 + rule("B3:E3", "xlCellValue", "xlGreater", '"22"'), 'ws.Columns("C").Delete'),
}
for _name, (_setup, _edit) in LOADED_COUNTED.items():
    LAYOUTS[_name] = (_setup, "")
    LAYOUTS[_name + "_after"] = (PROBE, "")
    RESAVE[_name] = _edit
    RESAVE[_name + "_after"] = 'ws.Columns("I").Delete'


def case_code(index: int, setup: str, edit: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim failed As String, r As Long, c As Long, f As Object, g As Object, h As Object, w2 As Object",
             "On Error Resume Next", "Err.Clear", *(GRID + setup).splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "Err.Clear", *edit.splitlines(),
             'If Err.Number <> 0 Then failed = failed & "E" & Err.Number & ";"', "On Error GoTo 0",
             "ws.Activate", f'Case{index} = failed & Dump(ws)', "End Function"]
    return "\n".join(lines) + "\n"


def resave_code(index: int, edit: str) -> str:
    lines = [f"Private Function Resave{index}(ws As Object) As String", "Dim failed As String",
             "On Error Resume Next", "ws.Activate", "Err.Clear", *edit.splitlines(),
             'If Err.Number <> 0 Then failed = "E" & Err.Number & ";"', "On Error GoTo 0",
             f"Resave{index} = failed & Dump(ws)", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "conditional_edits.xlsx"
    resaved = OUT / "conditional_edits_resaved.xlsx"
    for one in (path, resaved):
        if one.exists():
            one.unlink()
    build = [HELPER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, (setup, edit)) in enumerate(LAYOUTS.items()):
        assert len(name) + 2 <= 31, name
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup, edit))
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "<<PHASE>>"',
              f'Set wb = Workbooks.Open("{path}")', "For Each ws In wb.Worksheets",
              '    out = out & ws.Name & "^" & Dump(ws) & "|"', "Next", 'out = out & "<<PHASE>>"']
    for index, (name, edit) in enumerate(RESAVE.items()):
        build += [f'Set ws = wb.Worksheets("{name}")', f'out = out & ws.Name & "^" & Resave{index}(ws) & "|"']
        bodies.append(resave_code(index, edit))
    build += [f'wb.SaveAs "{resaved}", 51', "wb.Close False", "Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    strip_save_path(resaved)
    before, after, changed = str(result.value).split("<<PHASE>>")
    # An answer holding the separator would shift every answer after it onto the wrong layout.
    assert before.count("|") == len(LAYOUTS), "an answer holds the layout separator"
    answers = before.split("|")[: len(LAYOUTS)]
    reopened = dict(one.split("^", 1) for one in after.split("|") if one)
    resave_answers = dict(one.split("^", 1) for one in changed.split("|") if one)
    with zipfile.ZipFile(path) as package:
        parts = file_parts(package)
    with zipfile.ZipFile(resaved) as package:
        resaved_parts = file_parts(package)
    record = {"helper": HELPER, "grid": GRID,
              "layouts": [{"name": name, "setup": setup, "edit": edit, "answers": answer,
                           "reopened": reopened.get(name, ""), "file": parts.get(name, ""),
                           "helper_file": parts.get(name + "_r", "")}
                          for (name, (setup, edit)), answer in zip(LAYOUTS.items(), answers, strict=True)],
              "resaved": [{"name": name, "edit": edit, "answers": resave_answers.get(name, ""),
                           "file": resaved_parts.get(name, "")} for name, edit in RESAVE.items()]}
    (OUT / "conditional_edits.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']}: {one['answers'][:700]}")
        if one["reopened"] not in one["answers"]:
            print(f"    reopened: {one['reopened'][:700]}")
        print(f"    {one['file'][:900]}")
    for one in record["resaved"]:
        print(f"resaved {one['name']}: {one['answers'][:700]}")
        print(f"    {one['file'][:1500]}")


if __name__ == "__main__":
    main()
