"""Sheet protection from VBA, as Excel enforces it.

Each case runs on a sheet of its own, fresh from Worksheets.Add, so every
cell starts locked: setup lines, then the action under On Error, then a
read-out of the error the action raised and of the state it left. The
cases cover what Protect sets, which edits a protected sheet refuses and
with which error, the password rules for Unprotect, UserInterfaceOnly,
Contents:=False and the Allow options.

    python scripts/measure_protection.py

writes tests/fixtures/protection.json, which tests/test_excel_protection.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "protection.json"

#: What every case reads back after its action: the protection flags and A1.
STATE = ('ws.ProtectContents & "," & ws.ProtectDrawingObjects & "," & ws.ProtectScenarios & "," & '
         'ws.ProtectionMode & "," & TypeName(ws.Range("A1").Value) & ":" & CStr(ws.Range("A1").Value)')

#: (name, setup, action): the action's error, if any, is recorded, then the state.
CASES: list[tuple[str, str, str]] = [
    ("protect", "", "ws.Protect"),
    ("protect_flags", "", 'ws.Protect: out = out & ws.Protection.AllowFormattingCells & "," & '
                          'ws.Protection.AllowInsertingRows & "," & ws.Protection.AllowDeletingRows & "," & '
                          'ws.Protection.AllowSorting & "," & ws.Protection.AllowFiltering & ";"'),
    ("write_locked", "ws.Protect", 'ws.Range("A1").Value = 1'),
    ("write_unlocked", 'ws.Range("A1").Locked = False: ws.Protect', 'ws.Range("A1").Value = 1'),
    ("write_mixed", 'ws.Range("A1").Locked = False: ws.Protect', 'ws.Range("A1:A2").Value = 1'),
    ("write_formula", "ws.Protect", 'ws.Range("A1").Formula = "=1+1"'),
    ("write_same_value", 'ws.Range("A1").Value = 5: ws.Protect', 'ws.Range("A1").Value = 5'),
    ("clear_contents", 'ws.Range("A1").Value = 5: ws.Protect', 'ws.Range("A1").ClearContents'),
    ("clear", 'ws.Range("A1").Value = 5: ws.Protect', 'ws.Range("A1").Clear'),
    ("clear_unlocked", 'ws.Range("A1").Value = 5: ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1").ClearContents'),
    ("user_interface_only", "ws.Protect UserInterfaceOnly:=True", 'ws.Range("A1").Value = 1'),
    ("contents_false", "ws.Protect Contents:=False", 'ws.Range("A1").Value = 1'),
    ("drawing_false", "", "ws.Protect DrawingObjects:=False"),
    ("scenarios_false", "", "ws.Protect Scenarios:=False"),
    ("set_locked", "ws.Protect", 'ws.Range("A1").Locked = False'),
    ("read_locked", "ws.Protect", 'out = out & ws.Range("A1").Locked & ";"'),
    ("format_locked", "ws.Protect", 'ws.Range("A1").Interior.Color = vbRed'),
    ("format_allowed", "ws.Protect AllowFormattingCells:=True", 'ws.Range("A1").Interior.Color = vbRed'),
    ("number_format", "ws.Protect", 'ws.Range("A1").NumberFormat = "0.00"'),
    ("number_format_allowed", "ws.Protect AllowFormattingCells:=True", 'ws.Range("A1").NumberFormat = "0.00"'),
    ("font_bold", "ws.Protect", 'ws.Range("A1").Font.Bold = True'),
    ("format_unlocked", 'ws.Range("A1").Locked = False: ws.Protect', 'ws.Range("A1").Interior.Color = vbRed'),
    ("column_width", "ws.Protect", "ws.Columns(1).ColumnWidth = 20"),
    ("column_width_allowed", "ws.Protect AllowFormattingColumns:=True", "ws.Columns(1).ColumnWidth = 20"),
    ("row_height", "ws.Protect", "ws.Rows(1).RowHeight = 30"),
    ("row_height_allowed", "ws.Protect AllowFormattingRows:=True", "ws.Rows(1).RowHeight = 30"),
    ("insert_row", "ws.Protect", "ws.Rows(2).Insert"),
    ("insert_row_allowed", "ws.Protect AllowInsertingRows:=True", "ws.Rows(2).Insert"),
    ("delete_row", "ws.Protect", "ws.Rows(2).Delete"),
    ("delete_row_allowed", "ws.Protect AllowDeletingRows:=True", "ws.Rows(2).Delete"),
    ("delete_row_allowed_unlocked", "ws.Rows(2).Locked = False: ws.Protect AllowDeletingRows:=True",
     "ws.Rows(2).Delete"),
    ("insert_column", "ws.Protect", "ws.Columns(2).Insert"),
    ("insert_column_allowed", "ws.Protect AllowInsertingColumns:=True", "ws.Columns(2).Insert"),
    ("delete_column_allowed", "ws.Protect AllowDeletingColumns:=True", "ws.Columns(2).Delete"),
    ("hide_row", "ws.Protect", "ws.Rows(3).Hidden = True"),
    ("hide_row_allowed", "ws.Protect AllowFormattingRows:=True", "ws.Rows(3).Hidden = True"),
    ("merge", "ws.Protect", 'ws.Range("A1:B2").Merge'),
    ("copy_into_locked", 'ws.Range("A1").Value = 5: ws.Protect', 'ws.Range("A1").Copy ws.Range("B1")'),
    ("copy_into_unlocked", 'ws.Range("A1").Value = 5: ws.Range("B1").Locked = False: ws.Protect',
     'ws.Range("A1").Copy ws.Range("B1"): out = out & CStr(ws.Range("B1").Value) & ";"'),
    ("read_value", 'ws.Range("A1").Value = 5: ws.Protect', 'out = out & CStr(ws.Range("A1").Value) & ";"'),
    ("rename", "ws.Protect", 'ws.Name = "Renamed" & ws.Index'),
    ("password_unprotect", 'ws.Protect "pw"', 'ws.Unprotect "pw"'),
    ("password_wrong", 'ws.Protect "pw"', 'ws.Unprotect "bad"'),
    ("password_case", 'ws.Protect "pw"', 'ws.Unprotect "PW"'),
    ("password_empty", 'ws.Protect ""', "ws.Unprotect"),
    ("no_password_given_one", "ws.Protect", 'ws.Unprotect "anything"'),
    ("unprotect_unprotected", "", "ws.Unprotect"),
    ("protect_twice", "ws.Protect", "ws.Protect"),
    ("protect_twice_password", 'ws.Protect "pw"', 'ws.Protect "other"'),
    ("protect_twice_same_password", 'ws.Protect "pw"', 'ws.Protect "pw", Contents:=False'),
    ("protect_again_changes", "ws.Protect", "ws.Protect Contents:=False"),
    ("unlocked_after_unprotect", 'ws.Protect "pw": ws.Unprotect "pw"', 'ws.Range("A1").Value = 1'),
    ("select_locked", "ws.Protect", 'ws.Activate: ws.Range("A1").Select'),
    ("enable_selection", "ws.Protect", 'out = out & ws.EnableSelection & ";"'),
    # Which cells a refused write still reaches, in which order.
    ("write_mixed_reverse", 'ws.Range("A2").Locked = False: ws.Protect',
     'ws.Range("A1:A2").Value = 1: out = out & CStr(ws.Range("A2").Value) & ";"'),
    ("write_mixed_across", 'ws.Range("A1,C1").Locked = False: ws.Protect',
     'ws.Range("A1:C1").Value = 7: out = out & CStr(ws.Range("B1").Value) & "," & CStr(ws.Range("C1").Value) & ";"'),
    ("write_array_mixed", 'ws.Range("A1").Locked = False: ws.Range("B1").Locked = False: ws.Protect',
     'ws.Range("A1:C1").Value = Array(1, 2, 3): out = out & CStr(ws.Range("B1").Value) & ";"'),
    # A second Protect with another password, then which password opens it.
    ("protect_other_password_options", 'ws.Protect "pw"', 'ws.Protect "other", Contents:=False'),
    ("protect_other_password_keeps", 'ws.Protect "pw": ws.Protect "other"', 'ws.Unprotect "pw"'),
    ("protect_other_password_new", 'ws.Protect "pw": ws.Protect "other"', 'ws.Unprotect "other"'),
    ("protect_no_password_then_one", 'ws.Protect: ws.Protect "pw"', "ws.Unprotect"),
    ("user_interface_then_plain", "ws.Protect UserInterfaceOnly:=True: ws.Protect", 'ws.Range("A1").Value = 1'),
    ("user_interface_format", "ws.Protect UserInterfaceOnly:=True", 'ws.Range("A1").Interior.Color = vbRed'),
    # What each kind of format says when it is refused.
    ("interior_pattern", "ws.Protect", 'ws.Range("A1").Interior.Pattern = xlSolid'),
    ("interior_color_index", "ws.Protect", 'ws.Range("A1").Interior.ColorIndex = 3'),
    ("font_size", "ws.Protect", 'ws.Range("A1").Font.Size = 20'),
    ("font_color", "ws.Protect", 'ws.Range("A1").Font.Color = vbRed'),
    ("font_name", "ws.Protect", 'ws.Range("A1").Font.Name = "Arial"'),
    ("horizontal_alignment", "ws.Protect", 'ws.Range("A1").HorizontalAlignment = xlCenter'),
    ("wrap_text", "ws.Protect", 'ws.Range("A1").WrapText = True'),
    ("border_line", "ws.Protect", 'ws.Range("A1").Borders.LineStyle = xlContinuous'),
    ("border_weight", "ws.Protect", 'ws.Range("A1").Borders(xlEdgeTop).Weight = xlThick'),
    ("indent", "ws.Protect", 'ws.Range("A1").IndentLevel = 2'),
    ("formula_hidden", "ws.Protect", 'ws.Range("A1").FormulaHidden = True'),
    ("clear_formats", "ws.Protect", 'ws.Range("A1").ClearFormats'),
    ("sort", 'ws.Range("A1:A3").Value = Application.Transpose(Array(3, 1, 2)): ws.Protect',
     'ws.Range("A1:A3").Sort Key1:=ws.Range("A1"), Header:=xlNo'),
    ("autofilter", 'ws.Range("A1:A3").Value = Application.Transpose(Array("h", 1, 2)): ws.Protect',
     'ws.Range("A1:A3").AutoFilter Field:=1, Criteria1:="1"'),
    ("insert_cells", "ws.Protect AllowInsertingRows:=True", 'ws.Range("A2").Insert Shift:=xlDown'),
    # A second Protect that sets options on a sheet with no password, or passes one it already has.
    # Unprotect with no password on a sheet that has one opens a prompt, so these name one.
    ("protect_options_new_password", "ws.Protect", 'ws.Protect "pw", Contents:=False: ws.Unprotect "bad"'),
    ("protect_options_new_password_opens", "ws.Protect", 'ws.Protect "pw", Contents:=False: ws.Unprotect "pw"'),
    ("protect_user_interface_again", 'ws.Protect "pw"', 'ws.Protect "pw", UserInterfaceOnly:=True: '
                                                      'ws.Range("A1").Value = 1'),
    ("protect_user_interface_wrong", 'ws.Protect "pw"', 'ws.Protect "bad", UserInterfaceOnly:=True'),
    ("protect_same_option_other_password", 'ws.Protect "pw"', 'ws.Protect "other", Contents:=True'),
    ("protect_allow_then_plain", "ws.Protect AllowFormattingCells:=True", 'ws.Protect: out = out & '
                                                                           'ws.Protection.AllowFormattingCells & ";"'),
    ("delete_cells", 'ws.Range("A2").Locked = False: ws.Protect', 'ws.Range("A2").Delete Shift:=xlUp'),
    ("unmerge", 'ws.Range("A1:B1").Merge: ws.Protect', 'ws.Range("A1:B1").UnMerge'),
    ("replace", 'ws.Range("A1").Value = "abc": ws.Protect', 'ws.Range("A1").Replace "b", "x", LookAt:=xlPart'),
    ("replace_unlocked", 'ws.Range("A1").Value = "abc": ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1").Replace "b", "x", LookAt:=xlPart: out = out & ws.Range("A1").Value & ";"'),
    ("replace_unprotected", 'ws.Range("A1").Value = "abc"',
     'ws.Range("A1").Replace "b", "x", LookAt:=xlPart: out = out & ws.Range("A1").Value & ";"'),
    ("fill_down", 'ws.Range("A1").Value = 1: ws.Protect', 'ws.Range("A1:A3").FillDown'),
    ("autofill", 'ws.Range("A1").Value = 1: ws.Protect', 'ws.Range("A1").AutoFill ws.Range("A1:A3")'),
    ("paste_special", 'ws.Range("A1").Value = 1: ws.Range("A1").Copy: ws.Protect',
     'ws.Range("B1").PasteSpecial xlPasteValues'),
    ("remove_duplicates", 'ws.Range("A1:A3").Value = 1: ws.Protect', 'ws.Range("A1:A3").RemoveDuplicates 1'),
    ("sort_allowed_unlocked", 'ws.Range("A1:A3").Value = Application.Transpose(Array(3, 1, 2)): '
                              'ws.Range("A1:A3").Locked = False: ws.Protect AllowSorting:=True',
     'ws.Range("A1:A3").Sort Key1:=ws.Range("A1"), Header:=xlNo'),
    ("autofilter_allowed", 'ws.Range("A1:A3").Value = Application.Transpose(Array("h", 1, 2)): '
                           "ws.Range(\"A1:A3\").AutoFilter: ws.Protect AllowFiltering:=True",
     'ws.Range("A1:A3").AutoFilter Field:=1, Criteria1:="1"'),
    # Every other way the model changes a sheet.
    ("clear_contents_mixed", 'ws.Range("A1:A2").Value = 5: ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1:A2").ClearContents: out = out & CStr(ws.Range("A2").Value) & ";"'),
    ("clear_mixed", 'ws.Range("A1:A2").Value = 5: ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1:A2").Clear: out = out & CStr(ws.Range("A2").Value) & ";"'),
    ("clear_unlocked_locks", 'ws.Range("A1").Value = 5: ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1").Clear: out = out & ws.Range("A1").Locked & ";"'),
    ("clear_formats_unlocked", 'ws.Range("A1").Locked = False: ws.Protect',
     'ws.Range("A1").ClearFormats: out = out & ws.Range("A1").Locked & ";"'),
    ("border_around", "ws.Protect", 'ws.Range("A1").BorderAround xlContinuous'),
    ("cut_locked_source", 'ws.Range("A1").Value = 5: ws.Range("B1").Locked = False: ws.Protect',
     'ws.Range("A1").Cut ws.Range("B1")'),
    ("cut_unlocked", 'ws.Range("A1").Value = 5: ws.Range("A1:B1").Locked = False: ws.Protect',
     'ws.Range("A1").Cut ws.Range("B1"): out = out & CStr(ws.Range("B1").Value) & ";"'),
    ("paste_special_unlocked", 'ws.Range("A1").Value = 1: ws.Range("B1").Locked = False: ws.Protect: '
                               'ws.Range("A1").Copy',
     'ws.Range("B1").PasteSpecial xlPasteValues: out = out & CStr(ws.Range("B1").Value) & ";"'),
    ("paste_special_copied_after", 'ws.Range("A1").Value = 1: ws.Protect: ws.Range("A1").Copy',
     'ws.Range("B1").PasteSpecial xlPasteValues'),
    ("fill_right", 'ws.Range("A1").Value = 1: ws.Protect', 'ws.Range("A1:C1").FillRight'),
    ("fill_down_unlocked", 'ws.Range("A1").Value = 1: ws.Range("A2:A3").Locked = False: ws.Protect',
     'ws.Range("A1:A3").FillDown: out = out & CStr(ws.Range("A3").Value) & ";"'),
    ("replace_returns", 'ws.Range("A1").Value = "abc": ws.Protect',
     'out = out & ws.Range("A1").Replace("b", "x", LookAt:=xlPart) & ";"'),
    ("autofit", 'ws.Range("A1").Value = "a long text here": ws.Protect', "ws.Columns(1).AutoFit"),
    ("autofit_allowed", 'ws.Range("A1").Value = "a long text here": ws.Protect AllowFormattingColumns:=True',
     "ws.Columns(1).AutoFit"),
    ("show_all_data", 'ws.Range("A1:A3").Value = Application.Transpose(Array("h", 1, 2)): '
                      'ws.Range("A1:A3").AutoFilter 1, "1": ws.Protect', "ws.ShowAllData"),
    ("worksheet_paste", 'ws.Range("A1").Value = 1: ws.Protect: ws.Range("A1").Copy', 'ws.Paste ws.Range("B1")'),
    ("copy_sheet", "ws.Protect", 'ws.Copy After:=ws: out = out & wb.Worksheets(ws.Index + 1).ProtectContents & ";"'),
    ("range_name", "ws.Protect", 'ws.Range("A1").Name = "Here" & ws.Index'),
    ("merge_cells_property", "ws.Protect", 'ws.Range("A1:B1").MergeCells = True'),
    ("locked_allowed", "ws.Protect AllowFormattingCells:=True", 'ws.Range("A1").Locked = False'),
    ("hide_column", "ws.Protect", "ws.Columns(2).Hidden = True"),
    ("hide_column_allowed", "ws.Protect AllowFormattingColumns:=True", "ws.Columns(2).Hidden = True"),
    ("use_standard_height", "ws.Protect", "ws.Rows(1).UseStandardHeight = True"),
    ("use_standard_width", "ws.Protect", "ws.Columns(1).UseStandardWidth = True"),
    ("add_indent", "ws.Protect", 'ws.Range("A1").AddIndent = True'),
    ("reading_order", "ws.Protect", 'ws.Range("A1").ReadingOrder = xlRTL'),
    ("formula_r1c1", "ws.Protect", 'ws.Range("A1").FormulaR1C1 = "=1"'),
    ("value2", "ws.Protect", 'ws.Range("A1").Value2 = 1'),
    ("sort_object", 'ws.Range("A1:A3").Value = Application.Transpose(Array(3, 1, 2)): ws.Protect',
     'ws.Sort.SortFields.Clear: ws.Sort.SortFields.Add ws.Range("A1"): ws.Sort.SetRange ws.Range("A1:A3"): '
     "ws.Sort.Apply"),
    ("add_shape", "ws.Protect", "ws.Shapes.AddShape 1, 10, 10, 50, 50"),
    ("add_shape_drawing_false", "ws.Protect DrawingObjects:=False", "ws.Shapes.AddShape 1, 10, 10, 50, 50"),
    ("delete_sheet", "ws.Protect", "ws.Delete"),
    ("allow_after_unprotect", "ws.Protect AllowFormattingCells:=True, AllowSorting:=True: ws.Unprotect",
     'out = out & ws.Protection.AllowFormattingCells & "," & ws.Protection.AllowSorting & ";"'),
    ("enable_selection_set", "", 'ws.EnableSelection = xlUnlockedCells: out = out & ws.EnableSelection & ";"'),
]

#: Every format property the model writes: (object under the cell, property, a value to write).
FORMATS: list[tuple[str, str, str]] = [
    ("", "NumberFormat", '"0.00"'), ("", "HorizontalAlignment", "xlRight"), ("", "VerticalAlignment", "xlTop"),
    ("", "WrapText", "True"), ("", "Orientation", "45"), ("", "IndentLevel", "1"), ("", "ShrinkToFit", "True"),
    ("", "Locked", "False"), ("", "FormulaHidden", "True"),
    ("Font", "Name", '"Arial"'), ("Font", "Size", "14"), ("Font", "Bold", "True"), ("Font", "Italic", "True"),
    ("Font", "Underline", "xlUnderlineStyleSingle"), ("Font", "Strikethrough", "True"), ("Font", "Color", "vbRed"),
    ("Font", "ColorIndex", "3"), ("Font", "ThemeColor", "xlThemeColorAccent1"), ("Font", "TintAndShade", "0.5"),
    ("Font", "Superscript", "True"), ("Font", "Subscript", "True"), ("Font", "FontStyle", '"Bold"'),
    ("Interior", "Color", "vbRed"), ("Interior", "ColorIndex", "3"), ("Interior", "Pattern", "xlGray50"),
    ("Interior", "PatternColor", "vbBlue"), ("Interior", "PatternColorIndex", "5"),
    ("Interior", "ThemeColor", "xlThemeColorAccent1"), ("Interior", "TintAndShade", "0.5"),
    ("Borders", "LineStyle", "xlContinuous"), ("Borders", "Weight", "xlThick"), ("Borders", "Color", "vbRed"),
    ("Borders", "ColorIndex", "3"), ("Borders(xlEdgeLeft)", "LineStyle", "xlContinuous"),
    ("Borders(xlEdgeLeft)", "Weight", "xlThick"), ("Borders(xlEdgeLeft)", "Color", "vbRed"),
    ("Borders(xlEdgeLeft)", "ColorIndex", "3"), ("Borders(xlEdgeLeft)", "ThemeColor", "xlThemeColorAccent1"),
    ("Borders(xlEdgeLeft)", "TintAndShade", "0.5"),
]
CASES += [(f"format {one}.{prop}".replace(" .", " "), "ws.Protect",
           f'ws.Range("A1"){"." + one if one else ""}.{prop} = {value}') for one, prop, value in FORMATS]


#: Cases per procedure: a VBA procedure has a size limit.
BATCH = 12


def module() -> str:
    procedures: list[str] = []
    for start in range(0, len(CASES), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(wb As Object) As String", "Dim out As String, ws As Object"]
        for name, setup, action in CASES[start:start + BATCH]:
            lines += ["Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                      f'out = out & "{name}" & "="', "On Error Resume Next", "Err.Clear"]
            if setup:
                lines += [setup, 'If Err.Number <> 0 Then out = out & "setup " & Err.Number & ";"', "Err.Clear"]
            lines += [action, 'out = out & "E" & Err.Number & ":" & Err.Description & ";"', "Err.Clear",
                      f"out = out & {STATE}", "On Error GoTo 0", 'out = out & "|"']
        lines += [f"Batch{start // BATCH} = out", "End Function"]
        procedures.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim wb As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             *[f"out = out & Batch{index}(wb)" for index in range(len(procedures))],
             "wb.Close False", "Probe = out", "End Function"]
    return "\n".join([*procedures, *probe]) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    parts = [part for part in str(result.value).split("|") if part]
    answers = dict(part.split("=", 1) for part in parts)
    assert list(answers) == [name for name, _, _ in CASES], list(answers)
    record = {"state": STATE, "cases": [{"name": name, "setup": setup, "action": action, "answer": answers[name]}
                                        for name, setup, action in CASES]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, _, _ in CASES:
        print(f"{name:28} {answers[name]}")


if __name__ == "__main__":
    main()
