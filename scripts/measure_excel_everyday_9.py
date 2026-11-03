"""A ninth sweep of everyday Excel macros: what each ordinary edit leaves selected. B2:C3 is selected with C3 active before the
edit, and the probe reads the selection, the active cell and the active sheet after it.

    python scripts/measure_excel_everyday_9.py

writes tests/fixtures/excel_everyday_9.json, which tests/test_excel_everyday_7.py replays with the seventh and eighth.
"""
from __future__ import annotations

from pathlib import Path

import measure_excel_everyday_7 as seventh
import measure_excel_everyday_8 as eighth
from measure_excel_everyday_7 import LIST, PRE
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_9.json'
READ = ('\ns = Selection.Address & "|" & ActiveCell.Address & "|" & (ActiveSheet.Name = ws.Name) & "|" & ws.Range("B2").Value\n'
        'On Error Resume Next\nws.Activate\ns = s & "|" & Selection.Address & "|" & ActiveCell.Address\nOn Error GoTo 0')
#: Edits that reach the selected cells, or move them.
REACHING = {
    'rows_above_insert': 'ws.Rows(1).Insert',
    'rows_above_delete': 'ws.Rows(1).Delete',
    'row_in_insert': 'ws.Rows(3).Insert',
    'row_in_delete': 'ws.Rows(3).Delete',
    'rows_all_delete': 'ws.Rows("2:3").Delete',
    'rows_below_delete': 'ws.Rows(10).Delete',
    'columns_left_insert': 'ws.Columns(1).Insert',
    'columns_left_delete': 'ws.Columns(1).Delete',
    'column_in_delete': 'ws.Columns(2).Delete',
    'cell_in_delete': 'ws.Range("B2").Delete xlShiftUp',
    'cell_above_insert': 'ws.Range("B1").Insert xlShiftDown',
    'selection_delete': 'Selection.Delete xlShiftUp',
    'selection_insert': 'Selection.Insert xlShiftDown',
    'selection_clear': 'Selection.ClearContents',
    'selection_entire_row_delete': 'Selection.EntireRow.Delete',
    'selection_entire_row_insert': 'Selection.EntireRow.Insert',
    'selection_merge': 'Selection.Merge',
    'selection_sort': 'Selection.Sort ws.Range("B2")',
    'selection_cut': 'Selection.Cut ws.Range("H1")',
    'selection_copy': 'Selection.Copy ws.Range("H1")',
    'selection_copy_clipboard': 'Selection.Copy',
    'cut_onto_selection': 'ws.Range("H1:I2").Value = 1\nws.Range("H1:I2").Cut ws.Range("B2")',
    'cut_active_cell': 'ActiveCell.Cut ws.Range("H1")',
    'copy_paste_into_selection': 'ws.Range("A1").Copy\nws.Paste',
    'copy_paste_elsewhere': 'ws.Range("A1").Copy\nws.Paste ws.Range("H1")',
    'paste_special_elsewhere': 'ws.Range("A1:B2").Copy\nws.Range("H1").PasteSpecial xlPasteValues',
    'paste_special_all_elsewhere': 'ws.Range("A1:B2").Copy\nws.Range("H1").PasteSpecial xlPasteAll',
    'paste_special_formats_elsewhere': 'ws.Range("A1:B2").Copy\nws.Range("H1").PasteSpecial xlPasteFormats',
    'paste_special_big_selection': 'ws.Range("A1").Copy\nSelection.PasteSpecial xlPasteValues',
    'insert_copied_above': 'ws.Rows(9).Copy\nws.Rows(1).Insert',
    'insert_cut_above': 'ws.Rows(9).Cut\nws.Rows(1).Insert',
    'active_moves_in_selection': 'ws.Range("B2").Activate',
    'active_moves_outside': 'ws.Range("H1").Activate',
    'range_select_two_areas': 'ws.Range("A1,C3").Select',
    'rows_select': 'ws.Rows(3).Select',
    'columns_select': 'ws.Columns(2).Select',
    'cells_select': 'ws.Cells.Select',
    'used_select': 'ws.UsedRange.Select',
    'goto': 'Application.Goto ws.Range("H1")',
    'goto_scroll': 'Application.Goto ws.Range("H1"), True',
    'sheet_add_stays': 'Worksheets.Add',
    'sheet_add_after_stays': 'Worksheets.Add After:=ws',
    'sheet_copy_stays': 'ws.Copy After:=ws',
    'filter_on_selection': 'ws.Range("A1:E5").AutoFilter 2, "North"',
    'hide_selected_row': 'ws.Rows(3).Hidden = True',
    'hide_selected_column': 'ws.Columns(3).Hidden = True',
    'hide_all_selected_rows': 'ws.Rows("2:3").Hidden = True',
    'table_over_selection': 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:E5"), , xlYes',
    'text_to_columns_on_selection': 'ws.Range("A7").Value = "a b"\nws.Range("A7").TextToColumns ws.Range("A7"), Space:=True',
    'subtotal_stays': 'ws.Range("A1:E5").Subtotal 2, xlSum, Array(4)',
    'remove_subtotal_stays': 'ws.Range("A1:E5").Subtotal 2, xlSum, Array(4)\nws.Range("A1").CurrentRegion.RemoveSubtotal',
    'autofill_from_selection': 'ws.Range("Z1").Value = 1\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")',
    'shape_add_stays': 'ws.Shapes.AddShape msoShapeRectangle, 10, 10, 50, 50',
    'shape_selected': 'ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50).Select',
    'chart_add_stays': 'ws.ChartObjects.Add 10, 10, 200, 120',
}
EDITS = {**seventh.EDITS, **eighth.EDITS}
for left_out in ('sheet_add', 'sheet_copy', 'sheet_add_after', 'sheet_move', 'sheet_visible'):
    del EDITS[left_out]  # They leave a sheet behind or come back to this one, which REACHING asks instead.
CASES = {
    **{f'selected_after_{name}': f'{LIST}ws.Activate\nws.Range("B2:C3").Select\nws.Range("C3").Activate\n{edit}{READ}'
       for name, edit in {**EDITS, **REACHING}.items()},
}
#: A probe that adds a sheet deletes it again, whichever it is, once it has read the answer.
TIDY = ('\nApplication.DisplayAlerts = False\nFor i = Worksheets.Count To 1 Step -1\nIf Worksheets(i).Name <> ws.Name And Left(Worksheets(i).Name, 5) '
        '= "Sweep" Then Worksheets(i).Delete\nNext\nFor i = Worksheets.Count To 1 Step -1\nIf Worksheets(i).Name <> ws.Name And Worksheets(i).Index '
        '<= 2 And Worksheets.Count > 2 Then Worksheets(i).Delete\nNext\nApplication.DisplayAlerts = True')


def macro(name: str) -> str:
    cleanup = ("\nOn Error Resume Next\nws.Unprotect\nOn Error GoTo 0\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\n"
               "ws.Delete\nApplication.DisplayAlerts = True")
    tidy = TIDY if name.startswith('selected_after_sheet_') else ''
    return "\nFunction Probe() As String\n" + PRE + "Application.DisplayAlerts = False\n" + CASES[name] + tidy + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
