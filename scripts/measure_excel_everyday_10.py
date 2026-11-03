"""A tenth sweep of everyday Excel macros: when AutoFill moves the selection, what Selection is with a shape selected, and
ClearOutline with no outline to clear. Written to split the rival rules the ninth left open.

    python scripts/measure_excel_everyday_10.py

writes tests/fixtures/excel_everyday_10.json, which tests/test_excel_everyday_7.py replays with the others.
"""
from __future__ import annotations

from pathlib import Path

from measure_excel_everyday_7 import LIST, PRE
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_10.json'
WHERE = '\ns = Selection.Address & "|" & ActiveCell.Address & "|" & ws.Range("Z3").Value'
SEED = 'ws.Activate\nws.Range("Z1").Value = 1\nws.Range("Z2").Value = 2\n'
BOX = 'ws.Activate\nws.Range("B2:C3").Select\nws.Range("C3").Activate\nSet o = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)\no.Name = "Box"\n'
CASES = {
    'fill_source_selected': SEED + 'ws.Range("Z1").Select\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_source_is_selection': SEED + 'ws.Range("Z1:Z2").Select\nSelection.AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_two_selected_one_filled': SEED + 'ws.Range("Z1:Z2").Select\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_selection_elsewhere': SEED + 'ws.Range("A1").Select\nws.Range("Z1:Z2").AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_selection_is_destination': SEED + 'ws.Range("Z1:Z3").Select\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_selected_with_active_elsewhere': SEED + 'ws.Range("Z1:Z2").Select\nws.Range("Z2").Activate\nws.Range("Z1:Z2").AutoFill ws.Range("Z1:Z3")' + WHERE,
    'fill_right_source_selected': SEED + 'ws.Range("Z1").Select\nws.Range("Z1").AutoFill ws.Range("Z1:AB1")\ns = Selection.Address & "|" & ActiveCell.Address',
    'fill_other_sheet_active': SEED + 'Set ws2 = Worksheets.Add\nws.Range("Z1").AutoFill ws.Range("Z1:Z3")\ns = ActiveSheet.Name = ws2.Name\ns = s & ws.Range("Z3").Value\nws.Activate\ns = s & Selection.Address\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'shape_type_name': BOX + 'o.Select\ns = TypeName(Selection)',
    'shape_name': BOX + 'o.Select\ns = Selection.Name',
    'shape_left_top': BOX + 'o.Select\ns = Selection.Left & "|" & Selection.Top',
    'shape_left_set': BOX + 'o.Select\nSelection.Left = 40\ns = o.Left',
    'shape_range_count': BOX + 'o.Select\ns = Selection.ShapeRange.Count & Selection.ShapeRange(1).Name',
    'shape_address_fails': BOX + 'o.Select\ns = Selection.Address',
    'shape_active_cell_stays': BOX + 'o.Select\ns = ActiveCell.Address',
    'shape_then_range': BOX + 'o.Select\nws.Range("A1").Select\ns = TypeName(Selection) & Selection.Address',
    'shape_then_activate': BOX + 'o.Select\nws.Range("A1").Activate\ns = TypeName(Selection) & Selection.Address & ActiveCell.Address',
    'shape_selection_delete': BOX + 'o.Select\nSelection.Delete\ns = TypeName(Selection) & Selection.Address & ActiveCell.Address & ws.Shapes.Count',
    'shape_deleted_by_name': BOX + 'o.Select\no.Delete\ns = TypeName(Selection) & Selection.Address & ActiveCell.Address',
    'shape_other_deleted': BOX + 'Set c = Nothing\nDim p As Object\nSet p = ws.Shapes.AddShape(msoShapeOval, 80, 10, 50, 50)\no.Select\np.Delete\ns = TypeName(Selection)',
    'shape_oval': BOX + 'ws.Shapes.AddShape(msoShapeOval, 80, 10, 50, 50).Select\ns = TypeName(Selection)',
    'shape_rounded': BOX + 'ws.Shapes.AddShape(msoShapeRoundedRectangle, 80, 10, 50, 50).Select\ns = TypeName(Selection)',
    'shape_line': BOX + 'ws.Shapes.AddLine(1, 1, 50, 50).Select\ns = TypeName(Selection)',
    'shape_textbox': BOX + 'ws.Shapes.AddTextbox(1, 1, 1, 50, 50).Select\ns = TypeName(Selection)',
    'shape_two': BOX + 'Dim p As Object\nSet p = ws.Shapes.AddShape(msoShapeOval, 80, 10, 50, 50)\np.Name = "Ring"\nws.Shapes.Range(Array("Box", "Ring")).Select\ns = TypeName(Selection) & Selection.Count & Selection.ShapeRange.Count',
    'shape_add_to_selection': BOX + 'Dim p As Object\nSet p = ws.Shapes.AddShape(msoShapeOval, 80, 10, 50, 50)\no.Select\np.Select False\ns = TypeName(Selection) & Selection.Count',
    'shape_replace_true': BOX + 'Dim p As Object\nSet p = ws.Shapes.AddShape(msoShapeOval, 80, 10, 50, 50)\no.Select\np.Select True\ns = TypeName(Selection) & Selection.Name = p.Name',
    'shape_on_other_sheet': BOX + 'Set ws2 = Worksheets.Add\nOn Error Resume Next\no.Select\ns = Err.Number & TypeName(Selection)\nOn Error GoTo 0\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'shape_sheet_away_and_back': BOX + 'o.Select\nSet ws2 = Worksheets.Add\ns = TypeName(Selection)\nws.Activate\ns = s & TypeName(Selection)\nApplication.DisplayAlerts = False\nws2.Delete\nApplication.DisplayAlerts = True',
    'shape_insert_rows_keeps': BOX + 'o.Select\nws.Rows(1).Insert\ns = TypeName(Selection)',
    'shape_write_cell_keeps': BOX + 'o.Select\nws.Range("A1").Value = 1\ns = TypeName(Selection)',
    'shape_paste_elsewhere': BOX + 'o.Select\nws.Range("H1").Value = 1\nws.Range("H1").Copy\nOn Error Resume Next\nws.Paste\ns = Err.Number & TypeName(Selection) & ws.Range("B2").Value',
    'shape_range_on_selection': BOX + 'o.Select\nOn Error Resume Next\ns = Selection.Value\ns = s & "|" & Err.Number',
    'clear_outline_nothing': LIST + 'ws.Range("Z1").ClearOutline\ns = ws.Rows(1).OutlineLevel',
    'clear_outline_cell_in_group': LIST + 'ws.Rows("2:3").Group\nws.Range("A2").ClearOutline\ns = ws.Rows(2).OutlineLevel & ws.Rows(3).OutlineLevel & ws.Columns(1).OutlineLevel',
    'clear_outline_cells_in_group': LIST + 'ws.Rows("2:4").Group\nws.Range("A2:B3").ClearOutline\ns = ws.Rows(2).OutlineLevel & ws.Rows(3).OutlineLevel & ws.Rows(4).OutlineLevel',
    'clear_outline_column_group': LIST + 'ws.Columns("B:C").Group\nws.Range("B1").ClearOutline\ns = ws.Columns(2).OutlineLevel & ws.Columns(3).OutlineLevel',
    'clear_outline_both': LIST + 'ws.Rows("2:3").Group\nws.Columns("B:C").Group\nws.Range("B2").ClearOutline\ns = ws.Rows(2).OutlineLevel & ws.Rows(3).OutlineLevel & ws.Columns(2).OutlineLevel & ws.Columns(3).OutlineLevel',
}


def macro(name: str) -> str:
    cleanup = "\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
