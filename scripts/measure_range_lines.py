"""Range.Rows and Range.Columns: what the range they answer counts, indexes, walks and does.

A range Rows or Columns answers is the same cells taken a row or a column
at a time: its Count, its items and For Each go by rows or columns, while
its other members act on its cells. Each probe works on a sheet of its
own and gives what it read, an error as E and its number.

    python scripts/measure_range_lines.py

writes tests/fixtures/range_lines.json, which tests/test_excel_range_lines.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_lines.json'

#: Each read, by name: a VBA expression on ws, whose sheet has A1:F6 filled with their addresses.
READS = {
    'columns_count': 'ws.Range("A1:C3").Columns.Count',
    'rows_count': 'ws.Range("A1:C3").Rows.Count',
    'areas_columns_count': 'ws.Range("A1:C3,E1:F5").Columns.Count',
    'areas_rows_count': 'ws.Range("A1:C3,E1:F5").Rows.Count',
    'areas_cells_count': 'ws.Range("A1:C3,E1:F5").Cells.Count',
    'column_item': 'ws.Range("A1:C3").Columns(2).Address',
    'row_item': 'ws.Range("A1:C3").Rows(2).Address',
    'column_item_past': 'ws.Range("A1:C3").Columns(5).Address',
    'column_item_zero': 'ws.Range("A1:C3").Columns(0).Address',
    'row_item_negative': 'ws.Range("A1:C3").Rows(-1).Address',
    'columns_item_member': 'ws.Range("A1:C3").Columns.Item(2).Address',
    'columns_item_two': 'ws.Range("A1:C3").Columns.Item(1, 2).Address',
    'rows_item_two': 'ws.Range("A1:C3").Rows.Item(2, 1).Address',
    'columns_default_two': 'ws.Range("A1:C3").Columns(2, 1).Address',
    'column_then_cell': 'ws.Range("A1:C3").Columns(2)(3).Address',
    'columns_address': 'ws.Range("A1:C3").Columns.Address',
    'columns_value': 'ws.Range("B2").Columns.Value',
    'columns_typename': 'TypeName(ws.Range("A1:C3").Columns)',
    'columns_cells_count': 'ws.Range("A1:C3").Columns.Cells.Count',
    'columns_rows_count': 'ws.Range("A1:C3").Columns.Rows.Count',
    'columns_columns_count': 'ws.Range("A1:C3").Columns.Columns.Count',
    'column_columns_count': 'ws.Range("A1:C3").Columns(2).Columns.Count',
    'column_count': 'ws.Range("A1:C3").Columns(2).Count',
    'columns_offset_count': 'ws.Range("A1:C3").Columns.Offset(1, 0).Count',
    'columns_offset_item': 'ws.Range("A1:C3").Columns.Offset(1, 0)(2).Address',
    'columns_resize_count': 'ws.Range("A1:C3").Columns.Resize(2).Count',
    'columns_entire_count': 'ws.Range("A1:C3").Columns.EntireColumn.Count',
    'columns_areas_count': 'ws.Range("A1:C3").Columns.Areas(1).Count',
    'rows_cells_item': 'ws.Range("A1:C3").Rows.Cells(2).Address',
    'columns_first_cell': 'ws.Range("A1:C3").Columns.Cells(1, 2).Address',
    'columns_range_item': 'ws.Range("A1:C3").Columns.Range("B2").Address',
    'sheet_columns_item': 'ws.Columns("A:C")(2).Address',
    'sheet_columns_count': 'ws.Columns("A:C").Count',
    'sheet_rows_item': 'ws.Rows("1:3")(2).Address',
    'sheet_columns_cells_item': 'ws.Columns("A:C").Cells(2).Address',
    'sheet_columns_item_two': 'ws.Columns("B:C").Item(1, 1).Address',
    'range_columns_item': 'ws.Range("A:C")(2).Address',
    'range_columns_count': 'ws.Range("A:C").Columns.Count',
    'entire_column_count': 'ws.Range("B2").EntireColumn.Count',
    'entire_column_item': 'ws.Range("B2:C2").EntireColumn(2).Address',
    'entire_row_item': 'ws.Range("B2:B3").EntireRow(2).Address',
    'used_columns_count': 'ws.UsedRange.Columns.Count',
    'used_columns_item': 'ws.UsedRange.Columns(3).Address',
    'columns_hidden_read': 'ws.Range("A1:C3").Columns.Hidden',
    'columns_width_read': 'ws.Range("A1:C3").Columns.ColumnWidth',
    'rows_height_read': 'ws.Range("A1:C3").Rows.RowHeight',
    'columns_union': 'Union(ws.Range("A1:C3").Columns, ws.Range("E1:E2")).Count',
    'columns_intersect': 'Intersect(ws.Range("A1:C3").Columns, ws.Range("B1:D2")).Count',
    'sheet_areas_rows_count': 'ws.Rows("1:2,5:5").Count',
    'entire_areas_columns_count': 'ws.Range("A1,C1:D1").EntireColumn.Count',
    'entire_areas_item': 'ws.Range("A1,C1:D1").EntireColumn(2).Address',
    'columns_areas_item': 'ws.Range("A1:B2,E1:F1").Columns(3).Address',
    'columns_count_large': 'ws.Range("A1:C3").Columns.CountLarge',
    'rows_resize_item': 'ws.Range("A1:C3").Rows.Resize(2, 2)(2).Address',
    'column_entire_row': 'ws.Range("A1:C3").Columns(2).EntireRow.Address',
    'rows_value_count': 'UBound(ws.Range("A1:C3").Rows(2).Value, 2)',
}

#: Each walk: what For Each over the expression yields, address by address.
WALKS = {
    'walk_columns': 'ws.Range("A1:C3").Columns',
    'walk_rows': 'ws.Range("A1:C2").Rows',
    'walk_area_columns': 'ws.Range("A1:B2,E1:F1").Columns',
    'walk_sheet_columns': 'ws.Columns("B:D")',
    'walk_entire_rows': 'ws.Range("B2:B3").EntireRow',
    'walk_column_cells': 'ws.Range("A1:C3").Columns(2)',
    'walk_offset_columns': 'ws.Range("A1:B2").Columns.Offset(1, 0)',
    'walk_sheet_area_rows': 'ws.Rows("1:2,5:5")',
    'walk_cells_of_columns': 'ws.Range("A1:B2").Columns.Cells',
    'walk_entire_area_columns': 'ws.Range("A1,C1:D1").EntireColumn',
}

#: Each edit, then what it reads: (statement, reads after it).
EDITS = {
    'columns_width_write': ('ws.Range("A1:C3").Columns.ColumnWidth = 12',
                            'ws.Columns(1).ColumnWidth & "|" & ws.Columns(3).ColumnWidth & "|" & ws.Columns(4).ColumnWidth'),
    'rows_height_write': ('ws.Range("A1:C2").Rows.RowHeight = 30', 'ws.Rows(1).RowHeight & "|" & ws.Rows(3).RowHeight'),
    'columns_hidden_write': ('ws.Range("B1:C3").Columns.Hidden = True',
                             'ws.Columns(2).Hidden & "|" & ws.Columns(3).Hidden & "|" & ws.Columns(4).Hidden'),
    'rows_hidden_write': ('ws.Range("A2:C3").Rows.Hidden = True', 'ws.Rows(2).Hidden & "|" & ws.Rows(4).Hidden'),
    'columns_value_write': ('ws.Range("B2:C3").Columns.Value = 7', 'ws.Range("B2").Value & "|" & ws.Range("C3").Value & "|" & ws.Range("D3").Value'),
    'columns_font_write': ('ws.Range("A1:B1").Columns.Font.Bold = True', 'ws.Range("A1").Font.Bold & "|" & ws.Range("A2").Font.Bold'),
    'columns_clear': ('ws.Range("A1:B2").Columns.ClearContents', 'ws.Range("A1").Formula & "|" & ws.Range("C1").Formula'),
    'columns_delete': ('ws.Range("B1:B2").Columns.Delete', 'ws.Range("B1").Formula & "|" & ws.Range("B3").Formula'),
    'rows_delete': ('ws.Range("A2:B2").Rows.Delete', 'ws.Range("A2").Formula & "|" & ws.Range("C2").Formula'),
    'columns_insert': ('ws.Range("B1:B2").Columns.Insert', 'ws.Range("B1").Formula & "|" & ws.Range("C1").Formula & "|" & ws.Range("B3").Formula'),
    'columns_select': ('ws.Activate: ws.Range("B2:C3").Columns.Select', 'TypeName(Selection) & "|" & Selection.Address & "|" & Selection.Count'),
    'rows_autofit': ('ws.Range("A1").Font.Size = 30: ws.Range("A1:A2").Rows.AutoFit', 'ws.Rows(1).RowHeight & "|" & ws.Rows(2).RowHeight'),
    'cells_autofit': ('ws.Range("A1").Font.Size = 30: ws.Range("A1:A2").AutoFit', '"done"'),
}


def probe(body: str, reads: str) -> str:
    return '''Function Probe() As String
Dim ws As Worksheet, s As String, one As Range, r As Long, c As Long
Set ws = Worksheets.Add
For r = 1 To 6
    For c = 1 To 6
        ws.Cells(r, c).Value = "'" & ws.Cells(r, c).Address(False, False)
    Next
Next
On Error Resume Next
Err.Clear
''' + body + '''
If Err.Number <> 0 Then
    s = "E" & Err.Number
Else
    Err.Clear
    s = CStr(''' + reads + ''')
    If Err.Number <> 0 Then s = "E" & Err.Number
End If
On Error GoTo 0
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


def walk(expression: str) -> str:
    return probe(f'For Each one In {expression}\n    s = s & one.Address & ","\nNext', 's')


def main() -> None:
    modules = {name: probe('', expression) for name, expression in READS.items()}
    modules.update({name: walk(expression) for name, expression in WALKS.items()})
    modules.update({name: probe(statement, reads) for name, (statement, reads) in EDITS.items()})
    record(modules, OUT)


if __name__ == '__main__':
    main()
