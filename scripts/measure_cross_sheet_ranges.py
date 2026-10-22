"""What Worksheet.Range, Application.Range, Union and Intersect do with ranges on another sheet.

Each probe makes two sheets, ws and w2, asks for one range and records the
error number it raised and, where it gave one, the range's sheet and
address. The sheets are deleted afterwards.

    python scripts/measure_cross_sheet_ranges.py

writes tests/fixtures/cross_sheet_ranges.json, which
tests/test_excel_cross_sheet_ranges.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/cross_sheet_ranges.json'
#: Each probe's expression for the range, ws and w2 being the two sheets.
CASES = {
    'range_both_cells_other_sheet': 'ws.Range(w2.Cells(1, 1), w2.Cells(2, 2))',
    'range_one_cell_other_sheet': 'ws.Range(ws.Cells(1, 1), w2.Cells(2, 2))',
    'range_first_cell_other_sheet': 'ws.Range(w2.Cells(1, 1), ws.Cells(2, 2))',
    'range_one_range_other_sheet': 'ws.Range(w2.Range("A1:B2"))',
    'range_own_sheet': 'ws.Range(ws.Range("A1"), ws.Range("B2"))',
    'range_text_other_sheet': 'ws.Range("\'" & w2.Name & "\'!A1")',
    'range_text_own_sheet': 'ws.Range("\'" & ws.Name & "\'!A1:B2")',
    'range_name_other_sheet': 'ws.Range("CrossSheetProbe")',
    'range_name_own_sheet': 'w2.Range("CrossSheetProbe")',
    'application_range_other_sheet': 'Application.Range(w2.Cells(1, 1), w2.Cells(2, 2))',
    'application_range_mixed': 'Application.Range(ws.Cells(1, 1), w2.Cells(2, 2))',
    'application_range_text_other_sheet': 'Application.Range("\'" & w2.Name & "\'!B3")',
    'union_two_sheets': 'Union(ws.Range("A1"), w2.Range("A1"))',
    'union_one_sheet': 'Union(ws.Range("A1"), ws.Range("C3"))',
    'intersect_two_sheets': 'Intersect(ws.Range("A1:B2"), w2.Range("A1:B2"))',
    'intersect_one_sheet': 'Intersect(ws.Range("A1:B2"), ws.Range("B2:C3"))',
}


def macro(expression: str) -> str:
    return f'''Function Probe() As String
Dim ws As Worksheet, w2 As Worksheet, r As Range, s As String
Set ws = Worksheets.Add
Set w2 = Worksheets.Add
ThisWorkbook.Names.Add Name:="CrossSheetProbe", RefersTo:="='" & w2.Name & "'!$C$4"
On Error Resume Next
Set r = {expression}
s = CStr(Err.Number)
Err.Clear
If Not r Is Nothing Then s = s & "|" & IIf(r.Parent.Name = ws.Name, "ws", IIf(r.Parent.Name = w2.Name, "w2", "other")) & "|" & r.Address
On Error GoTo 0
ThisWorkbook.Names("CrossSheetProbe").Delete
Application.DisplayAlerts = False
ws.Delete
w2.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


if __name__ == '__main__':
    record({name: macro(expression) for name, expression in CASES.items()}, OUT)
