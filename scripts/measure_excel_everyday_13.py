"""A thirteenth sweep of everyday Excel macros: where End and the last cell stop past hidden lines and at a table's edge, and what a
protected sheet lets them read. Written to split the rules the twelfth left open.

    python scripts/measure_excel_everyday_13.py

writes tests/fixtures/excel_everyday_13.json, which tests/test_excel_everyday_7.py replays with the others.
"""
from __future__ import annotations

from pathlib import Path

from measure_excel_everyday_7 import PRE
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_13.json'
LIST = ('ws.Range("A1:B1").Value = Array("N", "T")\nws.Range("A2:B2").Value = Array("a", "N")\nws.Range("A3:B3").Value = Array("b", "S")\n'
        'ws.Range("A4:B4").Value = Array("c", "N")\nws.Range("A5:B5").Value = Array("d", "S")\n')
READ = ('\ns = t & ws.UsedRange.Address & "|" & ws.Cells.SpecialCells(xlCellTypeLastCell).Address & "|" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & "|" '
        '& ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column & "|" & ws.Range("A1").End(xlDown).Row & "|" & ws.Range("A1").End(xlToRight).Column')
TABLE = 'Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B5"), , xlYes)\n'
EDITS = {
    'plain': '',
    'hide_last_data_row': 'ws.Rows(5).Hidden = True',
    'hide_two_last_rows': 'ws.Rows("4:5").Hidden = True',
    'hide_middle_row': 'ws.Rows(3).Hidden = True',
    'hide_beyond_data': 'ws.Rows(20).Hidden = True',
    'hide_beyond_and_formatted': 'ws.Rows(20).Hidden = True\nws.Range("A20").Font.Bold = True',
    'hide_last_column': 'ws.Columns(2).Hidden = True',
    'hide_beyond_column': 'ws.Columns(8).Hidden = True',
    'value_in_hidden_beyond': 'ws.Range("A20").Value = 1\nws.Rows(20).Hidden = True',
    'value_below_hidden': 'ws.Range("A20").Value = 1\nws.Rows(19).Hidden = True',
    'height_zero': 'ws.Rows(5).RowHeight = 0',
    'filter_hides_last': 'ws.Range("A1:B5").AutoFilter 2, "N"',
    'filter_hides_none_last': 'ws.Range("A1:B5").AutoFilter 2, "S"',
    'filter_then_off': 'ws.Range("A1:B5").AutoFilter 2, "N"\nws.AutoFilterMode = False',
    'table_plain': TABLE,
    'table_row_added': TABLE + 'o.ListRows.Add',
    'table_resized': TABLE + 'o.Resize ws.Range("A1:B8")',
    'table_empty_column_b': TABLE + 'o.ListRows.Add\nt = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row & "|"',
    'table_column_added': TABLE + 'o.ListColumns.Add',
    'table_values_below': TABLE + 'o.ListRows.Add\nws.Range("A20").Value = 1',
    'table_totals': TABLE + 'o.ShowTotals = True',
    'table_unlisted_after_add': TABLE + 'o.ListRows.Add\no.Unlist',
    'protected_last_cell': 'ws.Protect\nOn Error Resume Next\nv = ws.Cells.SpecialCells(xlCellTypeLastCell).Address\nt = Err.Number & "|" & v & "|"\nErr.Clear\nOn Error GoTo 0\nws.Unprotect',
    'protected_end': 'ws.Protect\nt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & "|" & ws.UsedRange.Address & "|"\nws.Unprotect',
    'protected_constants': 'ws.Protect\nOn Error Resume Next\nv = ws.Cells.SpecialCells(xlCellTypeConstants).Address\nt = Err.Number & "|"\nErr.Clear\nOn Error GoTo 0\nws.Unprotect',
}
CASES = {f'stops_{name}': f'{LIST}{edit}{READ}' for name, edit in EDITS.items()}


def macro(name: str) -> str:
    cleanup = "\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True"
    return "\nFunction Probe() As String\n" + PRE + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
