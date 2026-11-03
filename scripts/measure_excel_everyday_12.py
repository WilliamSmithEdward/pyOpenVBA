"""A twelfth sweep of everyday Excel macros: where the sheet ends after each ordinary edit. The probe reads UsedRange, the last cell
SpecialCells finds, and End from the far edge of a row and a column, which is how macros find their last row.

    python scripts/measure_excel_everyday_12.py

writes tests/fixtures/excel_everyday_12.json, which tests/test_excel_everyday_7.py replays with the others.
"""
from __future__ import annotations

from pathlib import Path

from measure_excel_everyday_7 import LIST, PRE
from measure_excel_everyday_9 import EDITS, REACHING
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_12.json'
READ = ('\ns = ws.UsedRange.Address & "|" & ws.Cells.SpecialCells(xlCellTypeLastCell).Address & "|" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & "|" '
        '& ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column & "|" & ws.UsedRange.Rows.Count & "|" & ws.Range("A1").CurrentRegion.Address')
#: Selecting and adding sheets say nothing about where this one ends.
LEFT_OUT = {'range_select_two_areas', 'rows_select', 'columns_select', 'cells_select', 'used_select', 'goto', 'goto_scroll', 'active_moves_in_selection',
            'active_moves_outside', 'select', 'activate_cell', 'sheet_add_stays', 'sheet_add_after_stays', 'sheet_copy_stays', 'screen', 'events'}
CASES = {
    f'ends_after_{name}': f'{LIST}ws.Activate\nws.Range("B2:C3").Select\n{edit}{READ}'
    for name, edit in {**EDITS, **REACHING}.items() if name not in LEFT_OUT
}


def macro(name: str) -> str:
    cleanup = ("\nOn Error Resume Next\nws.Unprotect\nOn Error GoTo 0\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\n"
               "ws.Delete\nApplication.DisplayAlerts = True")
    return "\nFunction Probe() As String\n" + PRE + "Application.DisplayAlerts = False\n" + CASES[name] + cleanup + "\nProbe = s\nEnd Function\n"


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
