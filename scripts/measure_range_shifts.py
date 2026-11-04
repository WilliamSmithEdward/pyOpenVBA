"""Cells shifted on a sheet whose rows or columns carry formats: which format each position then shows, its value, and the
same after a save and reopen, late and early bound.

    python scripts/measure_range_shifts.py

writes tests/fixtures/range_shifts.json, which tests/test_excel_range_shifts.py replays. Its "early" holds the answers
that differ with o typed.
"""
from __future__ import annotations

from pathlib import Path

import measure_method_audit as audit

_VALUES = ('ws.Range("B2").Value = 1: ws.Range("B3").Value = 2: ws.Range("B5").Value = 5: ws.Range("B6").Value = 6: ws.Range("C5").Value = 7: '
           'ws.Range("D3").Value = 8: ws.Range("E3").Value = 9')
_ROW = 'ws.Rows(5).Font.Bold = True'
_ROWS = 'ws.Rows("5:6").Font.Bold = True'
_COLUMN = 'ws.Columns(4).Font.Bold = True'
_COLUMNS = 'ws.Columns("D:E").Font.Bold = True'
_BOTH = 'ws.Rows(5).Font.Bold = True: ws.Columns(4).Interior.Color = 255'
_OWN = 'ws.Rows(5).Font.Bold = True: ws.Range("B5").Font.Italic = True: ws.Range("B4").Font.Underline = xlUnderlineStyleSingle'
_SHEET = 'ws.Cells.Font.Bold = True'
CELLS = [f'{column}{row}' for column in 'ABCDEFG' for row in range(1, 9)]
BOLD = [f'ws.Range("{cell}").Font.Bold' for cell in CELLS]
ALL = [*BOLD, *[f'ws.Range("{cell}").Formula' for cell in ('B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'C3', 'C5', 'D3', 'E3', 'F3')],
       'ws.Range("B5").Font.Italic', 'ws.Range("B6").Font.Italic', 'ws.Range("B4").Font.Underline', 'ws.Range("B5").Font.Underline',
       'ws.Range("B7").Font.Bold', 'ws.Range("D2").Interior.Color', 'ws.Range("D5").Interior.Color', 'ws.Range("E5").Interior.Color', 'ws.Range("D6").Interior.Color',
       'ws.Rows(5).Font.Bold', 'ws.Rows(6).Font.Bold', 'ws.Columns(4).Font.Bold', 'ws.Columns(5).Font.Bold', 'ws.UsedRange.Address']
CALLS = [('o.Insert xlShiftDown', []), ('o.Insert xlShiftToRight', []), ('o.Delete xlShiftUp', []), ('o.Delete xlShiftToLeft', []),
         ('o.Insert xlShiftDown, xlFormatFromRightOrBelow', []), ('o.Insert xlShiftToRight, xlFormatFromRightOrBelow', [])]
METHODS: list[tuple[str, str, str, str, str, list[str], list[tuple[str, list[str]]]]] = [
    ('Range', 'Shift', 'Range', target, f'{_VALUES}: {setup}', ALL, CALLS)
    for setup in (_ROW, _ROWS, _COLUMN, _COLUMNS, _BOTH, _OWN, _SHEET)
    for target in ('ws.Range("B3")', 'ws.Range("B3:C4")', 'ws.Range("B5")', 'ws.Range("C2:D2")', 'ws.Range("B6:C7")')
]


def _read(number: int, reads: list[str]) -> str:
    """Every case reads the same cells, so one function does it, or the module is too large for Excel to load."""
    lines = ''.join(f'    v = Empty\n    v = {read}\n    If Err.Number <> 0 Then\n        s = s & "|e" & Err.Number\n'
                    f'        Err.Clear\n    Else\n        s = s & "|" & Show(v)\n    End If\n' for read in reads)
    shared = (f'\nPrivate Function Grid(ws As Object) As String\n    Dim s As String, v As Variant\n'
              f'    On Error Resume Next\n{lines}    Grid = s\nEnd Function\n') if number == 0 else ''
    return shared + f'\nPrivate Function R{number}(ws As Object) As String\n    R{number} = Grid(ws)\nEnd Function\n'


def main() -> None:
    audit.METHODS = METHODS
    audit.WRONG_COUNT = r'^$'
    audit._read = _read
    audit.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_shifts.json'
    audit.main()


if __name__ == '__main__':
    main()
