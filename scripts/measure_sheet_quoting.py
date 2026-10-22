"""Which sheet names Excel puts in apostrophes when a formula names the sheet.

Each probe adds a sheet of one name, writes ='name'!A1 to A1 of another
sheet with the name in apostrophes, and records the formula Excel reads
back, then deletes both sheets.

    python scripts/measure_sheet_quoting.py

writes tests/fixtures/sheet_quoting.json, which tests/test_excel_sheet_quoting.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sheet_quoting.json'
#: Each probe's sheet name, by what it is like.
NAMES = {
    'plain': 'Totals',
    'plain_digits': 'Data2',
    'underscore_first': '_x',
    'underscore_digit': 'x_1',
    'full_stop': 'x.y',
    'space': 'Data 1',
    'digit_first': '1Data',
    'hyphen': 'Ab-c',
    'dollar': 'a$b',
    'hash_first': '#x',
    'apostrophe': "O'Neil",
    'accented': 'Données',
    'cell': 'Q4',
    'cell_first': 'A1',
    'last_cell': 'XFD1048576',
    'past_last_column': 'XFE1',
    'four_letters': 'AAAA1',
    'past_last_row': 'A1048577',
    'cell_like_word': 'LOG10',
    'cell_then_letter': 'A1B',
    'letters': 'AB',
    'r1c1': 'R1C1',
    'r_alone': 'R',
    'c_alone': 'C',
    'rc': 'RC',
    'r_digits': 'R2',
    'c_digits': 'C3',
    'r_c_digits': 'R2C',
    'lower_r1c1': 'r1c1',
    'true': 'TRUE',
    'false': 'false',
    'function_name': 'SUM',
    'error_like': 'NA',
    'cell_like_not_a_function': 'LOH10',
    'cell_like_lower': 'q4',
    'rc_digit': 'RC1',
    'r1c1_then_letter': 'R1C1A',
    'full_stop_first': '.x',
    'full_stop_last': 'x.',
    'comma': 'x,y',
    'plus': 'x+y',
    'ampersand': 'x&y',
    'parenthesis': 'x(y)',
    'accented_first': 'Été',
    'cjk': '数据',
    'digits_only': '2020',
    'digit_full_stop': '1.5',
    'letter_digit_full_stop': 'x1.5',
    'r_letter': 'Rx',
    'c_letter': 'Cx',
    'r_digit_letter': 'R1x',
    'c_digit_letter': 'C1A',
    'rc_letter': 'RCx',
    'r1c_underscore': 'R1C1_',
    'rc_digit_letter': 'RC1x',
    'r_zero': 'R0',
    'c_zero': 'C0',
    'past_last_r1c1_row': 'R1048577',
    'past_last_r1c1_column': 'C16385',
    'last_r1c1_cell': 'R1048576C16384',
    'rr': 'RR',
    'cc': 'CC',
    'row_zero': 'A0',
    'leading_zero_row': 'A01',
    'two_letter_cell': 'AB12',
    'c_past_zzz': 'C18279',
    'c_65536': 'C65536',
    'c_65537': 'C65537',
    'c_large': 'C999999',
    'r1c_letter': 'R1Cx',
    'r1c0': 'R1C0',
    'r1c_past_last_column': 'R1C16385',
    'rc0': 'RC0',
    'r_leading_zero': 'R01',
    'r_digits_full_stop': 'R1.x',
    'c_digit_underscore': 'C1_',
    'r_digit_full_stop_last': 'R1.',
    'c_digit_full_stop': 'C1.x',
    'r1c1_full_stop': 'R1C1.x',
    'rc_digit_full_stop': 'RC1.x',
    'r_digits_decimal': 'R1.5',
    'c_last_row_number': 'C1048576',
    'c_past_last_row_number': 'C1048577',
    'rc_past_last_row_number': 'RC1048577',
    'log100': 'LOG100',
    'log1000': 'LOG1000',
    'log10_letter': 'LOG10X',
    'letter_log10': 'XLOG10',
    'log9': 'LOG9',
    'copyright_first': '©x',
    'copyright_last': 'x©',
    'euro_first': '€x',
    'euro_last': 'x€',
    'umlaut_digit': 'Ü1',
    'omega_word': 'Ωmega',
    'full_width_letters': 'ＡＢ',
    'full_width_digit': '１',
    'no_break_space': 'x y',
    'en_dash': 'x–y',
    'em_space_first': ' x',
}


def _vba_text(text: str) -> str:
    """``text`` as a VBA expression, each character past ASCII as ChrW, so the module's code page cannot change it."""
    pieces: list[str] = []
    run = ''
    for char in text:
        if ord(char) < 128:
            run += char.replace('"', '""')
            continue
        if run:
            pieces.append(f'"{run}"')
            run = ''
        pieces.append(f'ChrW({ord(char)})')
    if run or not pieces:
        pieces.append(f'"{run}"')
    return ' & '.join(pieces)


def macro(name: str) -> str:
    return f'''Function Probe() As String
Dim ws As Worksheet, target As Worksheet
Set ws = Worksheets.Add
Set target = Worksheets.Add(After:=ws)
target.Name = {_vba_text(name)}
target.Range("A1").Value = 7
ws.Range("A1").Formula = "='" & Replace(target.Name, "'", "''") & "'!A1"
Probe = ws.Range("A1").Formula & "|" & CStr(ws.Range("A1").Value)
Application.DisplayAlerts = False
ws.Delete
target.Delete
Application.DisplayAlerts = True
End Function
'''


if __name__ == '__main__':
    record({case: macro(name) for case, name in NAMES.items()}, OUT)
