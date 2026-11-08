"""Which digits of other scripts make a name read as a cell's address. "a" with U+0661 is refused as A1 is, but "a" with the
fullwidth U+FF11 is taken (tests/fixtures/name_later_characters.json). Rules written before this probe:

A. Every decimal digit (Unicode Nd) beyond ASCII is a digit of an address.
B. Those Windows classes C1_DIGIT (GetStringTypeW) are.

Each BMP Nd digit d is tried as "a" & d (A1's shape), "xfe" & d (XFE1, the last column's) and "xfz" & d (no column, so a
name either way), on live Excel and the model, and written to tests/fixtures/name_address_digits.json.

    python scripts/measure_name_address_digits.py
"""
from __future__ import annotations

import ctypes
import json
import os
import subprocess
import tempfile
import time
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'name_address_digits.json'
PREFIXES = ('a', 'xfe', 'xfz')
DIGITS = [point for point in range(0x80, 0x10000) if not 0xD800 <= point <= 0xDFFF and unicodedata.category(chr(point)) == 'Nd']

SOURCE = '''
Public Function Probe(ByVal points As String) As String
    Dim wb As Object, p As Variant, s As String, prefix As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error Resume Next
    For Each p In Split(points, " ")
        For Each prefix In Array("a", "xfe", "xfz")
            wb.Names.Add prefix & ChrW(CLng(p)), "=1"
            s = s & " " & Err.Number
            Err.Clear
            wb.Names(prefix & ChrW(CLng(p))).Delete
            Err.Clear
        Next
    Next
    On Error GoTo 0
    wb.Close False
    Probe = Mid(s, 2)
End Function
'''


def windows_digit(point: int) -> bool:
    out = ctypes.c_ushort()
    ctypes.WinDLL('kernel32').GetStringTypeW(1, ctypes.c_wchar_p(chr(point)), 1, ctypes.byref(out))
    return bool(out.value & 0x4)


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyopenvba.apps.excel._names import valid

    lock = _lock()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(10)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True), ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            done = excel.run_vba(SOURCE, 'Probe', (' '.join(map(str, DIGITS)),), timeout=300.0, line_numbers=False)
            assert done.ok, done
            answers = [int(part) for part in str(done.value).split()]
    finally:
        os.rmdir(lock)
    cases = []
    for at, point in enumerate(DIGITS):
        for offset, prefix in enumerate(PREFIXES):
            name = prefix + chr(point)
            address = prefix != 'xfz'
            cases.append({'point': f'U+{point:04X}', 'prefix': prefix, 'excel': answers[3 * at + offset], 'model': 0 if valid(name) else 1004,
                          'rule_a': 1004 if address else 0, 'rule_b': 1004 if address and windows_digit(point) else 0})
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for rule in ('rule_a', 'rule_b', 'model'):
        print(rule, 'agrees on', sum(case['excel'] == case[rule] for case in cases), 'of', len(cases))
    print('refused:', sorted({case['point'] for case in cases if case['excel']}))


if __name__ == '__main__':
    main()
