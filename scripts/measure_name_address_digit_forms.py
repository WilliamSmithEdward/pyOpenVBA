"""Whether the digits tests/fixtures/name_address_digits.json found read as address digits (Arabic-Indic, Devanagari, Thai and
the rest from U+0660 to U+1819) make a name every kind of address its ASCII spelling would. The rule written first: such a
digit stands for its ASCII digit wherever Excel reads a name as an A1 or R1C1 reference; the fullwidth and N'Ko digits never.

Each name is a prefix and digits, tried on live Excel and the model, and written to tests/fixtures/name_address_digit_forms.json.

    python scripts/measure_name_address_digit_forms.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'name_address_digit_forms.json'
#: Devanagari, Thai and Mongolian fold; fullwidth and N'Ko do not.
ZEROS = {0x0966: True, 0x0E50: True, 0x1810: True, 0xFF10: False, 0x07C0: False}
#: (prefix, digits, between, digits after, what the ASCII spelling is)
SHAPES = [('xfd', '1', '', '', 'XFD1, an address'), ('xfd', '1048576', '', '', 'XFD1048576'), ('xfd', '1048577', '', '', 'past the last row'),
          ('ab', '12', '', '', 'AB12'), ('r', '1', '', '', 'R1'), ('c', '2', '', '', 'C2'), ('r', '1', 'c', '1', 'R1C1'), ('rc', '3', '', '', 'RC3'),
          ('r', '1', '.', '', 'R1., a name'), ('q', '1', '', '', 'Q1'), ('a', '1', '_', '', 'A1_, a name')]


def names() -> list[tuple[str, str, bool, str]]:
    out = []
    for zero, folds in ZEROS.items():
        for prefix, digits, between, after, what in SHAPES:
            spell = ''.join(chr(zero + int(d)) for d in digits)
            tail = ''.join(chr(zero + int(d)) for d in after)
            out.append((prefix + spell + between + tail, f'U+{zero:04X} {prefix}{digits}{between}{after}', folds, what))
    return out


def _vba(text: str) -> str:
    return ' & '.join(f'ChrW({ord(ch)})' if ord(ch) > 127 else f'"{ch}"' for ch in text)


def source() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Object, s As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next']
    for name, _, _, _ in names():
        lines += [f'    wb.Names.Add {_vba(name)}, "=1"', '    s = s & " " & Err.Number: Err.Clear']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def ascii_refused() -> dict[str, int]:
    """What the model, whose ASCII rules are measured (tests/fixtures/name_validity.json), says of each ASCII spelling."""
    from pyopenvba.apps.excel._names import valid

    return {f'{prefix}{digits}{between}{after}': 0 if valid(f'{prefix}{digits}{between}{after}') else 1004 for prefix, digits, between, after, _ in SHAPES}


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyopenvba.apps.excel._names import valid

    spelled = ascii_refused()
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
            done = excel.run_vba(source(), 'Probe', (), timeout=120.0, line_numbers=False)
            assert done.ok, done
            answers = [int(part) for part in str(done.value).split()]
    finally:
        os.rmdir(lock)
    cases = []
    for (name, label, folds, what), answer in zip(names(), answers):
        shape = label.split(' ', 1)[1]
        cases.append({'name': name, 'label': label, 'what': what, 'excel': answer, 'model': 0 if valid(name) else 1004,
                      'rule': spelled[shape] if folds else 0})
    OUT.write_text(json.dumps({'cases': cases}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print('rule agrees on', sum(case['excel'] == case['rule'] for case in cases), 'of', len(cases))
    for case in cases:
        if case['excel'] != case['rule']:
            print(case)


if __name__ == '__main__':
    main()
