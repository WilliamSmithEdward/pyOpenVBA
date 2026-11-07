"""Which characters Excel takes in a name after its first, beyond ASCII. Three rules were written down before Excel answered:

A. Letters and digits (Python's \\w), and the Latin combining accents U+0300-U+036F: what the model took.
B. Letters, digits and combining marks of any script (Unicode categories L, N and M).
C. Every character from U+0080 up, as Excel takes for the first character (tests/fixtures/name_validity.json).

Each name is "a" and the character, added with Names.Add on a new book; Excel's error number and the model's are written to
tests/fixtures/name_later_characters.json with each rule's prediction.

    python scripts/measure_name_later_characters.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'name_later_characters.json'
#: Marks of several scripts, letters, digits, symbols, punctuation, spaces and format characters.
POINTS = [0x0E37, 0x0E48, 0x0E31, 0x0301, 0x093F, 0x094D, 0x0651, 0x05B4, 0x3099, 0x20DD, 0xFE0F,
          0x4E2D, 0x3042, 0xAC00, 0x0627, 0x0E01, 0x00E9, 0x0661, 0x0E51, 0xFF11, 0x00B2, 0x2167,
          0x20AC, 0x00D7, 0x00F7, 0x2014, 0x00BF, 0x00A1, 0x3001, 0xFF01, 0x00B7, 0x2026, 0x00AB, 0x2018,
          0x00A0, 0x3000, 0x2003, 0x200B, 0x200D, 0x00AD, 0x2060, 0xFEFF, 0x0085, 0x2028,
          0x00A9, 0x2122, 0x263A, 0x2190, 0x00B0, 0x00B5, 0x00AA, 0xE000, 0xFFFD]


def rule_a(ch: str) -> bool:
    import re
    return re.fullmatch(r'[\w.\\?̀-ͯ]', ch) is not None


def rule_b(ch: str) -> bool:
    return unicodedata.category(ch)[0] in 'LNM'


def rule_c(ch: str) -> bool:
    return ord(ch) >= 0x80


def source() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Object, s As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next']
    for point in POINTS:
        lines += [f'    wb.Names.Add "a" & ChrW({point}), "=1"', '    s = s & " " & Err.Number: Err.Clear']
    lines += ['    s = s & " | " & wb.Names.Count', '    On Error GoTo 0', '    wb.Close False', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def model() -> str:
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(), name='Later')
    return str(app.run('Later.Probe'))


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    modelled = model()
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
            measured = str(done.value)
    finally:
        os.rmdir(lock)
    excel_errors, excel_count = measured.split(' | ')
    model_errors, model_count = modelled.split(' | ')
    cases = []
    for point, live, ours in zip(POINTS, excel_errors.split(), model_errors.split()):
        ch = chr(point)
        cases.append({'point': f'U+{point:04X}', 'category': unicodedata.category(ch), 'excel': int(live), 'model': int(ours),
                      'rule_a': 0 if rule_a(ch) else 1004, 'rule_b': 0 if rule_b(ch) else 1004, 'rule_c': 0 if rule_c(ch) else 1004})
    OUT.write_text(json.dumps({'excel_count': int(excel_count), 'model_count': int(model_count), 'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for rule in ('rule_a', 'rule_b', 'rule_c'):
        print(rule, 'agrees on', sum(case['excel'] == case[rule] for case in cases), 'of', len(cases))
    for case in cases:
        if case['excel'] != case['model']:
            print(case)


if __name__ == '__main__':
    main()
