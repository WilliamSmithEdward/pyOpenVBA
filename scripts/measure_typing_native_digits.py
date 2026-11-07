"""What a cell holds when VBA sets its Value to digits of another script: tests/fixtures/sort_scripts.json had Excel sort
a lone Arabic-Indic digit before every text, and leave a Thai digit out of a text filter's >, as a number would.
The rule written before this probe:

N. Excel reads a string of decimal digits of one script as the number they spell, for the 17 digit sets a name's
   address check folds (Arabic-Indic, extended Arabic-Indic, Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil,
   Telugu, Kannada, Malayalam, Thai, Lao, Tibetan, Myanmar, Khmer, Mongolian); any other digits stay text.

Every decimal digit of the Basic Multilingual Plane is set alone, then each set's 1 and 2 as "12", "-12", "1.2" and
"1,2" (sets whose zero is at a multiple of 16 or 10, so the run is contiguous). Written to
tests/fixtures/typing_native_digits.json, live Excel beside the model.

    python scripts/measure_typing_native_digits.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
import unicodedata
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'typing_native_digits.json'
FOLDED = {0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6, 0x0D66, 0x0E50, 0x0ED0, 0x0F20,
          0x1040, 0x17E0, 0x1810}
CHUNK = 40


def zeros() -> list[int]:
    return [code for code in range(0x80, 0x10000) if unicodedata.category(chr(code)) == 'Nd' and unicodedata.decimal(chr(code)) == 0]


def texts() -> list[str]:
    found = [chr(code) for code in range(0x80, 0x10000) if unicodedata.category(chr(code)) == 'Nd']
    for zero in zeros():
        one, two = chr(zero + 1), chr(zero + 2)
        found += [one + two, '-' + one + two, one + '.' + two, one + ',' + two]
    return found


def vba(text: str) -> str:
    return ' & '.join(f'ChrW({ord(c)})' if ord(c) > 126 else f'"{c}"' for c in text)


def source(chunk: list[str]) -> str:
    lines = ['Public Function One() As String', '    Dim wb As Object, ws As Object, s As String, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    for text in chunk:
        lines += [f'    ws.Range("A1").Value = {vba(text)}', '    v = ws.Range("A1").Value',
                  '    s = s & TypeName(v) & " " & IIf(TypeName(v) = "String", "", CStr(v)) & "|"']
    lines += ['    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def rule(text: str) -> str:
    digits = [c for c in text if unicodedata.category(c) == 'Nd']
    zero = ord(digits[0]) - unicodedata.decimal(digits[0])
    if zero not in FOLDED:
        return 'String '
    plain = ''.join(str(unicodedata.decimal(c)) if unicodedata.category(c) == 'Nd' else c for c in text)
    if plain == '1,2':
        return 'Double 12'
    return f'Double {float(plain):g}'


def model(chunk: list[str]) -> list[str]:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(chunk), name='Digits')
    return str(app.run('Digits.One')).split('|')[:-1]


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found = texts()
    chunks = [found[at:at + CHUNK] for at in range(0, len(found), CHUNK)]
    cases: list[dict[str, Any]] = []
    for chunk in chunks:
        for text, answer in zip(chunk, model(chunk)):
            cases.append({'text': text, 'rule': rule(text), 'model': answer})
    lock = _lock()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(10)
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            at = 0
            for chunk in chunks:
                done = excel.run_vba(source(chunk), 'One', (), timeout=120.0, line_numbers=False)
                answers = str(done.value).split('|')[:-1] if done.ok else [f'no answer: {str(done)[:200]}'] * len(chunk)
                for answer in answers:
                    cases[at]['excel'] = answer
                    at += 1
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': cases}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print('rule N agrees on', sum(c['excel'] == c['rule'] for c in cases), 'of', len(cases))
    print('model agrees on', sum(c['excel'] == c['model'] for c in cases), 'of', len(cases))
    for c in cases:
        if c['excel'] != c['rule']:
            print(' '.join(f'{ord(ch):04X}' for ch in c['text']), '| Excel', c['excel'], '| N', c['rule'], '| model', c['model'])


if __name__ == '__main__':
    main()
