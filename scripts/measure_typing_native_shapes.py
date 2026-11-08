"""Whether Excel types digits of another script as it types the same text in ASCII digits. tests/fixtures/
typing_native_digits.json had Arabic-Indic, Thai and 15 more sets read as numbers alone, with a minus or a point,
and kept text about a comma, as ASCII "1,2" is. The rule written before this probe:

F. Excel reads each digit of those 17 sets as its ASCII digit and types the text as it types that: the same value,
   type and number format, for dates, times, percents, currency, exponents and fractions too; the digits of two
   sets, or of a set and ASCII, together, the same.

Each shape is set in a cell spelt with ASCII digits and spelt with a set's digits, and the cell's TypeName, value
and NumberFormat are read; the rule's prediction is that the two answer alike. Also Arabic's own decimal and
thousands separators and percent sign, which F leaves as they are. Written to
tests/fixtures/typing_native_shapes.json.

    python scripts/measure_typing_native_shapes.py
"""
from __future__ import annotations

import json
import os
import subprocess
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'typing_native_shapes.json'
ZEROS = [0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6, 0x0D66, 0x0E50, 0x0ED0,
         0x0F20, 0x1040, 0x17E0, 0x1810]
SHAPES = ['12', '1,2', '12,345', '1/2', '1/2/2020', '12/31/1999', '1:2', '12:30 PM', '1:02:03', '1%', '(12)', '$12',
          '1e3', ' 12', '12 ', '1 1/2', '12-Jan', 'Jan 12', '-1.5', '+12', '1.2.3', '12a', 'a12', '1/2/2020 3:04',
          '0012', '1234567890123456789', '.5', '1-2']
#: Arabic's decimal separator, thousands separator and percent sign beside Arabic-Indic digits.
ARABIC = ['١٫٢', '١٢٬٣٤٥', '١٢٪', '1٫2']
CHUNK = 30


def native(text: str, zero: int) -> str:
    return ''.join(chr(zero + int(c)) if c.isdigit() else c for c in text)


def pairs() -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    for zero in ZEROS[:4] + [0x0E50]:
        for shape in SHAPES:
            found.append({'shape': shape, 'zero': f'{zero:04X}', 'text': native(shape, zero)})
    for zero in ZEROS[4:]:
        if zero == 0x0E50:
            continue
        for shape in ('1/2/2020', '12:30 PM', '1%', '$12', '1e3'):
            found.append({'shape': shape, 'zero': f'{zero:04X}', 'text': native(shape, zero)})
    # Mixed: a set with ASCII, and two sets.
    for shape, text in (('12', '1' + chr(0x0662)), ('12', chr(0x0661) + '2'), ('12', chr(0x0661) + chr(0x0E52)),
                        ('1/2/2020', chr(0x0661) + '/2/' + native('2020', 0x0E50)), ('1.5', chr(0x0661) + '.' + chr(0x0965 + 6))):
        found.append({'shape': shape, 'zero': 'mixed', 'text': text})
    for text in ARABIC:
        found.append({'shape': None, 'zero': 'arabic signs', 'text': text})
    return found


def vba(text: str) -> str:
    return ' & '.join(f'ChrW({ord(c)})' if ord(c) > 126 or c == '"' else f'"{c}"' for c in text) or '""'


def source(texts: list[str]) -> str:
    lines = ['Public Function One() As String', '    Dim wb As Object, ws As Object, s As String, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    for text in texts:
        lines += ['    ws.Range("A1").Clear', f'    ws.Range("A1").Value = {vba(text)}', '    v = ws.Range("A1").Value',
                  '    s = s & TypeName(v) & " " & IIf(TypeName(v) = "String", "", CStr(v)) & " " & ws.Range("A1").NumberFormat & "|"']
    lines += ['    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(texts: list[str]) -> list[str]:
    """The model's answers for ``texts``, read as the probe reads Excel's."""
    import sys

    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    answers: list[str] = []
    for text in texts:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source([text]), name='Shapes')
        try:
            answers += str(app.run('Shapes.One')).split('|')[:-1]
        except VBAError as failed:
            answers.append(f'model {type(failed).__name__}: {str(failed)[:120]}')
    return answers


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = pairs()
    texts: list[str] = []
    for case in cases:
        texts.append(case['text'])
        texts.append(case['shape'] if case['shape'] is not None else case['text'])
    chunks = [texts[at:at + CHUNK] for at in range(0, len(texts), CHUNK)]
    answers: list[str] = []
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
            for chunk in chunks:
                done = excel.run_vba(source(chunk), 'One', (), timeout=120.0, line_numbers=False)
                answers += str(done.value).split('|')[:-1] if done.ok else [f'no answer: {str(done)[:200]}'] * len(chunk)
    finally:
        os.rmdir(lock)
    for at, case in enumerate(cases):
        case['excel'], case['excel_ascii'] = answers[2 * at], answers[2 * at + 1]
    OUT.write_text(json.dumps({'cases': cases}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    shaped = [case for case in cases if case['shape'] is not None]
    print('rule F agrees on', sum(case['excel'] == case['excel_ascii'] for case in shaped), 'of', len(shaped))
    for case in cases:
        if case['shape'] is None or case['excel'] != case['excel_ascii']:
            print(case['zero'], repr(case['shape']), ' '.join(f'{ord(c):04X}' for c in case['text']), '| Excel', case['excel'],
                  '| ASCII', case['excel_ascii'])


if __name__ == '__main__':
    main()
