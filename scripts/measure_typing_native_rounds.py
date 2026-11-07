"""Held-out rounds over typing another script's digits into a cell, the rule src/pyopenvba/apps/excel/_typing.py
gives (_native_digits, _native_format) after tests/fixtures/typing_native_digits.json and typing_native_shapes.json.

Each round draws seeded texts from all 17 digit sets: shapes those fixtures had and shapes they did not, numbers with
more places, longer whole numbers, signs, parentheses, currency with cents, percents with places, thousands,
times with seconds, months and years, dates with times, exponents, fractions, ASCII and a second set mixed in, and
text round the digits. Each is written to a cleared cell through a Range held As Range and As Object, and the cell's
Value (with its TypeName), Value2, NumberFormat and Text are read. The model's answers are written first, with
their digest, and committed before Excel answers:

    python scripts/measure_typing_native_rounds.py predict <seed>
    python scripts/measure_typing_native_rounds.py measure <seed>

Rounds go to tests/fixtures/typing_native_rounds/.
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'typing_native_rounds'
ZEROS = [0x0660, 0x06F0, 0x0966, 0x09E6, 0x0A66, 0x0AE6, 0x0B66, 0x0BE6, 0x0C66, 0x0CE6, 0x0D66, 0x0E50, 0x0ED0,
         0x0F20, 0x1040, 0x17E0, 0x1810]
SHAPES = ['7', '12', '-12', '+3', '0', '00', '1.5', '1.25', '1.50', '12.125', '-0.75', '.25', '(12)', '(1.5)',
          '12345678901', '123456789012', '999999999999999', '12345678.5', '$12', '$12.50', '-$12', '$1,234',
          '1%', '1.5%', '12.25%', '12,345', '12,345.5', '1,234,567', '1,2', '12:30', '12:30:45', '12:30:45 PM',
          '1/2', '1/2/2020', '12/31/1999', 'Jan-2020', '1/2/2020 3:04 PM', '12-Jan-2020', '1e3', '1.5e-3', '1 1/4',
          '3/4', '12a', 'x7', '1-2-3', '1.2.3', ' 12 ']
#: Shapes the second round on: formats typing brings that the first round had none of.
MORE_SHAPES = ['Jan-2020', '12/2020', 'Jan 2020', '12:30.5', '1:30:45.5', '25:30:00', '30:45', '1/2/2020 25:00',
               '€12', '12€', '%12', '12.5E+3', '1/2/2020 12:30:45 PM', '2020-01-02', '-1.5%', '($12.50)']
#: How a case writes: into a cleared cell, into one formatted beforehand, or by TextToColumns over the cell.
HOWS = ['clear'] * 6 + ['format:0.00', 'format:m/d/yyyy', 'format:@', 'format:0%', 'format:h:mm', 'textToColumns']
CHUNK = 20


def native(text: str, zero: int) -> str:
    return ''.join(chr(zero + int(c)) if c.isdigit() else c for c in text)


def texts(seed: int) -> list[dict[str, Any]]:
    pick = random.Random(seed)
    found: list[dict[str, Any]] = []
    # Rounds from 54102 on draw the further shapes and ways of writing too.
    shapes = SHAPES if seed <= 54101 else SHAPES + MORE_SHAPES * 2
    for _ in range(80 if seed <= 54101 else 120):
        shape = pick.choice(shapes)
        zero = pick.choice(ZEROS)
        text = native(shape, zero)
        roll = pick.random()
        digits = [at for at, c in enumerate(shape) if c.isdigit()]
        if roll < 0.15 and len(digits) > 1:
            # One digit left in ASCII.
            at = pick.choice(digits)
            text = text[:at] + shape[at] + text[at + 1:]
        elif roll < 0.25 and len(digits) > 1:
            # One digit of a second set.
            at = pick.choice(digits)
            other = pick.choice([z for z in ZEROS if z != zero])
            text = text[:at] + chr(other + int(shape[at])) + text[at + 1:]
        case = {'shape': shape, 'zero': f'{zero:04X}', 'text': text, 'typed': pick.random() < 0.5}
        if seed > 54101:
            case['how'] = pick.choice(HOWS)
        found.append(case)
    return found


def vba(text: str) -> str:
    return ' & '.join(f'ChrW({ord(c)})' if ord(c) > 126 or c == '"' else f'"{c}"' for c in text) or '""'


def source(cases: list[dict[str, Any]]) -> str:
    lines = ['Public Function One() As String', '    Dim wb As Object, ws As Object, s As String, v As Variant',
             '    Dim r As Range, o As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    Set r = ws.Range("A1"): Set o = ws.Range("A1")', '    ws.Columns(1).ColumnWidth = 40']
    for case in cases:
        holder = 'r' if case['typed'] else 'o'
        how = case.get('how', 'clear')
        lines.append(f'    {holder}.Clear')
        if how.startswith('format:'):
            lines.append(f'    {holder}.NumberFormat = "{how[7:]}"')
        if how == 'textToColumns':
            # The text goes in as a string, kept by the apostrophe, and TextToColumns reads it.
            lines += [f'    {holder}.Value = "\'" & {vba(case["text"])}',
                      f'    {holder}.TextToColumns Destination:={holder}, DataType:=xlDelimited, Tab:=True, Semicolon:=False, '
                      'Comma:=False, Space:=False, Other:=False']
        else:
            lines.append(f'    {holder}.Value = {vba(case["text"])}')
        lines += [f'    v = {holder}.Value',
                  f'    s = s & TypeName(v) & " " & IIf(TypeName(v) = "String", "", CStr(v)) & " ~ " & CStr({holder}.Value2) & " ~ " '
                  f'& {holder}.NumberFormat & " ~ " & {holder}.Text & "|"']
    lines += ['    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(cases: list[dict[str, Any]]) -> list[str]:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    answers: list[str] = []
    for case in cases:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source([case]), name='Rounds')
        try:
            answers += str(app.run('Rounds.One')).split('|')[:-1]
        except VBAError as failed:
            answers.append(f'model {type(failed).__name__}: {str(failed)[:120]}')
    return answers


def digest(cases: list[dict[str, Any]]) -> str:
    return hashlib.sha256(json.dumps([[c['text'], c['typed'], c['model']] for c in cases], ensure_ascii=False).encode()).hexdigest()[:16]


def predict(seed: int) -> None:
    path = OUT / f'round_{seed}.json'
    assert not path.exists(), f'{path} exists: a round is predicted once'
    cases = texts(seed)
    for case, answer in zip(cases, model(cases)):
        case['model'] = answer
    OUT.mkdir(parents=True, exist_ok=True)
    record = {'seed': seed, 'predicted_at': datetime.now(timezone.utc).isoformat(), 'predictions_sha256_16': digest(cases), 'cases': cases}
    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print('predicted', len(cases), 'cases, digest', record['predictions_sha256_16'],
          '; model refuses', sum(c['model'].startswith('model ') for c in cases))


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def measure(seed: int) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    path = OUT / f'round_{seed}.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    cases = record['cases']
    assert digest(cases) == record['predictions_sha256_16']
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
            for at in range(0, len(cases), CHUNK):
                chunk = cases[at:at + CHUNK]
                done = excel.run_vba(source(chunk), 'One', (), timeout=120.0, line_numbers=False)
                answers += str(done.value).split('|')[:-1] if done.ok else [f'no answer: {str(done)[:200]}'] * len(chunk)
    finally:
        os.rmdir(lock)
    for case, answer in zip(cases, answers):
        case['excel'] = answer
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    refused = [c for c in cases if c['model'].startswith('model ')]
    wrong = [c for c in cases if c not in refused and c['model'] != c['excel']]
    print(f'{len(cases) - len(refused) - len(wrong)} right, {len(wrong)} wrong, {len(refused)} refused')
    for c in wrong + refused:
        print(c['zero'], repr(c['shape']), ' '.join(f'{ord(ch):04X}' for ch in c['text']), c['typed'], '| Excel', c['excel'],
              '| model', c['model'])


if __name__ == '__main__':
    (predict if sys.argv[1] == 'predict' else measure)(int(sys.argv[2]))
