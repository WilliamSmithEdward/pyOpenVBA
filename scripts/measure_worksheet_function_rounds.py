"""Held-out rounds over WorksheetFunction's PARTIAL rows: every function the checklist holds PARTIAL, called through a
WorksheetFunction object held typed and late bound, with arguments no fixture has asked.

Each call starts from one Excel answered before (tests/fixtures/worksheet_functions_more.json, worksheet_functions.json) and
is varied by the round's seed: its numbers scaled, negated or moved by one; one number given as text; one argument put in a
cell and the cell given; or one argument swapped for a value of another kind (Empty, Null, True, text, an error, a text cell,
a block of cells). FilterXML, Phonetic and Subtotal, which no fixture calls, start from calls written here. The calls run on
the fixtures' own sheet (their Fill) and are read with their Show.

Left out, and so left PARTIAL: the functions the model answers a unit or a few away in the fifteenth significant digit, and
Pi, which takes nothing to vary. A call the model refuses as not implemented is counted apart: no answer of the model's, so no
wrong one, and its row stays PARTIAL by its refusal.

From round 60103 an answer's characters below a space are written as ChrW codes, \\x01 and the like, by Show's caller, for
the harness hands a control character back as a space: round 60102's Unichar(True) read " " where Excel answered ChrW(1)
(tests/fixtures/worksheet_function_rivals.json).

    python scripts/measure_worksheet_function_rounds.py predict <seed>
    python scripts/measure_worksheet_function_rounds.py measure <seed>

tests/test_excel_worksheet_function_rounds.py replays the rounds.
"""
from __future__ import annotations

import csv
import hashlib
import json
import random
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as rounds  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests' / 'fixtures'
OUT = FIXTURES / 'worksheet_function_rounds'
MORE = json.loads((FIXTURES / 'worksheet_functions_more.json').read_text(encoding='utf-8'))
FIRST = json.loads((FIXTURES / 'worksheet_functions.json').read_text(encoding='utf-8'))
#: A call to each function no fixture asks.
WRITTEN = {'FilterXML': ['"<a><b>1</b><b>2</b></a>", "//b"'], 'Phonetic': ['ws.Range("B1")'], 'Subtotal': ['9, ws.Range("A1:A5")']}
#: The model answers these a unit or a few away in the fifteenth significant digit (tests/test_excel_worksheet_functions_more.py;
#: round 60101 for Confidence_T, F_Dist_RT and Gamma_Dist; round 60102 for ChiInv, and MultiNomial, whose log-gamma is not
#: Excel's: 16368 where Excel answers 16367.9999999998).
LAST_DIGIT = {'Beta_Dist', 'BetaDist', 'ChiInv', 'ChiSq_Inv', 'Confidence_T', 'F_Dist_RT', 'Gamma_Dist', 'ImSec',
              'MultiNomial', 'T_Dist_2T', 'T_Dist_RT', 'T_Inv_2T', 'TDist'}
#: Functions a round cannot vary: they take no argument.
NOTHING_TO_VARY = {'Pi'}
#: Each given through the Variant k, so a typed call meets it at run time rather than failing to compile.
OTHER_KINDS = ('k = Empty', 'k = Null', 'k = True', 'k = "abc"', 'k = CVErr(xlErrDiv0)', 'Set k = ws.Range("C2")', 'Set k = ws.Range("A1:A5")')
SAMPLE = 300
BATCH = 30
SEPARATOR = '|^|'


def partial_functions() -> list[str]:
    with open(ROOT / 'docs' / 'excel_checklist.csv', encoding='utf-8') as handle:
        return sorted(row['member'] for row in csv.DictReader(handle)
                      if row['object'] == 'WorksheetFunction' and row['status'] == 'PARTIAL'
                      and row['member'] not in LAST_DIGIT | NOTHING_TO_VARY)


def bases() -> dict[str, list[str]]:
    found: dict[str, list[str]] = {}
    for record in (MORE, FIRST):
        for case in record['cases']:
            found.setdefault(case['function'].lower(), []).append(case['arguments'])
    for name, calls in WRITTEN.items():
        found.setdefault(name.lower(), []).extend(calls)
    return found


def _outside(text: str) -> list[tuple[str, bool]]:
    """``text`` in pieces, each marked whether it lies outside a string literal and a #date#."""
    pieces = re.split(r'("(?:[^"]|"")*"|#[^#]*#)', text)
    return [(piece, index % 2 == 0) for index, piece in enumerate(pieces) if piece]


def _top_level(arguments: str) -> list[str]:
    parts, depth, current, quoted = [], 0, '', False
    for char in arguments:
        if char == '"':
            quoted = not quoted
        if not quoted:
            depth += char == '('
            depth -= char == ')'
            if char == ',' and depth == 0:
                parts.append(current.strip())
                current = ''
                continue
        current += char
    if current.strip():
        parts.append(current.strip())
    return parts


NUMBER = re.compile(r'(?<![\w.])\d+(?:\.\d+)?(?![\w.(])')


def _spell(value: float) -> str:
    return str(int(value)) if float(value).is_integer() and abs(value) < 1e15 else repr(value)


def _scaled(arguments: str, rng: random.Random) -> str:
    def move(found: re.Match[str]) -> str:
        number = float(found.group(0))
        return _spell(rng.choice((number * 2, number / 2, number + 1, number * 10, number * 3)))
    return ''.join(NUMBER.sub(move, piece) if outside else piece for piece, outside in _outside(arguments))


def variants(arguments: str, rng: random.Random) -> list[tuple[str, str]]:
    """The ways one measured call is varied: (setup statement, arguments)."""
    made = [('', _scaled(arguments, rng))]
    parts = _top_level(arguments)
    numeric = [at for at, part in enumerate(parts) if NUMBER.fullmatch(part.lstrip('-'))]
    if numeric:
        at = rng.choice(numeric)
        made.append(('', ', '.join(f'"{part}"' if index == at else part for index, part in enumerate(parts))))
        made.append((f'ws.Range("K1").Value = {parts[at]}', ', '.join('ws.Range("K1")' if index == at else part for index, part in enumerate(parts))))
    if parts:
        at = rng.randrange(len(parts))
        made.append((rng.choice(OTHER_KINDS), ', '.join('k' if index == at else part for index, part in enumerate(parts))))
    return [(setup, varied) for setup, varied in made if varied != arguments or setup]


def calls(seed: int) -> list[dict[str, str]]:
    rng = random.Random(seed)
    known = bases()
    every: list[dict[str, str]] = []
    first: list[dict[str, str]] = []
    for name in partial_functions():
        made = [{'function': name, 'setup': setup, 'arguments': varied}
                for arguments in known.get(name.lower(), []) for setup, varied in variants(arguments, rng)]
        if made:
            pick = rng.choice(made)
            first.append(pick)
            every += [one for one in made if one is not pick]
    chosen = first + rng.sample(every, max(0, min(len(every), SAMPLE - len(first))))
    rng.shuffle(chosen)
    return chosen


#: The first round whose answers write characters below a space as their codes.
ESCAPES_FROM = 60103
#: Show's answer with each character below a space written as \x and its code.
PLAIN = ('Private Function Plain(s As String) As String\n'
         '    Dim i As Long\n'
         '    For i = 0 To 31\n'
         '        s = Replace(s, ChrW(i), "\\x" & Right("0" & Hex(i), 2))\n'
         '    Next\n'
         '    Plain = s\n'
         'End Function\n')


def statement(case: dict[str, str], escape: bool = False) -> str:
    call = f"({case['arguments']})" if case['arguments'] else ''
    before = f"{case['setup']}: " if case['setup'] else ''
    shown = 'Plain(Show(v))' if escape else 'Show(v)'
    return (f"Err.Clear: v = Empty: {before}v = f.{case['function']}{call}\n"
            f"    If Err.Number <> 0 Then out = out & \"E\" & Err.Number & \"{SEPARATOR}\" Else out = out & {shown} & \"{SEPARATOR}\"")


def module(cases: list[dict[str, str]], typed: bool, escape: bool = False) -> str:
    kind = 'WorksheetFunction' if typed else 'Object'
    body = '\n'.join(f'    {statement(case, escape)}' for case in cases)
    return (MORE['helper'] + '\n' + MORE['setup'] + '\n' + (PLAIN if escape else '') +
            'Public Function One() As String\n'
            f'    Dim wb As Object, ws As Object, f As {kind}, v As Variant, k As Variant, out As String\n'
            '    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)\n'
            '    Fill ws\n    Set f = Application.WorksheetFunction\n    On Error Resume Next\n' + body + '\n'
            '    On Error GoTo 0\n    wb.Close False\n    One = out\nEnd Function\n')


def refused(answer: str) -> bool:
    """Whether the model's answer is its refusal: a call it says is not implemented."""
    return answer.startswith('model VBAUnsupportedError')


def model_answers(cases: list[dict[str, str]], typed: bool, escape: bool = False) -> list[str]:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    answers: list[str] = []
    for start in range(0, len(cases), BATCH):
        batch = cases[start:start + BATCH]
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        try:
            app.add_module(module(batch, typed, escape), name='Round')
            text = str(app.run('Round.One'))
            answers += text.split(SEPARATOR)[:len(batch)]
        except VBAError:
            # One call the model refuses outright takes the batch with it: ask each alone.
            for case in batch:
                single = ExcelApplication(sandbox=False)
                single.add_workbook()
                try:
                    single.add_module(module([case], typed, escape), name='Round')
                    answers.append(str(single.run('Round.One')).split(SEPARATOR)[0])
                except VBAError as failed:
                    answers.append(f'model {type(failed).__name__}: {str(failed)[:120]}')
    return answers


def digest(cases: list[dict[str, str]]) -> str:
    text = json.dumps([[case['function'], case['setup'], case['arguments'], case['model_late'], case['model_typed']] for case in cases])
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:16]


def predict(seed: int) -> None:
    path = OUT / f'round_{seed}.json'
    assert not path.exists(), f'{path} exists: a round is predicted once'
    cases = calls(seed)
    for tag, typed in (('late', False), ('typed', True)):
        for case, answer in zip(cases, model_answers(cases, typed, seed >= ESCAPES_FROM), strict=True):
            case[f'model_{tag}'] = answer
    OUT.mkdir(parents=True, exist_ok=True)
    record = {'seed': seed, 'predicted_at': datetime.now(timezone.utc).isoformat(), 'predictions_sha256_16': digest(cases), 'cases': cases}
    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print('predicted', len(cases), 'cases over', len({case['function'] for case in cases}), 'functions, digest', record['predictions_sha256_16'])


def measure(seed: int) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    path = OUT / f'round_{seed}.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    cases = record['cases']
    assert digest(cases) == record['predictions_sha256_16'], 'the predictions changed after they were written'
    lock = rounds._lock()
    owner = rounds.take_lock(lock, 'Verification 02')
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True):
            for tag, typed in (('late', False), ('typed', True)):
                todo = [case for case in cases if str(case.get(tag, 'no answer')).startswith('no answer')]
                for start in range(0, len(todo), BATCH):
                    batch = todo[start:start + BATCH]
                    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                        excel.new_document()
                        done = excel.run_vba(module(batch, typed, seed >= ESCAPES_FROM), 'One', (), timeout=180.0, line_numbers=False)
                    answers = str(done.value).split(SEPARATOR)[:len(batch)] if done.ok else []
                    if len(answers) != len(batch):
                        answers = []
                        for case in batch:
                            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                                excel.new_document()
                                alone = excel.run_vba(module([case], typed, seed >= ESCAPES_FROM), 'One', (), timeout=60.0, line_numbers=False)
                            answers.append(str(alone.value).split(SEPARATOR)[0] if alone.ok else f'no answer: {str(alone)[:300]}')
                    for case, answer in zip(batch, answers, strict=True):
                        case[tag] = answer
                    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
                    print(tag, start + len(batch), 'of', len(todo), flush=True)
    finally:
        rounds.release_lock(lock, owner)
    if any(str(case[tag]).startswith('no answer') for case in cases for tag in ('late', 'typed')):
        print('cases with no answer: run again to ask them again')
        return
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    answered = [(case, tag) for case in cases for tag in ('late', 'typed') if not refused(case[f'model_{tag}'])]
    differ = [(case, tag) for case, tag in answered if case[tag] != case[f'model_{tag}']]
    sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[union-attr]
    for case, tag in differ:
        print(tag, case['function'], case['setup'], '|', case['arguments'], '| Excel', case[tag], '| model', case[f'model_{tag}'])
    print('wrong', len(differ), 'of', len(answered), 'answered;', 2 * len(cases) - len(answered), 'refused')


if __name__ == '__main__':
    action, number = sys.argv[1], int(sys.argv[2])
    predict(number) if action == 'predict' else measure(number)
