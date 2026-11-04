"""The VERIFIED standard (docs/host_completeness.md) for the rest of the worksheet functions: each with no argument and each
kind of value for each argument of a call Excel answered, typed and late bound.

    python scripts/measure_worksheet_function_kinds.py [seconds]

writes tests/fixtures/worksheet_function_kinds.json, function by function, which tests/test_excel_worksheet_function_kinds.py
replays both ways. Of a function taking any number of arguments, those of its longest measured call are varied. A call
with an argument left out does not compile typed where one is required, so only late binding answers it.
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))

from measure_worksheet_function_types import FIXTURES, HELPER, split_arguments  # noqa: E402
from pyvbaharness import ExcelSession, HarnessConfig  # noqa: E402

from pyopenvba.interpreter._type_library import find_class  # noqa: E402

OUT = ROOT / 'tests/fixtures/worksheet_function_kinds.json'
#: The functions measured elsewhere with every kind of value (tests/fixtures/worksheet_function_args.json).
DONE = {'VLookup', 'HLookup', 'Match', 'Index', 'Sum', 'CountIf', 'SumIf', 'Max', 'Min', 'Round', 'Count', 'CountA', 'Average', 'Subtotal'}
KINDS = ['x = Null', 'x = Empty', 'x = True', 'x = False', 'x = "x"', 'x = ""', 'x = "2"', 'x = 1.5', 'x = 0', 'x = -1', 'x = 99', 'x = 1E+300',
         'Set x = ws.Range("A2")', 'Set x = ws.Range("B2")', 'Set x = ws.Range("A1:A3")', 'Set x = ws.Range("Z9")', 'x = Array(1, 2)', 'x = CVErr(xlErrNA)',
         'Set x = Nothing', 'Set x = ws']
def functions() -> dict[str, list[str]]:
    """Each function to measure and the arguments of its longest measured call."""
    library = find_class('excel', 'worksheetfunction')
    assert library is not None
    calls: dict[str, list[list[str]]] = {}
    for name in ('worksheet_functions.json', 'worksheet_functions_more.json'):
        for case in json.loads((FIXTURES / name).read_text(encoding='utf-8'))['cases']:
            calls.setdefault(case['function'], []).append(split_arguments(case['arguments']))
    return {name: max(found, key=len) for name, found in sorted(calls.items()) if name not in DONE and library.members.get(name.lower())}


def statements(name: str, arguments: list[str]) -> list[str]:
    """The calls of one function: its measured call, none of its arguments, and each kind in each place."""
    found = [f'v = f.{name}({", ".join(arguments)})', f'v = f.{name}()']
    for index in range(len(arguments)):
        rest = [*arguments[:index], 'x', *arguments[index + 1:]]
        found += [f'{kind}: v = f.{name}({", ".join(rest)})' for kind in KINDS]
    return found


def source(calls: list[str], typed: bool) -> str:
    """A probe answering each call joined by @@. ``f.Name()`` does not compile typed, so answers no compile there."""
    fill = json.loads((FIXTURES / 'worksheet_functions.json').read_text(encoding='utf-8'))['setup']
    kind = 'WorksheetFunction' if typed else 'Object'
    body = HELPER + fill.split('Private Function Item')[0]
    parts = [calls[at:at + 200] for at in range(0, len(calls), 200)]
    for number, part in enumerate(parts):
        lines = [f'Private Function Part{number}(f As {kind}, ws As Object) As String', '    Dim v As Variant, x As Variant, s As String', '    On Error Resume Next']
        for call in part:
            if typed and call.endswith('()'):
                lines.append('    s = s & "@@no compile"')
            else:
                lines.append(f'    v = Empty: x = Empty: {call}: s = s & "@@" & Note(Err.Number, v): Err.Clear')
        lines += [f'    Part{number} = Mid(s, 3)', 'End Function', '']
        body += '\n'.join(lines) + '\n'
    joined = ' & "@@" & '.join(f'Part{number}(f, ws)' for number in range(len(parts)))
    return body + (f'Public Function Probe() As String\n    Dim wb As Object, ws As Object, f As {kind}, s As String\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                   f'    Set ws = wb.Worksheets(1)\n    Fill ws\n    Set f = Application.WorksheetFunction\n    s = {joined}\n    wb.Close False\n    Probe = s\nEnd Function\n')


def main() -> None:
    """Measures function by function, keeping what it has, so it can be run again until "left 0"; a function that does not
    answer within its time is recorded as such and not measured."""
    fill = json.loads((FIXTURES / 'worksheet_functions.json').read_text(encoding='utf-8'))['setup']
    record = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'functions': {}, 'unanswered': []}
    record['helper'] = HELPER + fill.split('Private Function Item')[0]
    every = functions()
    started = time.monotonic()
    budget = float(sys.argv[1]) if len(sys.argv) > 1 else 500.0
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        for name, arguments in every.items():
            if name in record['functions'] or name in record['unanswered']:
                continue
            if time.monotonic() - started > budget:
                break
            calls = statements(name, arguments)
            late = excel.run_vba(source(calls, False), 'Probe', (), timeout=40.0)
            typed = excel.run_vba(source(calls, True), 'Probe', (), timeout=40.0) if late.ok and not late.dialogs else late
            if not (late.ok and typed.ok) or late.dialogs or typed.dialogs:
                record['unanswered'].append(name)
                print('unanswered', name, late.outcome, typed.outcome, str(typed.message)[:80], flush=True)
            else:
                one, two = str(late.value).split('@@'), str(typed.value).split('@@')
                assert len(one) == len(two) == len(calls), (name, len(one), len(two), len(calls))
                record['functions'][name] = [[call, a, b] for call, a, b in zip(calls, one, two, strict=True)]
            OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    left = [name for name in every if name not in record['functions'] and name not in record['unanswered']]
    print('measured', len(record['functions']), 'unanswered', len(record['unanswered']), 'left', len(left), flush=True)


if __name__ == '__main__':
    main()
