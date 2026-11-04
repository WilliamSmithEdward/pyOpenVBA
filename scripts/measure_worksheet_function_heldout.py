"""Ask live Excel the calls of tests/fixtures/worksheet_function_heldout.json again, typed and late bound, and say which answers
have changed. Writes nothing: a held-out case is only held out until the model has been fixed against it, so new rounds need new
calls, with the model's answers written down first.

    python scripts/measure_worksheet_function_heldout.py
"""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
RECORD = json.loads((FIXTURES / 'worksheet_function_heldout.json').read_text(encoding='utf-8'))
HELPER = json.loads((FIXTURES / 'worksheet_function_kinds.json').read_text(encoding='utf-8'))['helper']


def source(calls: list[str], typed: bool) -> str:
    kind = 'WorksheetFunction' if typed else 'Object'
    lines = ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, f As {kind}, v As Variant, x As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Fill ws', '    Set f = Application.WorksheetFunction', '    On Error Resume Next']
    lines += [f'    v = Empty: x = Empty: {call}: s = s & "@@" & Note(Err.Number, v): Err.Clear' for call in calls]
    return HELPER + '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', ''])


def main() -> None:
    by_function: dict[str, list[dict[str, str]]] = defaultdict(list)
    for case in RECORD['cases']:
        by_function[case['function']].append(case)
    changed = 0
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        for name, cases in sorted(by_function.items()):
            calls = [case['call'] for case in cases]
            for tag, typed in (('late', False), ('typed', True)):
                result = excel.run_vba(source(calls, typed), 'Probe', (), timeout=120.0)
                answers = str(result.value).split('@@') if result.ok else ['no answer'] * len(cases)
                for case, answer in zip(cases, answers, strict=True):
                    if answer != case[tag]:
                        changed += 1
                        print(name, tag, case['call'], 'was', case[tag], 'now', answer, flush=True)
    print('changed', changed, 'of', 2 * len(RECORD['cases']))


if __name__ == '__main__':
    main()
