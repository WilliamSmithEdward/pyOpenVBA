"""Worksheet functions tested on calls the refusal rules were not made from. The sample was seeded, the model's answers were
written down and their digest recorded before anything was measured, and only then did live Excel answer
(tests/fixtures/worksheet_function_heldout.json, scripts/measure_worksheet_function_heldout.py).

The fixture keeps the cases of the functions with no difference at all, last digits included. Its `scores` say how each round
went over every function it sampled, those that missed included: they stay PARTIAL, and fixing one against these cases would
use them up, so it needs new cases.
"""
from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'worksheet_function_heldout.json').read_text(encoding='utf-8'))
HELPER: str = json.loads((FIXTURES / 'worksheet_function_kinds.json').read_text(encoding='utf-8'))['helper']
BY_FUNCTION: dict[str, list[dict[str, str]]] = defaultdict(list)
for _case in RECORD['cases']:
    BY_FUNCTION[_case['function']].append(_case)


def _answers(calls: list[str], typed: bool) -> list[str]:
    kind = 'WorksheetFunction' if typed else 'Object'
    lines = ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, f As {kind}, v As Variant, x As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Fill ws', '    Set f = Application.WorksheetFunction', '    On Error Resume Next']
    lines += [f'    v = Empty: x = Empty: {call}: s = s & "@@" & Note(Err.Number, v): Err.Clear' for call in calls]
    source = HELPER + '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', ''])
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Functions')
    return str(app.run('Functions.Probe')).split('@@')


@pytest.mark.parametrize('name', sorted(BY_FUNCTION))
def test_each_function_as_excel_answered(name: str) -> None:
    cases = BY_FUNCTION[name]
    calls = [case['call'] for case in cases]
    late, typed = _answers(calls, False), _answers(calls, True)
    assert [(case['call'], one) for case, one in zip(cases, late, strict=True) if one != case['late']] == []
    assert [(case['call'], two) for case, two in zip(cases, typed, strict=True) if two != case['typed']] == []


def test_the_scores_count_the_misses() -> None:
    for score in RECORD['scores'].values():
        assert score['comparisons'] and score['functions_clean'] <= score['functions']
        assert len(score['missed']) == score['functions'] - score['functions_clean']
