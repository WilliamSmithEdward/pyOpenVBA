"""XLookup, Sort, Unique and YearFrac on rounds of seeded calls, with the model's answers written down first
(tests/fixtures/worksheet_function_four.json). Rules were fixed against each round before the next, so the rounds are now only
regressions. Their scores say how each did when it counted. YearFrac is VERIFIED: its last two rounds, new values and calls
weighted to arrays, had no difference. The other three stay PARTIAL. KNOWN lists what the model still answers differently, all of
it an array given where one value goes, and none of it YearFrac's.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/worksheet_function_four.json').read_text(encoding='utf-8'))
KNOWN: set[str] = set(json.loads((Path(__file__).parent / 'fixtures/worksheet_function_four_known.json').read_text(encoding='utf-8')))
BATCH = 300


def _answers(calls: list[str], fill: str, typed: bool) -> list[str]:
    kind = 'WorksheetFunction' if typed else 'Object'
    lines = ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, f As {kind}, v As Variant, x As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Fill ws', '    Fresh ws', '    Set f = Application.WorksheetFunction', '    On Error Resume Next']
    lines += [f'    v = Empty: x = Empty: Err.Clear: {call}: s = s & "@@" & Show(Err.Number, v)' for call in calls]
    source = RECORD['helper'] + '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']) + '\n' + fill
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Functions')
    return str(app.run('Functions.Probe')).split('@@')


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('name', ['first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth'])
def test_each_call_as_excel_answered_but_the_known(name: str, binding: str) -> None:
    round_ = RECORD['rounds'][name]
    cases = round_['cases']
    wrong: list[tuple[str, str, str]] = []
    for start in range(0, len(cases), BATCH):
        part = cases[start:start + BATCH]
        got = _answers([case['call'] for case in part], round_['fill'], binding == 'typed')
        wrong += [(case['call'], case[binding], one) for case, one in zip(part, got, strict=True)
                  if one != case[binding] and case['call'] not in KNOWN]
    assert wrong == []


def test_the_scores_say_what_was_missed_when_it_counted() -> None:
    scores = RECORD['scores']
    assert sum(scores['first']['differ'].values()) == 204 and sum(scores['first']['comparisons'].values()) == 1880
    assert scores['second']['differ'] == {'XLookup': 10, 'YearFrac': 0, 'Sort': 4, 'Unique': 22}
    assert [scores[name]['differ']['YearFrac'] for name in ('third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth')] == [125, 101, 1, 1, 0, 0]


def test_year_frac_has_no_known_difference() -> None:
    assert not [call for call in KNOWN if 'YearFrac' in call]


def test_every_known_difference_gives_an_array_or_several_cells_for_one_value() -> None:
    assert not [call for call in KNOWN if 'Array(' not in call and ':I' not in call]
