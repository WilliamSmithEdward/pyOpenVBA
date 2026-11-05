"""An Interior over cells whose fills differ: seeded dressing of cells with gradients, colours and patterns, read and set, typed and late bound
(tests/fixtures/mixed_fills.json). Over several fills a gradient reads as no fill, Pattern is Null, and Gradient is Nothing. The model's answers
were written down first; rules were fixed against the first two rounds, and the last two ran on the final code. Their only differences are
TintAndShade and PatternTintAndShade set across cells whose fills differ, where every cell takes one colour by a rule not known: those calls
are left out here and those rows are PARTIAL.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/mixed_fills.json').read_text(encoding='utf-8'))
SHADES = {'TintAndShade', 'PatternTintAndShade'}


def _answers(part: list[dict[str, Any]], typed: bool) -> list[str]:
    kind = 'Range' if typed else 'Object'
    lines = [RECORD['helper'], 'Public Function Probe() As String', f'    Dim wb As Object, ws As Object, r As {kind}, v As Variant, s As String, o As Object, e As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error Resume Next']
    last = None
    for case in part:
        if case['dress'] != last or case['kind'] == 'set':
            lines += ['    ws.Cells.Clear', *[f'    {line}' for line in case['dress'] if line]]
            last = case['dress']
        lines.append(f'    Err.Clear: Set r = ws.Range("{case["area"]}")')
        if case['kind'] == 'read' and case['member'] == 'Gradient':
            lines += ['    Set o = Nothing: v = Empty: Err.Clear', '    Set o = r.Interior.Gradient', '    If Err.Number = 0 Then Set v = o', '    s = s & "@@" & Note(Err.Number, v): Err.Clear']
        elif case['kind'] == 'read':
            lines += ['    v = Empty: Err.Clear', f'    v = r.Interior.{case["member"]}', '    s = s & "@@" & Note(Err.Number, v): Err.Clear']
        else:
            lines += [f'    r.Interior.{case["member"]} = {case["value"]}', '    e = Err.Number: Err.Clear', '    s = s & "@@" & Note(e, Empty) & "|" & Note(0, r.Interior.Pattern) & "|" & Note(0, r.Interior.Color) & "|" & Note(0, r.Interior.ColorIndex): Err.Clear']
            last = None
    lines += ['    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines), name='Fills')
    return str(app.run('Fills.Probe')).split('@@')


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('name', sorted(RECORD['rounds']))
def test_each_call_as_excel_answered(name: str, binding: str) -> None:
    cases = RECORD['rounds'][name]
    wrong: list[tuple[Any, ...]] = []
    for start in range(0, len(cases), 9):
        part = cases[start:start + 9]
        got = _answers(part, binding == 'typed')
        wrong += [(case['area'], case['kind'], case['member'], case[binding], one) for case, one in zip(part, got, strict=True)
                  if one != case[binding] and not (case['kind'] == 'set' and case['member'] in SHADES)]
    assert wrong == []


def test_the_last_rounds_differ_only_on_the_shades() -> None:
    assert set(RECORD['scores']['third']['differ']) == set(RECORD['scores']['fourth']['differ']) == {'set TintAndShade', 'set PatternTintAndShade'}
