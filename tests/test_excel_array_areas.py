"""Arrays written to ranges through Value, Value2, Formula and FormulaLocal, typed and late bound, on seeded areas, shapes and kinds of item
(tests/fixtures/array_areas.json). A row or column of one repeats across an area, past the array's size each cell is #N/A, and every
second area reads a two-dimensional array with its sizes the other way about. The model's answers were written down first; rules were
fixed against the first two rounds, and the last two ran on the final code with no difference.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/array_areas.json').read_text(encoding='utf-8'))
BATCH = 60


def _statement(given: str) -> list[str]:
    fills = RECORD['fills']
    if given.startswith('2d'):
        _, low, high, left, right, fill = given.split()
        return [f'    ReDim y({low} To {high}, {left} To {right})', f'    For i = {low} To {high}: For j = {left} To {right}: y(i, j) = {fills[int(fill)]}: Next: Next', '    x = y']
    if given.startswith('1d'):
        _, base, count, fill = given.split()
        last = int(base) + int(count) - 1
        return [f'    ReDim z({base} To {last})', f'    For j = {base} To {last}: i = 1: z(j) = {fills[int(fill)]}: Next', '    x = z']
    return [f'    x = {given}']


def _answers(part: list[dict[str, Any]], typed: bool) -> list[str]:
    lines = ['Public Function Probe() As String', '    Dim wb As Object, ws As Object, r As ' + ('Range' if typed else 'Object') + ', x As Variant, y() As Variant, z() As Variant, s As String, i As Long, j As Long',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error Resume Next']
    for case in part:
        lines += ['    ws.Cells.Clear: Erase y: Erase z: x = Empty', f'    {RECORD["cells"]}', *_statement(case['given']), '    Err.Clear', f'    Set r = ws.Range("{case["area"]}")', f'    r.{case["member"]} = x', '    s = s & "@@e" & Err.Number & " " & Grid(ws): Err.Clear']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['helper'] + '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']), name='Arrays')
    return str(app.run('Arrays.Probe')).split('@@')


@pytest.mark.parametrize('typed', [False, True], ids=['late', 'typed'])
@pytest.mark.parametrize('name', sorted(RECORD['rounds']))
def test_each_write_as_excel_answered(name: str, typed: bool) -> None:
    cases = [case for case in RECORD['rounds'][name] if case['typed'] == typed]
    wrong: list[tuple[str, str, str, str, str]] = []
    for start in range(0, len(cases), BATCH):
        part = cases[start:start + BATCH]
        wrong += [(case['area'], case['member'], case['given'], case['answer'], got)
                  for case, got in zip(part, _answers(part, typed), strict=True) if got != case['answer']]
    assert wrong == []


def test_the_scores_when_each_round_counted() -> None:
    assert [RECORD['scores'][name]['differ'] for name in ('first', 'second', 'third', 'fourth')] == [31, 0, 0, 0]
