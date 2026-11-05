"""An Interior over fills from cell styles, alone and with direct fills, read and set, typed and late bound: in books made in the run, and in books
Excel saved that both open (tests/fixtures/style_fills.json, tests/fixtures/style_fills/). A conditional format does not reach an Interior, only its
DisplayFormat. The model's answers were written down first; the first round was fixed against and the second ran on the final code. Their only
differences left are pattern colours beside a gradient loaded from a file, whose rows are PARTIAL and whose calls are left out here.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

HERE = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((HERE / 'style_fills.json').read_text(encoding='utf-8'))


def _source(group: dict[str, Any], typed: bool, path: str) -> str:
    kind = 'Range' if typed else 'Object'
    lines = [RECORD['helper'], 'Public Function Probe() As String', f'    Dim wb As Object, ws As Object, r As {kind}, v As Variant, s As String, o As Object, e As Long']
    fresh = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Styles wb', *[f'    {line}' for line in group['dress'] if line]]
    opened = [f'    Set wb = Workbooks.Open("{path}")', '    Set ws = wb.Worksheets(1)']
    for case in group['cases']:
        lines += opened if path else fresh
        lines += ['    On Error Resume Next', f'    Err.Clear: Set r = ws.Range("{case["area"]}")']
        if case['kind'] == 'read' and case['member'] == 'Gradient':
            lines += ['    Set o = Nothing: v = Empty: Err.Clear', '    Set o = r.Interior.Gradient', '    If Err.Number = 0 Then Set v = o', '    s = s & "@@" & Note(Err.Number, v): Err.Clear']
        elif case['kind'] == 'read':
            lines += ['    v = Empty: Err.Clear', f'    v = r.Interior.{case["member"]}', '    s = s & "@@" & Note(Err.Number, v): Err.Clear']
        else:
            lines += [f'    r.Interior.{case["member"]} = {case["value"]}', '    e = Err.Number: Err.Clear', '    s = s & "@@" & Note(e, Empty) & "|" & Note(0, r.Interior.Pattern) & "|" & Note(0, r.Interior.Color) & "|" & Note(0, r.Interior.ColorIndex): Err.Clear']
        lines += ['    On Error GoTo 0', '    wb.Close False']
    lines += ['    Probe = Mid(s, 3)', 'End Function', '']
    return '\n'.join(lines)


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('name', sorted(RECORD['rounds']))
def test_each_call_as_excel_answered(name: str, binding: str) -> None:
    wrong: list[tuple[Any, ...]] = []
    for group in RECORD['rounds'][name]:
        book = group.get('book_file')
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(_source(group, binding == 'typed', str(HERE / 'style_fills' / book) if book else ''), name='Fills')
        got = str(app.run('Fills.Probe')).split('@@')
        wrong += [(book, case['area'], case['member'], case[binding], one) for case, one in zip(group['cases'], got, strict=True)
                  if one != case[binding] and not case.get('known')]
    assert wrong == []


def test_the_scores_when_each_round_counted() -> None:
    assert RECORD['scores']['second']['differ'] == {'saved read PatternTintAndShade': 2}
