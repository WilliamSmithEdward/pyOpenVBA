"""A Currency, a Date, a Single, a Decimal and ranges given to each settable property of Range, Font, Interior, Border, Borders, Style
and Characters, typed and late bound: the error, and what the property then reads (scripts/measure_given_values.py,
tests/fixtures/given_value_rounds.json).

Each round's values were new, with the model's answers written down first; rules were fixed against each round up to the fifth, so
those are now only regressions. The sixth ran on the fifth's code, on enum numbers past what was measured, separate areas, whole rows
and a sample of whole columns. A row goes VERIFIED only with no difference in both, and with no reported gap in reach. The scores name
what missed, and those members' cases are left out here: their rows are PARTIAL.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/given_value_rounds.json').read_text(encoding='utf-8'))
MEMBERS = sorted({key for one in RECORD['rounds'].values() for key in one['answers']})


def _source(owner: str, member: str, typed: bool, round_: dict[str, Any]) -> str:
    kind = owner if typed else 'Object'
    lines = [RECORD['helper'], 'Public Function Probe() As String', f'    Dim wb As Object, ws As Object, o As {kind}, x As Variant, v As Variant, s As String, e As Long, sn As String, n As Long',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next']
    for value in round_['kinds']:
        lines += ['    n = n + 1: sn = "Given" & n: wb.Styles.Add sn: Err.Clear', '    Set ws = wb.Worksheets.Add',
                  f'    ws.Range("B2").Value = "abc": {round_["cells"]}',
                  f'    Set o = {RECORD["reach"][owner]}', '    x = Empty: Err.Clear', f'    {value}', f'    o.{member} = x', '    e = Err.Number: Err.Clear',
                  f'    v = Empty: v = o.{member}', '    s = s & "@@" & Note(e, Empty) & "|" & Note(Err.Number, v): Err.Clear']
    lines += ['    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']
    return '\n'.join(lines)


def _differences(key: str, name: str, binding: str) -> list[tuple[str, str, str]]:
    round_ = RECORD['rounds'][name]
    if key not in round_['answers']:
        return []
    owner, member = key.split('.')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(owner, member, binding == 'typed', round_), name='Given')
    got = str(app.run('Given.Probe')).split('@@')
    wrong = [(value, want, one) for value, want, one in zip(round_['kinds'], round_['answers'][key][binding], got, strict=True) if want != one]
    # A column's width is in the display's own steps: only the error is compared (docs/host_completeness.md).
    if member == 'ColumnWidth':
        wrong = [one for one in wrong if one[1].split('|')[0] != one[2].split('|')[0]]
    return wrong


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('name', sorted(RECORD['rounds']))
@pytest.mark.parametrize('key', MEMBERS)
def test_each_member_as_excel_answered(key: str, name: str, binding: str) -> None:
    assert _differences(key, name, binding) == []


def test_the_scores_name_the_misses_and_none_is_replayed() -> None:
    fifth, sixth = RECORD['scores']['fifth'], RECORD['scores']['sixth']
    assert set(fifth['wrong']) == {'Border.TintAndShade', 'Borders.TintAndShade'}
    assert fifth['comparisons'] == 2860 and sixth['comparisons'] == 3888
    assert not set(sixth['wrong']) & set(MEMBERS)
