"""Formula and FormulaLocal read back after each formula is written through Formula, Formula2, FormulaLocal or FormulaArray: spills and their ghost cells, @ and SINGLE,
LET and LAMBDA, tables, names defined after the formula and names never defined, typed and late bound, in books made in the run and books Excel saved
(tests/fixtures/formula_reads.json, tests/fixtures/formula_reads/). Each round's formulas were new, with the model's answers written down first, and the two here ran on the
final code with no difference. Formula2 and FormulaArray writes of a table reference are set aside: they report a gap, and their rows are PARTIAL.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

HERE = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((HERE / 'formula_reads.json').read_text(encoding='utf-8'))
READERS = ('Formula', 'FormulaLocal')


def _statements(group: dict[str, Any]) -> list[str]:
    lines = ['    Base ws', '    ws.Range("J1").Value = "a": ws.Range("J2").Value = 7: ws.Range("J3").Value = 8', '    ws.ListObjects.Add(xlSrcRange, ws.Range("J1:J3"), , xlYes).Name = "Tab"']
    for write in group['writes']:
        text = write['text'].replace('"', '""')
        lines.append(f'    ws.Range("{write["cell"]}").{write["writer"]} = "{text}"')
    lines += ['    wb.Names.Add "Later", "=Sheet1!$A$1:$A$3"', '    wb.Names.Add "Soon", "=Sheet1!$A$2:$A$4"']
    return lines


def _source(group: dict[str, Any], typed: bool, path: str) -> str:
    kind = 'Range' if typed else 'Object'
    lines = ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, r As {kind}, v As Variant, s As String', '    Application.DisplayAlerts = False']
    if path:
        lines += [f'    Set wb = Workbooks.Open("{path}")', '    Set ws = wb.Worksheets(1)', '    On Error Resume Next']
    else:
        lines += ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"', '    On Error Resume Next', *_statements(group)]
    for area in group['reads']:
        for member in READERS:
            lines += [f'    Set r = ws.Range("{area}")', '    v = Empty: Err.Clear', f'    v = r.{member}', '    s = s & "@@" & Show(Err.Number, v): Err.Clear']
    lines += ['    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', '']
    return RECORD['helper'] + '\n'.join(lines)


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('name', sorted(RECORD['rounds']))
def test_each_read_as_excel_answered(name: str, binding: str) -> None:
    wrong: list[tuple[Any, ...]] = []
    for group in RECORD['rounds'][name]:
        book = group.get('book_file')
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(_source(group, binding == 'typed', str(HERE / 'formula_reads' / book) if book else ''), name='Reads')
        got = str(app.run('Reads.Probe')).split('@@')
        wanted = group[binding]
        wrong += [(book, one, two) for one, two in zip(wanted, got, strict=True) if one != two]
    assert wrong == []


def test_the_scores_when_each_round_counted() -> None:
    assert [RECORD['scores'][name]['differ'] for name in ('third', 'fourth')] == [0, 0]
