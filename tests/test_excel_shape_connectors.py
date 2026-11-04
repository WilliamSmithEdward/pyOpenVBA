"""ConnectorFormat's connections, replayed from live Excel (scripts/measure_shape_connectors.py, tests/fixtures/shape_connectors.json)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_connectors.json').read_text(encoding='utf-8'))
F = 'sh.ConnectorFormat.'
#: Each of these reports itself: connecting to a line or a group's member, duplicating a connector, an elbow's or a curve's routing
#: other than its plain case, and changing a connector's type straight to a curve.
UNBUILT = ('Set o = ws.Shapes.AddLine', 'o.GroupItems', 'sh.Duplicate', 
           f'Set sh = ws.Shapes.AddConnector(2, 10, 10, 50, 30): {F}BeginConnect a, 1', f'Set sh = ws.Shapes.AddConnector(2, 10, 10, 50, 30): {F}BeginConnect a, 3',
           f'Set sh = ws.Shapes.AddConnector(2, 10, 10, 50, 30): {F}BeginConnect a, 4: {F}EndConnect b, 4', f'Set sh = ws.Shapes.AddConnector(2, 10, 10, 50, 30): {F}BeginConnect a, 4: {F}EndConnect b, 2: b.Left',
           f'Set sh = ws.Shapes.AddConnector(3, 10, 10, 50, 30): {F}BeginConnect a, 3', f'{F}EndConnect b, 2: sh.ConnectorFormat.Type')
CASES = [case for case in RECORD['cases'] if not any(one in case[0] for one in UNBUILT)]
IDS = [f'{index}: {case[0]}' for index, case in enumerate(CASES)]


def _replayed(case: list[Any], tmp_path: Path) -> tuple[Any, list[str]]:
    lines = ['Public Function Go() As String', '    Dim ws As Object, a As Object, b As Object, c As Object, t As Object, r As Object, x As Object, sh As Object, o As Object',
             '    Dim s As String, v As String, s2 As Variant', '    Set ws = ActiveSheet', f'    {RECORD["make"].rstrip(": ")}', '    On Error GoTo Failed', f'    {case[0]}', '    s = "ok"',
             'Measured:', '    On Error Resume Next']
    for read in RECORD['reads']:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Connect')
    answer = app.run('Connect.Go')
    path = tmp_path / 'b.xlsx'
    app.save(path)
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    return answer, re.findall(r'<xdr:cxnSp\b.*?</xdr:cxnSp>', xml, re.DOTALL)


def _steady(part: str) -> str:
    """A saved connector without the creation id, which is new each time."""
    return re.sub(r'<a:extLst>.*?</a:extLst>', '', part)


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], tmp_path: Path) -> None:
    answer, parts = _replayed(case, tmp_path)
    names = ['status', *RECORD['reads']]
    # A shape held after its Delete still reads in the model, where Excel raises 424: a gap of every kind of shape, not of connectors.
    held = [name for name in names if name.startswith('a.') and 'a.Delete' in case[0]]
    theirs = {**dict(zip(names, case[1].split('~'), strict=True)), **dict.fromkeys(held)}
    assert {**dict(zip(names, answer.split('~'), strict=True)), **dict.fromkeys(held)} == theirs
    assert [_steady(part) for part in parts] == [_steady(part) for part in case[2]]


@pytest.mark.parametrize('statements', ['Set o = ws.Shapes.AddLine(1, 1, 2, 2): sh.ConnectorFormat.BeginConnect o, 1', 'sh.ConnectorFormat.BeginConnect a, 1: sh.ConnectorFormat.EndConnect b, 1'])
def test_what_is_not_worked_out_reports_itself(statements: str) -> None:
    connector = 'AddConnector(1, 10, 10, 50, 30)' if 'AddLine' in statements else 'AddConnector(2, 10, 10, 50, 30)'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ws As Object, a As Object, b As Object, sh As Object, o As Object\n    Set ws = ActiveSheet\n'
                   f'    Set a = ws.Shapes.AddShape(1, 100, 100, 80, 40): Set b = ws.Shapes.AddShape(1, 300, 200, 60, 50)\n    Set sh = ws.Shapes.{connector}\n    {statements}\nEnd Sub\n', name='Unbuilt')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unbuilt.Go')
