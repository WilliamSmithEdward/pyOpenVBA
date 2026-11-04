"""The rest of the worksheet functions, each with no argument and each kind of value for each argument of a call Excel
answered, typed and late bound, replayed from live Excel (scripts/measure_worksheet_function_kinds.py,
tests/fixtures/worksheet_function_kinds.json).

The functions in VERIFIED answer every call as Excel did; their checklist rows say VERIFIED. The rest of the fixture is
what is still to be matched.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/worksheet_function_kinds.json').read_text(encoding='utf-8'))
VERIFIED: list[str] = json.loads(
    (Path(__file__).parent / 'fixtures/worksheet_function_kinds_verified.json').read_text(encoding='utf-8'))


def _source(calls: list[str], typed: bool) -> str:
    """The probe, as the script measured it: f.Name() does not compile typed, so answers no compile there."""
    kind = 'WorksheetFunction' if typed else 'Object'
    lines = ['Public Function Probe() As String', f'    Dim wb As Object, ws As Object, f As {kind}, v As Variant, x As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Fill ws', '    Set f = Application.WorksheetFunction', '    On Error Resume Next']
    for call in calls:
        lines.append('    s = s & "@@no compile"' if typed and call.endswith('()')
                     else f'    v = Empty: x = Empty: {call}: s = s & "@@" & Note(Err.Number, v): Err.Clear')
    return RECORD['helper'] + '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', ''])


def _answers(calls: list[str], typed: bool) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(calls, typed), name='Functions')
    return str(app.run('Functions.Probe')).split('@@')


@pytest.mark.parametrize('name', VERIFIED)
def test_each_function_as_excel_answered(name: str) -> None:
    cases: list[list[str]] = RECORD['functions'][name]
    calls = [case[0] for case in cases]
    late, typed = _answers(calls, False), _answers(calls, True)
    assert [(case[0], one) for case, one in zip(cases, late, strict=True) if one != case[1]] == []
    assert [(case[0], two) for case, two in zip(cases, typed, strict=True) if two != case[2]] == []
