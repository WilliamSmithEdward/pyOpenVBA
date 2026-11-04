"""When a procedure is refused, against Excel with Compile On Demand at its default (scripts/measure_vba_compile_timing.py).

A line that does not compile stops only a run that enters its procedure, first at that entry, whatever ran before it and
under any On Error.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

ANSWERS = json.loads((Path(__file__).parent / 'fixtures/vba_compile_timing.json').read_text(encoding='utf-8'))['answers']
REFUSED_HERE = {'read_only': 'ws.Index = 1', 'argument_count': 'v = r.Offset(1, 2, 3)', 'named_argument': 'v = r.Offset(Foo:=1)', 'sub_as_value': 'v = ws.Activate',
                'unknown_member': 'v = ws.Nope', 'undeclared_explicit': 'v = NotDeclared', 'project_function_count': 'v = Good(1, 2)'}
GOOD = 'Function Good() As String\n    Good = "ran"\nEnd Function\n'


def bad(line: str) -> str:
    return f'Function Bad() As String\n    Dim ws As Worksheet, r As Range, v As Variant\n    Set ws = ActiveSheet: Set r = ws.Range("A1")\n    {line}\n    Bad = "bad ran"\nEnd Function\n'


def workbook(*modules: tuple[str, str]) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    for name, source in modules:
        app.add_module('Option Explicit\n' + source, name=name)
    return app


@pytest.mark.parametrize(('family', 'line'), REFUSED_HERE.items())
def test_a_bad_line_stops_only_its_own_procedure(family: str, line: str) -> None:
    assert ANSWERS[f'{family} neighbour_run']['value'] == 'ran'
    assert ANSWERS[f'{family} itself_run']['dialogs'][0]['kind'] == 'compile-error'
    app = workbook(('Harness', GOOD + bad(line)))
    assert app.run('Harness.Good') == 'ran'
    with pytest.raises(VBACompileError):
        app.run('Harness.Bad')


def test_an_other_module_is_refused_only_when_called() -> None:
    assert ANSWERS['other_module_not_called']['value'] == 'ran'
    assert ANSWERS['other_module_called']['dialogs'][0]['kind'] == 'compile-error'
    app = workbook(('Harness', GOOD + 'Function CallsBad() As String\n    CallsBad = Bad()\nEnd Function\n'), ('Other', bad('v = NotDeclared')))
    assert app.run('Harness.Good') == 'ran'
    with pytest.raises(VBACompileError):
        app.run('Harness.CallsBad')


def test_on_error_does_not_catch_a_refusal() -> None:
    assert ANSWERS['on_error_in_caller']['dialogs'][0]['kind'] == 'compile-error'
    assert ANSWERS['on_error_other_module']['dialogs'][0]['kind'] == 'compile-error'
    app = workbook(('Harness', 'Function Guarded() As String\n    On Error Resume Next\n    Guarded = Bad()\n    Guarded = "went on"\nEnd Function\n' + bad('v = NotDeclared')))
    with pytest.raises(VBACompileError):
        app.run('Harness.Guarded')
