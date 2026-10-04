"""Native VBA module qualification and ByRef conformance."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter import Interpreter

MEASURED = json.loads((Path(__file__).parent / 'fixtures/module_namespaces.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('probe', MEASURED['probes'], ids=[probe['name'] for probe in MEASURED['probes']])
def test_native_module_namespace_operations(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('kind,source,expression', [
    ('function', 'Private Function Hidden() As Long\nHidden = 3\nEnd Function', 'Helpers.Hidden()'),
    ('variable', 'Private Hidden As Long', 'Helpers.Hidden'),
    ('constant', 'Private Const Hidden As Long = 3', 'Helpers.Hidden'),
])
def test_namespace_does_not_expose_private_members(kind: str, source: str, expression: str) -> None:
    interpreter = Interpreter()
    interpreter.add_module(source, name='Helpers')
    interpreter.add_module('Function Probe() As Long\nProbe = ' + expression + '\nEnd Function', name='Caller')
    with pytest.raises(VBACompileError, match='private'):
        interpreter.run('Caller.Probe')
    assert kind
