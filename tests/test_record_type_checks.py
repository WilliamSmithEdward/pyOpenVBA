"""Native nominal VBA record type checking before procedure execution."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError

PROBES: list[dict[str, str]] = json.loads(
    (Path(__file__).parent / 'fixtures/record_type_checks.json').read_text(encoding='utf-8')
)['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_type_checks(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'compile_error' in probe:
        message = probe['compile_error'].split('\n')[-1]
        with pytest.raises(VBACompileError, match=message):
            app.run('Harness.Probe')
    elif 'runtime_error' in probe:
        with pytest.raises(VBARuntimeError) as caught:
            app.run('Harness.Probe')
        assert caught.value.number == int(probe['runtime_error'])
    else:
        assert app.run('Harness.Probe') == probe['result']


def test_nominal_mismatch_in_dead_branch_precedes_mutation() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Type FirstData\nValue As Long\nEnd Type\nPublic Type SecondData\nValue As Long\nEnd Type\nPublic Counter As Long\nSub Change(ByRef item As FirstData)\nEnd Sub\nFunction Check() As Long\nDim item As SecondData\nCounter = Counter + 1\nIf False Then\nChange item\nEnd If\nEnd Function\nFunction ReadCounter() As Long\nReadCounter = Counter\nEnd Function', name='Caller')
    with pytest.raises(VBACompileError, match='ByRef argument type mismatch'):
        app.run('Caller.Check')
    assert app.run('Caller.ReadCounter') == 0
