"""Native VBA record declaration visibility and initialized return storage."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

PROBES: list[dict[str, str]] = []
for fixture in ('record_scope', 'record_returns'):
    PROBES.extend(json.loads(
        (Path(__file__).parent / f'fixtures/{fixture}.json').read_text(encoding='utf-8')
    )['probes'])


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_scope_and_returns(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'compile_error' in probe:
        message = 'Ambiguous name detected' if 'Ambiguous' in probe['compile_error'] else 'User-defined type not defined'
        with pytest.raises(VBACompileError, match=message):
            app.run('Harness.Probe')
    else:
        assert app.run('Harness.Probe') == probe['result']


def test_inaccessible_record_declaration_is_checked_before_body_execution() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Private Type RecordData\nValue As Long\nEnd Type', name='Helpers')
    app.add_module('Private Counter As Long\nFunction Check() As Long\nCounter = Counter + 1\nIf False Then\nDim item As Helpers.RecordData\nEnd If\nEnd Function\nFunction ReadCounter() As Long\nReadCounter = Counter\nEnd Function', name='Caller')
    with pytest.raises(VBACompileError, match='User-defined type not defined'):
        app.run('Caller.Check')
    assert app.run('Caller.ReadCounter') == 0
