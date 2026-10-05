"""Native procedure-demand compilation of whole-record arguments."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/compile_demand.json').read_text(encoding='utf-8'))['probes']
@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_compile_demand(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'compile_error' in probe:
        with pytest.raises(VBACompileError, match='Variable required'):
            app.run('Harness.Probe')
    else:
        assert app.run('Harness.Probe') == probe['result']


def test_invalid_dead_branch_is_checked_before_body_mutates_globals() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Type RecordData\nValue As Long\nEnd Type\nPublic Sub Change(ByRef item As RecordData)\nEnd Sub', name='Helpers')
    app.add_module('Private Counter As Long\nFunction Check() As Long\nDim item As RecordData\nCounter = Counter + 1\nIf False Then\nHelpers.Change (item)\nEnd If\nEnd Function\nFunction ReadCounter() As Long\nReadCounter = Counter\nEnd Function', name='Caller')
    with pytest.raises(VBACompileError, match='Variable required'):
        app.run('Caller.Check')
    assert app.run('Caller.ReadCounter') == 0


def test_dependency_edit_invalidates_successful_procedure_check() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Type RecordData\nValue As Long\nEnd Type\nPublic Sub Change(ByRef value As Long)\nvalue = value + 10\nEnd Sub', name='Helpers')
    app.add_module('Function Check() As Long\nDim item As RecordData\nHelpers.Change (item.Value)\nCheck = 13\nEnd Function', name='Caller')
    assert app.run('Caller.Check') == 13
    app.add_module('Public Type InnerData\nValue As Long\nEnd Type\nPublic Type RecordData\nValue As InnerData\nEnd Type\nPublic Sub Change(ByRef value As InnerData)\nvalue.Value = value.Value + 10\nEnd Sub', name='Helpers')
    with pytest.raises(VBACompileError, match='Variable required'):
        app.run('Caller.Check')
