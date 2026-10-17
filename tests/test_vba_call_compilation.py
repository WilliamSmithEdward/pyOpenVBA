"""Native call binding, argument diagnostics and returned default members."""
from __future__ import annotations
import importlib.util
import json
import os
from pathlib import Path
from typing import Any
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError
from pyopenvba.interpreter._ast import Procedure
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._runtime import ModuleRuntime
from pyopenvba.interpreter._values import EMPTY
RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/call_compilation.json').read_text(encoding='utf-8'))

@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_call_compilation(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    if case.get('compile_error'):
        with pytest.raises(VBACompileError):
            for module in case.get('modules', []):
                app.add_module(module['source'], name=module['name'], kind=module['kind'])
            app.add_module(case['source'], name='Harness')
            if 'demand_result' in case:
                assert app.run('Harness.Probe') == case['demand_result']
            app.run('Harness.Unused' if 'demand_result' in case else 'Harness.Probe')
    else:
        for module in case.get('modules', []):
            app.add_module(module['source'], name=module['name'], kind=module['kind'])
        app.add_module(case['source'], name='Harness')
        assert app.run('Harness.Probe') == case['result']


def test_invalid_call_fails_before_caller_body_side_effects() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Alpha(value As Long)\nEnd Sub\nSub Probe()\nRange("A1").Value2="executed"\nAlpha\nEnd Sub', name='Harness')
    with pytest.raises(VBACompileError, match='Argument not optional'):
        app.run('Harness.Probe')
    cell = app.evaluate('Range("A1")')
    assert isinstance(cell, VBAObject)
    assert cell.vba_get('Value2') is EMPTY


def test_direct_macro_invocation_keeps_runtime_argument_errors() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Alpha(value As Long) As Long\nAlpha=value\nEnd Function', name='Harness')
    for args, number in (((), 449), ((1, 2), 450)):
        with pytest.raises(VBARuntimeError) as caught:
            app.run('Harness.Alpha', *args)
        assert caught.value.number == number
    assert app.run('Harness.Alpha', 3) == 3


def test_edited_signature_invalidates_cached_call_checks() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Alpha(value As Long) As Long\nAlpha=value\nEnd Function', name='Other')
    app.add_module('Function Probe() As Long\nProbe=Other.Alpha(1)\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 1
    app.add_module('Function Alpha() As Long\nAlpha=1\nEnd Function', name='Other')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    app.add_module('Function Alpha(value As Long) As Long\nAlpha=value+1\nEnd Function', name='Other')
    assert app.run('Harness.Probe') == 2


def test_deep_member_chain_has_bounded_resolution_work(monkeypatch: pytest.MonkeyPatch) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == 'call_more_deep_chain')
    original = ModuleRuntime.procedure
    lookups = 0

    def counted(runtime: ModuleRuntime, name: str, kind: str = '') -> Procedure | None:
        nonlocal lookups
        lookups += 1
        assert lookups < 2000, 'member type resolution expanded repeatedly'
        return original(runtime, name, kind)

    monkeypatch.setattr(ModuleRuntime, 'procedure', counted)
    app = ExcelApplication()
    app.add_workbook()
    for module in case['modules']:
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == case['result']


def _save_call_project(path: Path) -> str:
    case = next(case for case in RECORD['probes'] if case['name'] == 'returned_default_calls_collection_standard_True')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\nSub StoreProbe()\nRange("A1").NumberFormat="@"\nRange("A1").Value2=Probe()\nEnd Sub\n'
                   'Function ReadStored() As String\nReadStored=CStr(Range("A1").Value2)\nEnd Function', name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)
    return str(case['result'])


def test_call_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'call-compilation.xlsm'
    expected = _save_call_project(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('Harness.ReadStored') == expected
    assert reopened.run('Harness.Probe') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_call_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'call-compilation.xlsm'
    expected = _save_call_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.ReadStored') == expected
        assert excel.run_raw('Harness.Probe') == expected
        assert not any(event.get('dialogs') for event in excel.events)
