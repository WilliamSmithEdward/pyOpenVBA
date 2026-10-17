"""Native module default types, declaration precedence and compiler restrictions."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/default_types.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_default_type(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sentinel = app.evaluate('Range("A1")')
    assert isinstance(sentinel, VBAObject)
    for module in case.get('modules', []):
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    if case.get('compile_error'):
        with pytest.raises(VBACompileError):
            app.add_module(case['source'], name='Harness')
            app.run('Harness.Probe')
        assert sentinel.vba_get('Value2') is EMPTY
    else:
        app.add_module(case['source'], name='Harness')
        assert app.run('Harness.Probe') == case['result']


def test_source_edits_rebind_redim_against_current_project_globals() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('DefInt A\nFunction Probe() As String\nReDim apple(0)\nProbe = TypeName(apple)\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == 'Integer()'
    app.add_module('Public apple() As String', name='Other')
    assert app.run('Harness.Probe') == 'String()'
    app.add_module('Public apple() As Double', name='Other')
    assert app.run('Harness.Probe') == 'Double()'


def _save_default_type_project(path: Path) -> str:
    case = next(case for case in RECORD['probes'] if case['name'] == 'scope_class_scope')
    app = ExcelApplication()
    app.add_workbook()
    for module in case['modules']:
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    app.add_module(case['source'] + '\nSub StoreProbe()\nRange("A1").Value2 = Probe()\nEnd Sub\n'
                   'Function ReadStored() As String\nReadStored = CStr(Range("A1").Value2)\nEnd Function', name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)
    return str(case['result'])


def test_default_type_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'default-types.xlsm'
    expected = _save_default_type_project(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('Harness.ReadStored') == expected
    assert reopened.run('Harness.Probe') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_default_type_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'default-types.xlsm'
    expected = _save_default_type_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.ReadStored') == expected
        assert excel.run_raw('Harness.Probe') == expected
        assert not any(event.get('dialogs') for event in excel.events)


@pytest.mark.parametrize('name', ['demand_missing_label', 'demand_invalid_exit', 'demand_optional_default', 'demand_deftype_body'])
def test_deferred_procedure_error_still_fails_when_called(name: str) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == name)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == 'done'
    with pytest.raises(VBACompileError):
        app.run('Harness.Unused')
    assert app.run('Harness.Probe') == 'done'
