"""Native compiler rules for numeric For storage, including lexical With."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/numeric_loop_validation.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_numeric_loop_validation(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    for module in case.get('modules', []):
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    app.add_module(case['source'], name='Harness')
    if case.get('compile_error'):
        if 'demand_result' in case:
            assert app.run('Harness.Probe') == case['demand_result']
        with pytest.raises(VBACompileError):
            app.run('Harness.Unused' if 'demand_result' in case else 'Harness.Probe')
    else:
        assert app.run('Harness.Probe') == case['result']


def test_invalid_unreachable_counter_fails_before_body_side_effects() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Probe() As String\nRange("A1").Value2="executed"\nExit Function\n'
                   'Dim apple As String\nFor apple=1 To 2\nNext\nEnd Function', name='Harness')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    cell = app.evaluate('Range("A1")')
    assert isinstance(cell, VBAObject)
    assert cell.vba_get('Value2') is EMPTY


@pytest.mark.parametrize('target', ['apple', 'Other.apple'])
def test_global_source_edits_revalidate_cached_counter_types(target: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public apple As Integer', name='Other')
    app.add_module('Function Probe() As String\nFor ' + target + '=1 To 2\nNext\nProbe=CStr(' + target + ')\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == '3'
    app.add_module('Public apple As String', name='Other')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    app.add_module('Public apple As Long', name='Other')
    assert app.run('Harness.Probe') == '3'


def _save_counter_project(path: Path) -> str:
    case = next(case for case in RECORD['probes'] if case['name'] == 'with_record_function')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\nSub StoreProbe()\nRange("A1").Value2=Probe()\nEnd Sub\n'
                   'Function ReadStored() As String\nReadStored=CStr(Range("A1").Value2)\nEnd Function', name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)
    return str(case['result'])


def test_counter_validation_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'counter-validation.xlsm'
    expected = _save_counter_project(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('Harness.ReadStored') == expected
    assert reopened.run('Harness.Probe') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_counter_validation_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'counter-validation.xlsm'
    expected = _save_counter_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.ReadStored') == expected
        assert excel.run_raw('Harness.Probe') == expected
        assert not any(event.get('dialogs') for event in excel.events)
