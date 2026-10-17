"""Native declaration syntax, typed references and late-bound dispatch."""
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
from pyopenvba.interpreter._parse import parse_module
from pyopenvba.interpreter._values import EMPTY

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/type_characters.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_type_characters(case: dict[str, Any]) -> None:
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


def test_invalid_unreachable_function_reference_precedes_body_side_effects() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Alpha() As Integer\nEnd Function\nFunction Probe() As String\n'
                   'Range("A1").Value2="executed"\nExit Function\nProbe=Alpha$()\nEnd Function', name='Harness')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    cell = app.evaluate('Range("A1")')
    assert isinstance(cell, VBAObject)
    assert cell.vba_get('Value2') is EMPTY


def test_edited_function_return_type_invalidates_reference_checks() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Alpha() As Integer\nAlpha=1\nEnd Function', name='Other')
    app.add_module('Function Probe() As String\nProbe=CStr(Other.Alpha%())\nEnd Function', name='Harness')
    assert app.run('Harness.Probe') == '1'
    app.add_module('Function Alpha() As String\nAlpha="x"\nEnd Function', name='Other')
    with pytest.raises(VBACompileError):
        app.run('Harness.Probe')
    app.add_module('Function Alpha() As Integer\nAlpha=2\nEnd Function', name='Other')
    assert app.run('Harness.Probe') == '2'


@pytest.mark.parametrize('name', ['local_duplicate', 'local_constant_duplicate'])
def test_strict_parser_rejects_duplicate_types_even_in_unused_body(name: str) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == name)
    with pytest.raises(VBACompileError):
        parse_module(case['source'])


def _save_reference_project(path: Path) -> str:
    case = next(case for case in RECORD['probes'] if case['name'] == 'class_byref')
    app = ExcelApplication()
    app.add_workbook()
    for module in case['modules']:
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    app.add_module(case['source'] + '\nSub StoreProbe()\nRange("A1").NumberFormat="@"\nRange("A1").Value2=Probe()\nEnd Sub\n'
                   'Function ReadStored() As String\nReadStored=CStr(Range("A1").Value2)\nEnd Function', name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)
    return str(case['result'])


def test_type_character_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'type-characters.xlsm'
    expected = _save_reference_project(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('Harness.ReadStored') == expected
    assert reopened.run('Harness.Probe') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_type_character_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'type-characters.xlsm'
    expected = _save_reference_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.ReadStored') == expected
        assert excel.run_raw('Harness.Probe') == expected
        assert not any(event.get('dialogs') for event in excel.events)
