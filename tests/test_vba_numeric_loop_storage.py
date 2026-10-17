"""Native numeric For bounds, changing array indices and replaced storage."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/numeric_loop_storage.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_numeric_loop_storage(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    for module in case.get('modules', []):
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == case['result']


def _save_loop_project(path: Path) -> str:
    names = ('array_fraction_step', 'Other_fraction_limit', 'index_changes', 'counter_resize', 'record_field')
    cases = [case for case in RECORD['probes'] if case['name'] in names]
    app = ExcelApplication()
    app.add_workbook()
    expected: list[str] = []
    calls: list[str] = []
    for index, case in enumerate(cases):
        for module in case.get('modules', []):
            app.add_module(module['source'], name=module['name'], kind=module['kind'])
        name = f'Loop{index}'
        app.add_module(case['source'], name=name)
        expected.append(case['result'])
        calls.append(name + '.Probe()')
    app.add_module('Function ProbeAll() As String\nProbeAll=' + ' & "/" & '.join(calls)
                   + '\nEnd Function\nSub StoreProbe()\nRange("A1").Value2=ProbeAll()\nEnd Sub\n'
                   'Function ReadStored() As String\nReadStored=CStr(Range("A1").Value2)\nEnd Function', name='Harness')
    app.run('Harness.StoreProbe')
    app.save(path)
    return '/'.join(expected)


def test_numeric_loop_project_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'numeric-loops.xlsm'
    expected = _save_loop_project(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('Harness.ReadStored') == expected
    assert reopened.run('Harness.ProbeAll') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_numeric_loop_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'numeric-loops.xlsm'
    expected = _save_loop_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.ReadStored') == expected
        assert excel.run_raw('Harness.ProbeAll') == expected
        assert not any(event.get('dialogs') for event in excel.events)
