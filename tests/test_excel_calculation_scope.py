"""Native calculation boundaries, dependent caches and deferred Calculate notifications."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((ROOT / 'calculation_scope.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, Any]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD.get('files', [])


def _app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['prepare'], name='Setup')
    app.run('Setup.Prepare')
    app.add_module(RECORD['helpers'], name='EventLog')
    app.add_module(RECORD['book_handler'], name=str(app.evaluate('ThisWorkbook.CodeName')), kind='document')
    for name in ['Left', 'Right', 'Blank']:
        code_name = str(app.evaluate(f'Worksheets("{name}").CodeName'))
        app.add_module(RECORD['sheet_handler'], name=code_name, kind='document')
    return app


@pytest.mark.parametrize('case', PROBES, ids=[case['name'] for case in PROBES])
def test_native_calculation_scope(case: dict[str, Any]) -> None:
    app = _app()
    app.add_module(case['macro'], name='Harness')
    assert app.run('Harness.Probe') in case.get('results', [case['result']])


def _make_file(case: dict[str, Any], path: Path) -> None:
    app = _app()
    app.add_module(case['macro'], name='Harness')
    assert app.run('Harness.Probe') == case['setup']
    assert app.run('Harness.SavedValues') == case['before_save']
    app.run('Harness.SetSaveFlag')
    app.save(path)


def _read_file(case: dict[str, Any], path: Path) -> None:
    app = ExcelApplication.open(path)
    assert app.run('Harness.SavedValues') == case['before']
    assert app.run('Harness.CalculateTrace') == case['after']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_scope_file(case: dict[str, Any]) -> None:
    _read_file(case, ROOT / case['file'])


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_model_scope_file(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (case['name'] + '.xlsm')
    _make_file(case, path)
    _read_file(case, path)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_model_scope_files(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        for case in FILES:
            path = tmp_path / (case['name'] + '.xlsm')
            _make_file(case, path)
            excel.new_document()
            mode = excel.run_vba('Sub SetMode()\nApplication.EnableEvents = False\nApplication.Calculation = -4135\nEnd Sub', 'SetMode')
            assert mode.ok and not mode.dialogs, mode
            excel.open_document(path)
            assert excel.run_raw('Harness.SavedValues') == case['before']
            assert excel.run_raw('Harness.CalculateTrace') == case['after']
            assert not any(event.get('dialogs') for event in excel.events)
