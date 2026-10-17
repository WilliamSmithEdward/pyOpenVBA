"""Native CSE fill/replacement results, errors, styles and caches."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((ROOT / 'cse_fill_replace.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, str]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD['files']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_cse_fill_replace(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_file_retains_cse_fill_and_replace(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'])
    app.add_module(RECORD['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', name='Readback')
    assert app.evaluate('Application.Calculation') == case['calculation']
    assert app.run('ReadSaved') == case['result']
    app.run('CalculateNow')
    assert app.run('ReadSaved') == case['after_calculate']


def _make_file(case: dict[str, Any], path: Path) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    probe = next(probe for probe in PROBES if probe['name'] == case['name'])
    app.add_module(probe['macro'] + '\nSub SetSaveFlag()\nApplication.CalculateBeforeSave = False\nEnd Sub', name='Harness')
    app.run('Harness.Probe')
    app.run('SetSaveFlag')
    app.save(path)
    return app


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_model_cse_fill_replace_file_reopens(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (case['name'] + '.xlsx')
    app = _make_file(case, path)
    assert app.run('ReadSaved') == case['result']
    reopened = ExcelApplication.open(path)
    reopened.add_module(RECORD['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', name='Readback')
    assert reopened.run('ReadSaved') == case['result']
    reopened.run('CalculateNow')
    assert reopened.run('ReadSaved') == case['after_calculate']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_opens_cse_fill_replace_files(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    saved: list[tuple[Path, dict[str, Any]]] = []
    for case in FILES:
        path = tmp_path / (case['name'] + '.xlsx')
        _make_file(case, path)
        saved.append((path, case))
    with harness.ExcelSession() as excel:
        excel.new_document()
        for path, case in saved:
            mode = excel.run_vba('Sub SetMode()\nApplication.Calculation = ' + str(case['calculation']) + '\nEnd Sub', 'SetMode')
            assert mode.ok and not mode.dialogs, mode
            excel.open_document(path)
            excel.add_module('Readback', RECORD['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', line_numbers=False)
            assert excel.run_raw('Readback.ReadSaved') == case['result']
            excel.run_raw('Readback.CalculateNow')
            assert excel.run_raw('Readback.ReadSaved') == case['after_calculate']
            assert not any(event.get('dialogs') for event in excel.events)
