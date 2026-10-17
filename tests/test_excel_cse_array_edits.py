"""Native CSE array fragments, copying, moving and refused operations."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((ROOT / 'cse_array_edits.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, str]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD['files']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_cse_edit(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_file_retains_cse_ranges_and_caches(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'])
    app.add_module(RECORD['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', name='Readback')
    assert app.evaluate('Application.Calculation') == case['calculation']
    assert app.run('ReadSaved') == case['result']
    app.run('CalculateNow')
    name = case['name'].rsplit('_', 1)[0] + '_calculated'
    expected = next(probe for probe in PROBES if probe['name'] == name)
    assert app.run('ReadSaved') == '0|' + expected['result'].split('|', 1)[1]


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
def test_model_cse_file_survives_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (case['name'] + '.xlsx')
    app = _make_file(case, path)
    assert app.run('ReadSaved') == case['result']
    reopened = ExcelApplication.open(path)
    reopened.add_module(RECORD['read_saved'], name='Readback')
    assert reopened.run('ReadSaved') == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_opens_model_cse_arrays(tmp_path: Path) -> None:
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
            result = excel.run_vba(RECORD['read_saved'], 'ReadSaved', timeout=60.0)
            assert result.ok and not result.dialogs, result
            assert result.value == case['result']
