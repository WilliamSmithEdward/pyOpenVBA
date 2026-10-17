"""Native manual-calculation spill caches, before and after explicit calculation."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((ROOT / 'manual_spill_edits.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, str]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD['files']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_manual_spill_edit(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_manual_file_reads_caches_then_calculates(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'])
    app.add_module(RECORD['read_saved'] + '\nSub CalculateNow()\nApplication.Calculate\nEnd Sub', name='Readback')
    assert app.evaluate('Application.Calculation') == -4135
    assert app.evaluate('Application.CalculateBeforeSave') == case['calculate_before_save']
    assert app.run('ReadSaved') == case['result']
    app.run('CalculateNow')
    expected = next(probe for probe in PROBES if probe['name'] == case['case'] + '_calculated')
    assert app.run('ReadSaved') == expected['result']


def _make_file(case: dict[str, Any], path: Path) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    probe = next(probe for probe in PROBES if probe['name'] == case['case'])
    app.add_module(probe['macro'] + '\n' + RECORD['read_saved'] + '\nSub SetSaveFlag()\nApplication.CalculateBeforeSave = '
                   + str(case['calculate_before_save']) + '\nEnd Sub', name='Harness')
    app.run('Harness.Probe')
    app.run('SetSaveFlag')
    app.save(path)
    return app


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_model_manual_file_preserves_caches_and_settings(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / f"{case['name']}.xlsx"
    app = _make_file(case, path)
    assert app.run('ReadSaved') == case['result']
    with ZipFile(path) as package:
        assert b'calcMode="manual"' in package.read('xl/workbook.xml')
    reopened = ExcelApplication.open(path)
    reopened.add_module(RECORD['read_saved'], name='Readback')
    assert reopened.evaluate('Application.Calculation') == -4135
    assert reopened.evaluate('Application.CalculateBeforeSave') == case['calculate_before_save']
    assert reopened.run('ReadSaved') == case['result']


def test_opening_a_second_workbook_keeps_application_mode() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub OpenOther(ByVal path As String)\nWorkbooks.Open path\nEnd Sub\n' + RECORD['read_saved'], name='Harness')
    case = next(case for case in FILES if case['name'] == 'input_insert_inside')
    app.run('OpenOther', str((ROOT / case['file']).resolve()))
    assert app.evaluate('Application.Calculation') == -4105
    expected = next(probe for probe in PROBES if probe['name'] == 'input_insert_inside_calculated')
    assert app.run('ReadSaved') == expected['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_opens_model_manual_caches(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    saved: list[tuple[Path, dict[str, Any]]] = []
    for case in FILES:
        path = tmp_path / f"{case['name']}.xlsx"
        _make_file(case, path)
        saved.append((path, case))
    with harness.ExcelSession() as excel:
        for path, case in saved:
            excel.open_document(path)
            result = excel.run_vba(RECORD['read_saved'], 'ReadSaved', timeout=60.0)
            assert result.ok and not result.dialogs, result
            assert result.value == case['result']
