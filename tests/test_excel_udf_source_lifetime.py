"""Replay native UDF source-edit results within and across macro calls."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/udf_source_lifetime.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, Any]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD['files']


@pytest.mark.parametrize('case', PROBES, ids=[case['name'] for case in PROBES])
def test_native_udf_source_lifetime(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['macro'], name='Harness')
    for step in case['steps']:
        assert app.run('Harness.' + step['entry']) == step['result']


def test_python_source_replacement_invalidates_cached_udf_dependents() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function MyCalc(value As Double) As Double\nMyCalc = value * 2\nEnd Function', name='UserCode')
    app.add_module('Sub Setup()\nApplication.Calculation = xlCalculationManual\nRange("A1").Value = 3\n'
                   'Range("B1").Formula = "=MyCalc(A1)"\nRange("C1").Formula = "=B1+1"\nApplication.Calculate\nEnd Sub\n'
                   'Sub CalculateNow()\nApplication.Calculate\nEnd Sub', name='Harness')
    app.run('Harness.Setup')
    app.add_module('Function MyCalc(value As Double) As Double\nMyCalc = value * 3\nEnd Function', name='UserCode')
    assert app.evaluate('Range("B1").Text') == '6'
    assert app.evaluate('Range("C1").Text') == '7'
    app.run('Harness.CalculateNow')
    assert app.evaluate('Range("B1").Value2') == 9
    assert app.evaluate('Range("C1").Value2') == 10


def _make_file(case: dict[str, Any], path: Path) -> None:
    probe = next(probe for probe in PROBES if probe['name'] == case['name'])
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'] + '\n' + RECORD['read_saved'], name='Harness')
    for step in probe['steps']:
        assert app.run('Harness.' + step['entry']) == step['result']
    app.run('Harness.SetSaveFlag')
    app.save(path)


def _read_file(path: Path, case: dict[str, Any]) -> None:
    app = ExcelApplication.open(path)
    assert app.evaluate('Application.Calculation') == case['calculation']
    assert app.run('Harness.ReadSaved') == case['result']
    app.run('Harness.CalculateNow')
    assert app.run('Harness.ReadSaved') == case['calculated']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_udf_source_file(case: dict[str, Any]) -> None:
    _read_file(Path(__file__).parent / 'fixtures' / case['file'], case)


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_model_udf_source_file(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (case['name'] + '.xlsm')
    _make_file(case, path)
    _read_file(path, case)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_model_udf_source_files(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        for case in FILES:
            path = tmp_path / (case['name'] + '.xlsm')
            _make_file(case, path)
            excel.new_document()
            mode = excel.run_vba('Sub SetMode()\nApplication.Calculation = ' + str(case['calculation']) + '\nEnd Sub', 'SetMode')
            assert mode.ok and not mode.dialogs, mode
            excel.open_document(path)
            before = excel.run_macro('Harness.ReadSaved')
            assert before.ok and not before.dialogs and before.value == case['result'], before
            calculated = excel.run_macro('Harness.CalculateNow')
            assert calculated.ok and not calculated.dialogs, calculated
            after = excel.run_macro('Harness.ReadSaved')
            assert after.ok and not after.dialogs and after.value == case['calculated'], after
