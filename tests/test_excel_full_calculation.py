"""Native full/rebuild calculation, formula entry and persisted calculation state."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((ROOT / 'full_calculation.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, Any]] = RECORD['probes']
FILES: list[dict[str, Any]] = RECORD['files']


def test_automatic_calculation_does_not_hide_unsupported_udf_operations() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Private Flag As Boolean\nFunction MyUnsupportedUDF() As Double\nApplication.Volatile\n'
                   'If Flag Then\nMyUnsupportedUDF = ThisWorkbook.PivotCaches.Create(2, "External source").RecordCount\nElse\nMyUnsupportedUDF = 1\nEnd If\nEnd Function\n'
                   'Sub Start()\nRange("A1").Formula = "=MyUnsupportedUDF()"\nEnd Sub\n'
                   'Sub Trigger()\nFlag = True\nRange("B1").Value = 1\nEnd Sub', name='Harness')
    app.run('Harness.Start')
    assert app.evaluate('Range("A1").Value2') == 1
    with pytest.raises(VBAUnsupportedError, match='Create'):
        app.run('Harness.Trigger')


@pytest.mark.parametrize('case', PROBES, ids=[case['name'] for case in PROBES])
def test_native_full_calculation(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['macro'], name='Harness')
    assert app.run('Harness.Probe') in case.get('results', [case['result']])


def _make_file(case: dict[str, Any], path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['macro'], name='Harness')
    assert app.run('Harness.Probe') == case['setup']
    app.run('Harness.SetSaveFlag')
    app.save(path)


def _read_file(path: Path, case: dict[str, Any]) -> None:
    app = ExcelApplication.open(path)
    assert app.evaluate('Application.Calculation') == case['calculation']
    assert app.run('Harness.ReadSaved') == case['result']
    app.run('Harness.FullNow')
    assert app.run('Harness.ReadSaved') == case['full']
    app.run('Harness.RebuildNow')
    assert app.run('Harness.ReadSaved') == case['rebuild']


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_native_calculation_file(case: dict[str, Any]) -> None:
    _read_file(ROOT / case['file'], case)


@pytest.mark.parametrize('case', FILES, ids=[case['name'] for case in FILES])
def test_model_calculation_file(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (case['name'] + '.xlsm')
    _make_file(case, path)
    _read_file(path, case)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_model_calculation_files(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        for case in FILES:
            path = tmp_path / (case['name'] + '.xlsm')
            _make_file(case, path)
            excel.new_document()
            mode = excel.run_vba('Sub SetMode()\nApplication.Calculation = ' + str(case['calculation']) + '\nEnd Sub', 'SetMode')
            assert mode.ok and not mode.dialogs, mode
            excel.open_document(path)
            for entry, expected in [('ReadSaved', 'result'), ('FullNow', None), ('ReadSaved', 'full'),
                                    ('RebuildNow', None), ('ReadSaved', 'rebuild')]:
                result = excel.run_raw('Harness.' + entry)
                assert not excel.events[-1].get('dialogs'), result
                if expected is not None:
                    assert result == case[expected], (case['name'], expected, result)
        assert not any(event.get('dialogs') for event in excel.events), excel.events[-1]
