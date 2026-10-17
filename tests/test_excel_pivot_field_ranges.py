"""Native PivotTable field and page ranges, including multi-area results."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'field_ranges.json').read_text(encoding='utf-8'))
RETAINED: dict[str, Any] = json.loads((ROOT / 'retained_fields.json').read_text(encoding='utf-8'))
IMPORTED: dict[str, Any] = json.loads((ROOT / 'imported_field_ranges.json').read_text(encoding='utf-8'))


def _project(case: dict[str, Any], path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    path = app.save(path)
    assert ExcelApplication.open(path).run('FieldRanges') == case['result']
    return path


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_field_ranges(case: dict[str, Any], tmp_path: Path) -> None:
    _project(case, tmp_path / 'ranges.xlsm')


def _imported_project(case: dict[str, Any], path: Path) -> Path:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run('FieldRanges') == case['result']
    app.save(path)
    assert ExcelApplication.open(path).run('FieldRanges') == case['result']
    return path


@pytest.mark.parametrize('case', IMPORTED['probes'], ids=lambda case: case['file'])
def test_imported_field_ranges(case: dict[str, Any], tmp_path: Path) -> None:
    _imported_project(case, tmp_path / 'imported.xlsm')


@pytest.mark.parametrize('name,expected', [('row-values-down', 150), ('rows-two', 300),
                                          ('rows1-columns2', 150)])
def test_measure_ranges_enumerate_displayed_cells(name: str, expected: int) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == name)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module('''Function DisplayedSum() As Double
Dim cell As Range,p As PivotTable
Set p=Worksheets(1).PivotTables("Made")
For Each cell In p.DataFields(1).DataRange.Cells
DisplayedSum=DisplayedSum+cell.Value2
Next
End Function''', name='Values')
    assert app.run('DisplayedSum') == expected


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_field_ranges(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    native_excel.open_document(_project(case, tmp_path / 'ranges.xlsm'))
    assert native_excel.run_raw('Harness.FieldRanges') == case['result']
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', IMPORTED['probes'], ids=lambda case: case['file'])
def test_native_imported_field_ranges(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    native_excel.open_document(_imported_project(case, tmp_path / 'imported.xlsm'))
    assert native_excel.run_raw('Harness.FieldRanges') == case['result']
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.mark.parametrize('case', RETAINED['probes'], ids=lambda case: case['name'])
def test_retained_field_references(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
