"""Native source-axis differences between AddDataField and Orientation."""
from __future__ import annotations

import json
import importlib.util
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE = Path(__file__).parent / 'fixtures/pivots/source_axis_data.json'
CASES: list[dict[str, Any]] = json.loads(FIXTURE.read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_source_axis_data_native_control(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    path = app.save(tmp_path / 'source-axis.xlsm')
    assert ExcelApplication.open(path).run('Snapshot') == case['result']


@pytest.mark.parametrize('axis', [1, 2, 3])
def test_add_data_field_preserves_proxy_and_cache_snapshot(axis: int, tmp_path: Path) -> None:
    case = next(case for case in CASES if case['name'] == f'Amount-axis{axis}-add')
    source = case['source'].replace(
        'p.AddDataField p.PivotFields("Amount")',
        'Set f=p.PivotFields("Amount")\n'
        'ws.Range("C2:C6").Value=999\n'
        'p.AddDataField f\n'
        'If f.Orientation<>xlHidden Then Err.Raise 999')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    assert app.run('Probe') == case['result']
    path = app.save(tmp_path / 'snapshot.xlsm')
    assert ExcelApplication.open(path).run('Snapshot') == case['result']


@pytest.mark.parametrize('action', [
    'p.AddDataField f,"Amount"',
    'p.AddDataField f,"Too long" & String(256,"x")',
    'p.AddDataField f,"Invalid",123',
])
@pytest.mark.parametrize('axis', [1, 2, 3])
def test_invalid_add_data_field_leaves_source_axis_and_cells(action: str, axis: int) -> None:
    case = next(case for case in CASES if case['name'] == f'Amount-axis{axis}-orient')
    source = case['source'].replace(
        'p.PivotFields("Amount").Orientation=4',
        'Set f=p.PivotFields("Amount")\n'
        'Dim before As String\n'
        'before=Snapshot()\n'
        'On Error Resume Next\n' + action + '\n'
        'If Err.Number=0 Then Probe="NO ERROR":Exit Function\n'
        'Err.Clear\nOn Error GoTo Failed\n'
        'If f.Orientation<>' + str(axis) + ' Then Err.Raise 999\n'
        'If Snapshot()<>before Then Err.Raise 998\n'
        'Probe="unchanged"\nExit Function')
    source = source.replace('p.PivotFields("Amount").Orientation=' + str(axis),
                            'p.AddDataField p.PivotFields("Region"),"Seed"\n'
                            'p.PivotFields("Amount").Orientation=' + str(axis))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    assert app.run('Probe') == 'unchanged'


def test_removing_page_field_clears_previous_filter_cells(tmp_path: Path) -> None:
    case = next(case for case in CASES if case['name'] == 'Region-axis3-add')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module('''Function OldFilterCells() As String
Dim ws As Worksheet
Set ws=Worksheets(1)
OldFilterCells=CStr(IsEmpty(ws.Range("F3").Value2)) & "|" & CStr(IsEmpty(ws.Range("G3").Value2))
End Function''', name='Inspect')
    assert app.run('OldFilterCells') == 'True|True'
    assert ExcelApplication.open(app.save(tmp_path / 'removed-filter.xlsm')).run('OldFilterCells') == 'True|True'


def _wrapped_project(order: int, wrap: int, tmp_path: Path) -> Path:
    name = f'native-ranges-pages-order{order}-wrap{wrap}.xlsx'
    imported = json.loads((FIXTURE.parent / 'imported_field_ranges.json').read_text(encoding='utf-8'))
    case = next(case for case in imported['probes'] if case['file'] == name)
    app = ExcelApplication.open(FIXTURE.parent / name, with_vba=False)
    app.add_module(case['source'], name='Harness')
    app.add_module('''Sub AddMeasure()
Dim p As PivotTable
Set p=Worksheets(1).PivotTables("Made")
p.AddDataField p.PivotFields("Countable"),"Added",xlCount
End Sub
Function PageCells() As String
Dim p As PivotTable,f As PivotField
Set p=Worksheets(1).PivotTables("Made")
For Each f In p.PivotFields
If f.Orientation=xlPageField Then
If f.LabelRange.Value2<>f.Name Or f.DataRange.Value2<>"(All)" Then PageCells="bad":Exit Function
End If
Next
PageCells="ok"
End Function''', name='Edit')
    app.run('AddMeasure')
    result = app.run('FieldRanges')
    assert isinstance(result, str)
    assert result.split('|Values:')[0] == case['result'].split('|data=')[0]
    assert app.run('PageCells') == 'ok'
    app = ExcelApplication.open(app.save(tmp_path / 'wrapped.xlsm'))
    result = app.run('FieldRanges')
    assert isinstance(result, str)
    assert result.split('|Values:')[0] == case['result'].split('|data=')[0]
    assert app.run('PageCells') == 'ok'

    return tmp_path / 'wrapped.xlsm'


@pytest.mark.parametrize('order', [1, 2])
@pytest.mark.parametrize('wrap', [0, 1, 2, 3])
def test_adding_measure_preserves_imported_wrapped_pages(order: int, wrap: int, tmp_path: Path) -> None:
    _wrapped_project(order, wrap, tmp_path)


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_reopens_source_axis_edit(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    native_excel.open_document(app.save(tmp_path / 'source-axis.xlsm'))
    assert native_excel.run_raw('Harness.Snapshot') == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('order', [1, 2])
@pytest.mark.parametrize('wrap', [0, 1, 2, 3])
def test_native_reopens_edited_wrapped_pages(order: int, wrap: int, tmp_path: Path, native_excel: Any) -> None:
    path = _wrapped_project(order, wrap, tmp_path)
    expected = ExcelApplication.open(path).run('FieldRanges')
    native_excel.open_document(path)
    assert native_excel.run_raw('Harness.FieldRanges') == expected
    assert native_excel.run_raw('Edit.PageCells') == 'ok'
