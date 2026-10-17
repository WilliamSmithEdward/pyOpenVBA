"""Native aggregate controls and usable populated report persistence."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._pivots import NS

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'aggregation.json').read_text(encoding='utf-8'))
ARGUMENTS: dict[str, Any] = json.loads((ROOT / 'datafield_arguments.json').read_text(encoding='utf-8'))
INSPECT = '''Function InspectAggregate() As String
Dim p As PivotTable,f As PivotField
Set p=Worksheets(1).PivotTables("Made")
Set f=p.DataFields(1)
InspectAggregate=f.Name & "|" & CStr(f.Function) & "|" & p.TableRange1.Address & "|" & CStr(VarType(p.DataBodyRange.Cells(1,1).Value2)) & "|" & CStr(p.DataBodyRange.Cells(1,1).Value2)
End Function'''
EMOJI_INSPECT = '''Function InspectEmojiCaption() As String
Dim caption As String,p As PivotTable
caption=ChrW(&HD83D) & ChrW(&HDE00)
Set p=Worksheets(1).PivotTables("Made")
InspectEmojiCaption=CStr(p.DataFields(caption).Name=caption) & "|" & CStr(p.DataFields(1).Name=Worksheets(1).Range("E3").Value2) & "|" & CStr(p.GetPivotData(caption).Value2)
End Function'''


def _project(case: dict[str, Any], path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module(INSPECT, name='Inspection')
    assert app.run('InspectAggregate') == case['result']
    return app.save(path)


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_pivot_aggregation(case: dict[str, Any], tmp_path: Path) -> None:
    path = _project(case, tmp_path / 'aggregate.xlsm')
    app = ExcelApplication.open(path)
    assert app.run('InspectAggregate') == case['result']
    app.add_module('''Sub ChangeSource()
Worksheets(1).Range("B2:B6").Value=999
End Sub
Function ReadTotal() As String
ReadTotal=CStr(Worksheets(1).PivotTables("Made").GetPivotData("Amount").Value2)
End Function''', name='Query')
    app.run('ChangeSource')
    assert app.run('ReadTotal') == case['result'].rsplit('|', 1)[-1]
    assert app.run('InspectAggregate') == case['result']
    path = app.save(path)
    assert ExcelApplication.open(path).run('InspectAggregate') == case['result']


@pytest.mark.parametrize('case', ARGUMENTS['probes'], ids=lambda case: case['name'])
def test_native_datafield_arguments(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    for name, source in case.get('classes', {}).items():
        app.add_module(source, name=name, kind='class')
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


def test_add_fields_to_saved_cache_preserves_snapshot(tmp_path: Path) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == 'numeric-xlSum')
    path = _project(case, tmp_path / 'shared.xlsm')
    with ZipFile(path) as package:
        cached = {name: package.read(name) for name in package.namelist() if name.startswith('xl/pivotCache/')}
    app = ExcelApplication.open(path)
    app.add_module('''Sub AddMore()
Dim p As PivotTable
Set p=Worksheets(1).PivotTables("Made")
Worksheets(1).Range("B2:B6").Value=999
p.AddDataField p.PivotFields("Amount"),"Counted",xlCount
p.AddDataField p.PivotFields("Amount"),"Again",xlSum
End Sub
Function ReadMore() As String
Dim p As PivotTable
Set p=Worksheets(1).PivotTables("Made")
ReadMore=CStr(p.DataFields.Count) & "|" & CStr(p.GetPivotData("Sum of Amount").Value2) & "|" & CStr(p.GetPivotData("Counted").Value2) & "|" & CStr(p.GetPivotData("Again").Value2) & "|" & p.TableRange1.Address
End Function''', name='More')
    app.run('AddMore')
    for _ in range(3):
        assert app.run('ReadMore') == '3|150|5|150|$E$3:$G$4'
        path = app.save(path)
        with ZipFile(path) as package:
            assert {name: package.read(name) for name in cached} == cached
            part = next(name for name in package.namelist()
                        if name.startswith('xl/pivotTables/pivotTable') and name.endswith('.xml'))
            location = xml_tree(package.read(part)).find(NS + 'location')
            assert location is not None and location.get('firstHeaderRow') == '0'
        app = ExcelApplication.open(path)


def _emoji_project(path: Path) -> Path:
    case = next(case for case in ARGUMENTS['probes'] if case['name'] == 'caption-emoji')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module(EMOJI_INSPECT, name='Inspection')
    for _ in range(3):
        assert app.evaluate('Worksheets(1).PivotTables("Made").DataFields(1).Name') == '\U0001f600'
        assert app.run('InspectEmojiCaption') == 'True|True|150'
        app.save(path)
        app = ExcelApplication.open(path)
    return path


def test_emoji_field_caption_roundtrip(tmp_path: Path) -> None:
    _emoji_project(tmp_path / 'emoji.xlsm')


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_opens_populated_report(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    path = _project(case, tmp_path / 'aggregate.xlsm')
    native_excel.open_document(path)
    assert native_excel.run_raw('Inspection.InspectAggregate') == case['result']
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', [case for case in ARGUMENTS['probes'] if case['name'].startswith('getters-')],
                         ids=lambda case: case['name'])
def test_native_saved_field_ranges(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    tail = case['source'].split('On Error Resume Next\n', 1)[1].replace('Probe', 'InspectSaved')
    inspection = '''Function InspectSaved() As String
Dim p As PivotTable,f As PivotField
Set p=Worksheets(1).PivotTables("Made")
Set f=p.DataFields(p.DataFields.Count)
On Error Resume Next
''' + tail
    app.add_module(inspection, name='Inspection')
    assert app.run('InspectSaved') == case['result']
    path = app.save(tmp_path / 'fields.xlsm')
    native_excel.open_document(path)
    assert native_excel.run_raw('Inspection.InspectSaved') == case['result']
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
def test_native_emoji_field_caption(tmp_path: Path, native_excel: Any) -> None:
    path = _emoji_project(tmp_path / 'emoji.xlsm')
    native_excel.open_document(path)
    assert native_excel.run_raw('Inspection.InspectEmojiCaption') == 'True|True|150'
    assert not any(event.get('dialogs') for event in native_excel.events)
