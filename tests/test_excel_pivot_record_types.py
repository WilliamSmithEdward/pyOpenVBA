"""Typed cache snapshots measured against native Excel, including persisted aggregation."""
from __future__ import annotations

import datetime as dt
import importlib.util
import json
import os
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.apps.excel._model import Workbook

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/pivots/record_types.json').read_text(encoding='utf-8'))
INSPECT = '''Function InspectCache() As String
Dim p As PivotTable, f As PivotField
Set p=Worksheets(1).PivotTables(1)
Set f=p.AddDataField(p.PivotFields("Amount"))
InspectCache=CStr(f.Function) & "|" & CStr(p.DataBodyRange.Cells(1,1).Value2)
End Function'''


def _project(case: dict[str, Any], path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module(INSPECT, name='Inspection')
    return app.save(path)


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_typed_cache_snapshot(case: dict[str, Any], tmp_path: Path) -> None:
    path = _project(case, tmp_path / 'typed.xlsm')
    expected = case['record']
    if expected is None:
        assert ExcelApplication.open(path).evaluate('ThisWorkbook.PivotCaches.Count') == 0
        return
    for _ in range(3):
        with ZipFile(path) as package:
            part = next(name for name in package.namelist() if name.startswith(
                ('xl/pivotCache/pivotCacheRecords', 'xl/pivotCache/richPivotRecords')) and name.endswith('.xml'))
            assert ('richPivotRecords' in part) == expected['rich']
            row = xml_tree(package.read(part)).find(NS + 'r')
            assert row is not None
            item = row[1]
            assert item.tag == NS + expected['kind']
            actual, wanted = item.get('v', ''), expected['attributes'].get('v', '')
            if expected['kind'] == 'n':
                assert float(actual) == float(wanted)
            elif expected['kind'] == 'd':
                assert dt.datetime.fromisoformat(actual) == dt.datetime.fromisoformat(wanted)
                definition = next(name for name in package.namelist()
                                  if name.startswith('xl/pivotCache/pivotCacheDefinition') and name.endswith('.xml'))
                fields = xml_tree(package.read(definition)).find(NS + 'cacheFields')
                assert fields is not None
                summary = fields[1].find(NS + 'sharedItems')
                assert summary is not None
                assert 'minValue' not in summary.attrib and 'maxValue' not in summary.attrib
            else:
                assert actual == wanted
        app = ExcelApplication.open(path)
        assert app.evaluate('ThisWorkbook.PivotCaches(1).RecordCount') == 5
        # Editing the source must not silently refresh the immutable cache snapshot.
        app.add_module('Sub ChangeSource()\nWorksheets(1).Range("B2").Value=999\nEnd Sub', name='Change')
        app.run('ChangeSource')
        path = app.save(path)


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', [case for case in RECORD['probes'] if case['record'] is not None],
                         ids=lambda case: case['name'])
def test_native_aggregation_of_typed_snapshot(case: dict[str, Any], tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = _project(case, tmp_path / 'typed.xlsm')
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Inspection.InspectCache') == case['inspection']
        assert not any(event.get('dialogs') for event in excel.events)


def test_shared_and_independent_rich_cache_snapshots(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._pivot_rich import datasets

    case = next(case for case in RECORD['probes'] if case['name'] == 'spill-error')
    path = _project(case, tmp_path / 'rich.xlsm')
    app = ExcelApplication.open(path)
    app.add_module('''Sub MoreCaches()
ThisWorkbook.PivotCaches(1).CreatePivotTable Worksheets(1).Range("J3"), "Shared"
Worksheets(1).Range("A3").Value=CVErr(2050)
Worksheets(1).Range("B2").Value=100
ThisWorkbook.PivotCaches.Create(xlDatabase,Worksheets(1).Range("A1:B6")).CreatePivotTable Worksheets(1).Range("O3"), "Independent"
End Sub''', name='More')
    app.run('MoreCaches')
    for _ in range(3):
        path = app.save(path)
        app = ExcelApplication.open(path)
        assert app.evaluate('ThisWorkbook.PivotCaches.Count') == 2
        assert app.evaluate('Worksheets(1).PivotTables.Count') == 3
        book = app.evaluate('ThisWorkbook')
        assert isinstance(book, Workbook)
        snapshots = list(datasets(book.package).values())
        assert {'Amount': ['#SPILL!', None, None, None, None]} in snapshots
        assert {'Region': [None, '#CALC!', None, None, None]} in snapshots
