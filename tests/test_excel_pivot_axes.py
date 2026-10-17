"""Native Excel axis layouts, field moves and cached record persistence."""
from __future__ import annotations

import json
import importlib.util
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _pivot_aggregate

ROOT = Path(__file__).parent / 'fixtures/pivots'
CASES: list[dict[str, Any]] = []
for filename in ('axis_layouts', 'axis_mixed', 'orientation_arguments', 'orientation_extra'):
    record = json.loads((ROOT / (filename + '.json')).read_text(encoding='utf-8'))
    for case in record['probes']:
        CASES.append(dict(case, name=filename + '-' + case['name']))


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_pivot_axes(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    if not case['result'].startswith('ERR:'):
        path = app.save(tmp_path / 'axes.xlsm')
        assert ExcelApplication.open(path).run('Snapshot') == case['result']


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', [case for case in CASES if not case['result'].startswith('ERR:')],
                         ids=lambda case: case['name'])
def test_native_reopens_axis_report(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    native_excel.open_document(app.save(tmp_path / 'axes.xlsm'))
    assert native_excel.run_raw('Harness.Snapshot') == case['result']
    # Exercise Excel's own aggregation over the generated shared-item cache.
    result = native_excel.run_vba('''Function Recompute() As String
Dim p As PivotTable
Set p=Worksheets(1).PivotTables("Made")
p.AddDataField p.PivotFields("Amount"),"NativeCheck",xlSum
Recompute=CStr(p.GetPivotData("NativeCheck").Value2)
End Function''', 'Recompute', line_numbers=False)
    assert result.ok and not result.dialogs and result.value == '150', result


def test_layout_uses_snapshot_and_retained_field_proxy(tmp_path: Path) -> None:
    case = next(case for case in CASES if case['name'] == 'axis_layouts-row-sum')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    source = '''Function Move() As String
Dim p As PivotTable,f As PivotField
Set p=Worksheets(1).PivotTables("Made")
Set f=p.PivotFields("Region")
Worksheets(1).Range("C2:C6").Value=999
f.Orientation=xlColumnField
f.Orientation=xlRowField
Move=CStr(f.Orientation) & "|" & CStr(p.GetPivotData("Total").Value2)
End Function'''
    app.add_module(source, name='MoveModule')
    assert app.run('Move') == '1|150'
    path = app.save(tmp_path / 'snapshot.xlsm')
    app = ExcelApplication.open(path)
    assert app.run('Move') == '1|150'
    assert app.run('Snapshot') == case['result']


def test_sparse_axes_aggregate_only_observed_intersections(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = 0
    original = _pivot_aggregate.aggregate

    def counted(values: Any, function: str) -> Any:
        nonlocal calls
        calls += 1
        return original(values, function)

    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Sparse() As String
Dim i As Long,p As PivotTable,c As PivotCache,ws As Worksheet
Set ws=Worksheets(1)
ws.Range("A1:C1").Value=Array("Row","Column","Amount")
For i=1 To 40
ws.Cells(i+1,1).Value=i
ws.Cells(i+1,2).Value=i
ws.Cells(i+1,3).Value=1
Next
Set c=ThisWorkbook.PivotCaches.Create(xlDatabase,ws.Range("A1:C41"))
Set p=c.CreatePivotTable(ws.Range("F5"),"Made")
p.PivotFields("Row").Orientation=xlRowField
p.PivotFields("Column").Orientation=xlColumnField
p.AddDataField p.PivotFields("Amount"),"Total",xlSum
Sparse=CStr(p.GetPivotData("Total").Value2)
End Function''', name='Harness')
    monkeypatch.setattr(_pivot_aggregate, 'aggregate', counted)
    assert app.run('Sparse') == '40'
    assert calls == 3 * 40 + 1
