"""Series lifecycle observations projected from independent native chart controls."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/chart_lifecycle'
RAW: dict[str, Any] = json.loads((ROOT / 'native.json').read_bytes())
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_bytes())
DEFAULTS: dict[str, Any] = json.loads((ROOT / 'series-default-persistence-native.json').read_bytes())
CONTROLS = [control for control in RAW['probes'] if control['name'] in MANIFEST['admitted_projection']]
CHART_SNAP = '''Function ChartSnap(c As Object) As String
Dim s As String,n As Long,p As Object
On Error GoTo Failed
s="type=" & ReadP(c,"ChartType") & ";series=" & SeriesCount(c)
For n=1 To c.SeriesCollection.Count
If n>4 Then Exit For
Set p=c.SeriesCollection(n)
s=s & ";s" & n & "=[" & SeriesSnap(p) & "]"
Next
ChartSnap=s
Exit Function
Failed:
ChartSnap=s & ";SNAP-ERROR:" & Err.Number & ":" & Err.Description
End Function'''


def project(result: str) -> str:
    # Remove only complete, balanced axis observation values. Keep all series
    # and operation bodies, including every error and held-object observation.
    for marker in (';cat=[', ';val=['):
        while marker in result:
            start = result.index(marker)
            cursor = start + len(marker)
            depth = 1
            while depth:
                character = result[cursor]
                depth += (character == '[') - (character == ']')
                cursor += 1
            result = result[:start] + result[cursor:]
    result = re.sub(r';axes=[^;\n]*;hasCat=[^;\n]*;hasVal=[^;\n]*', '', result)
    result = re.sub(r';parent=(?:ChartGroup|ERR:1004:Application-defined or object-defined error)', '', result)
    return result.replace(";Index=ERR:438:Object doesn't support this property or method", '')


@pytest.mark.parametrize('control', CONTROLS, ids=lambda control: control['name'])
def test_native_series_lifecycle_observation(control: dict[str, Any]) -> None:
    source = lifecycle_source(control)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='ChartLifecycle')
    assert app.run('Probe') == project(control['raw_result'])


def test_raw_native_evidence_remains_exact_and_deferred_observations_retained() -> None:
    assert hashlib.sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['raw_sha256']
    assert RAW['completed'] and RAW['session_closed_utc']
    assert len(CONTROLS) == 2 and len(RAW['probes']) == 8
    for control in CONTROLS:
        assert control['status']['ok'] and not control['status']['dialogs']
        assert ';parent=ChartGroup' in control['raw_result']
        assert ';cat=[' in control['raw_result'] and ';val=[' in control['raw_result']


def lifecycle_source(control: dict[str, Any]) -> str:
    source = re.sub(r'Function ChartSnap\(.*?End Function', CHART_SNAP, control['source'], flags=re.DOTALL)
    source = source.replace(' & ";parent=" & ParentType(s)', '')
    return source.replace(' & ";Index=" & ReadP(c.SeriesCollection(n),"Index")', '')


@pytest.mark.parametrize('control', CONTROLS, ids=lambda control: control['name'])
def test_native_saved_lifecycle_import_and_model_generated_reopen(control: dict[str, Any], tmp_path: Path) -> None:
    source = lifecycle_source(control)
    native = ROOT / control['file']
    assert hashlib.sha256(native.read_bytes()).hexdigest() == MANIFEST['native_inputs'][control['file']]
    loaded = ExcelApplication.open(native)
    loaded.add_module(source, name='ChartLifecycle')
    expected = project(control['reopened'])
    assert loaded.run('Snapshot') == expected
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='ChartLifecycle')
    app.run('Probe')
    output = tmp_path / Path(control['file']).name
    app.save(output)
    reopened = ExcelApplication.open(output)
    reopened.add_module(source, name='ChartLifecycle')
    assert reopened.run('Snapshot') == expected


def test_invalid_order_and_imported_mutations_preserve_state_and_identity() -> None:
    from pyopenvba.apps.excel._charts import Chart, ChartObjects, Series, SeriesState
    from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
    app = ExcelApplication()
    app.add_workbook()
    chart = ChartObjects(app.workbook.sheets_[0]).Add(0, 0, 200, 100).Chart()
    assert isinstance(chart, Chart)
    entries = [SeriesState(name='identical', values=[1]) for _ in range(3)]
    chart.state.series[:] = entries
    held = Series(chart, entries[1])
    chart.state.dirty = False
    chart.sheet.drawing_dirty = False
    chart.sheet.book.saved = True
    original_list = chart.state.series
    for value in (0, 4):
        with pytest.raises(VBARuntimeError) as failure:
            held.vba_set('PlotOrder', value)
        assert failure.value.number == 1004
        assert chart.state.series is original_list
        assert all(actual is expected for actual, expected in zip(chart.state.series, entries, strict=True))
        assert not chart.state.dirty and not chart.sheet.drawing_dirty and chart.sheet.book.saved
    chart.state.original = '<opaque imported chart/>'
    for operation in (lambda: held.Delete(), lambda: held.vba_set('PlotOrder', 1)):
        with pytest.raises(VBAUnsupportedError):
            operation()
        assert not chart.state.dirty and chart.sheet.book.saved
    chart.state.original = ''
    assert held.Delete() is True
    assert chart.state.series is original_list
    assert chart.state.series[0] is entries[0] and chart.state.series[1] is entries[2]


def test_category_array_readback_does_not_own_or_mutate_chart_storage() -> None:
    from pyopenvba.apps.excel._charts import ChartObjects, Series, SeriesState
    from pyopenvba.interpreter._values import VBAArray
    app = ExcelApplication()
    app.add_workbook()
    chart = ChartObjects(app.workbook.sheets_[0]).Add(0, 0, 200, 100).Chart()
    entry = SeriesState(name='one', values=[1, 2], categories=['a', 'b'])
    chart.state.series.append(entry)
    series = Series(chart, entry)
    first = series.XValues()
    assert isinstance(first, VBAArray)
    first.items[0] = 'changed by caller'
    del first
    assert series.XValues().elements() == ['a', 'b']
    assert series.XValues().elements() == ['a', 'b']
    assert entry.categories == ['a', 'b']


def test_detached_chart_cannot_be_mutated_through_a_held_series() -> None:
    from pyopenvba.apps.excel._charts import ChartObjects, Series, SeriesState
    from pyopenvba.exceptions import VBAUnsupportedError
    app = ExcelApplication()
    app.add_workbook()
    item = ChartObjects(app.workbook.sheets_[0]).Add(0, 0, 200, 100)
    chart = item.Chart()
    entry = SeriesState(name='one', values=[1])
    chart.state.series.append(entry)
    held = Series(chart, entry)
    item.Delete()
    for operation in (lambda: held.Delete(), lambda: held.Name(), lambda: held.vba_set('PlotOrder', 1),
                      lambda: chart.SeriesCollection().vba_get('Count')):
        with pytest.raises(VBAUnsupportedError, match='detached chart'):
            operation()
    assert chart.state.series[0] is entry


@pytest.mark.parametrize('deleted_index', [0, 1, 2])
def test_plot_order_revalidates_identity_after_reentrant_argument_coercion(deleted_index: int) -> None:
    from pyopenvba.apps.excel._charts import ChartObjects, Series, SeriesState
    from pyopenvba.exceptions import VBARuntimeError
    from pyopenvba.interpreter._objects import VBAObject
    app = ExcelApplication()
    app.add_workbook()
    chart = ChartObjects(app.workbook.sheets_[0]).Add(0, 0, 200, 100).Chart()
    entries = [SeriesState(name=name, values=[1]) for name in ('A', 'B', 'C')]
    chart.state.series[:] = entries
    held = Series(chart, entries[1])

    class ReentrantOrder(VBAObject):
        def vba_value(self) -> object:
            Series(chart, entries[deleted_index]).Delete()
            return 1

    if deleted_index == 1:
        with pytest.raises(VBARuntimeError) as failure:
            held.vba_set('PlotOrder', ReentrantOrder())
        assert failure.value.number == -2147467261
        assert chart.state.series[0] is entries[0] and chart.state.series[1] is entries[2]
    else:
        held.vba_set('PlotOrder', ReentrantOrder())
        assert chart.state.series[0] is entries[1]
        assert chart.state.series[1] is entries[2 if deleted_index == 0 else 0]
    assert len({id(entry) for entry in chart.state.series}) == 2


@pytest.fixture(scope='module')
def native_lifecycle_session() -> Any:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('control', CONTROLS + DEFAULTS['probes'], ids=lambda control: control['name'])
def test_native_excel_reopens_generated_series_lifecycle(control: dict[str, Any], tmp_path: Path,
                                                       native_lifecycle_session: Any) -> None:
    source = lifecycle_source(control)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='ChartLifecycle')
    app.run('Probe')
    output = tmp_path / Path(control['file']).name
    app.save(output)
    native_lifecycle_session.open_document(output)
    result = native_lifecycle_session.run_vba(source, 'Snapshot', timeout=60.0)
    assert result.ok and not result.dialogs, result
    expected = control.get('reopened', control.get('reopened_result'))
    assert isinstance(expected, str)
    assert result.value == project(expected)


COERCIONS: dict[str, Any] = json.loads((ROOT / 'series-coercions-native.json').read_bytes())


@pytest.mark.parametrize('control', COERCIONS['probes'][:2], ids=lambda control: control['name'])
def test_completed_native_plot_order_coercion_control(control: dict[str, Any]) -> None:
    assert control['status']['ok'] and not control['status']['dialogs']
    assert hashlib.sha256(control['source'].encode()).hexdigest() == control['source_sha256']
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(lifecycle_source(control), name='ChartCoercions')
    assert app.run('Probe') == project(control['raw_result'])


def test_coercion_raw_bytes_and_timeout_diagnostic_are_retained() -> None:
    assert hashlib.sha256((ROOT / 'series-coercions-native.json').read_bytes()).hexdigest() == MANIFEST['coercions_raw_sha256']
    assert len(COERCIONS['probes']) == 3 and COERCIONS['session_closed_utc']
    assert not COERCIONS['probes'][2]['status']['ok']
    assert COERCIONS['probes'][2]['status']['outcome'] == 'timeout'





@pytest.mark.parametrize('control', DEFAULTS['probes'], ids=lambda control: control['name'])
def test_native_stable_default_names_runtime_and_persistence(control: dict[str, Any], tmp_path: Path) -> None:
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile
    source = lifecycle_source(control)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='StableSeries')
    assert app.run('Probe') == project(control['raw_result'])
    output = tmp_path / Path(control['file']).name
    app.save(output)
    expected = project(control['reopened_result'])
    for path in (ROOT / output.name, output):
        loaded = ExcelApplication.open(path)
        loaded.add_module(source, name='StableSeries')
        assert loaded.run('Snapshot') == expected
    ns = {'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    with ZipFile(output) as package:
        saved = ET.fromstring(package.read('xl/charts/chart1.xml'))
    native = ET.fromstring(control['chart_xml']['xl/charts/chart1.xml'])
    def identities(root: ET.Element) -> list[tuple[str | None, str | None]]:
        result: list[tuple[str | None, str | None]] = []
        for node in root.findall('.//c:ser', ns):
            identity, order = node.find('c:idx', ns), node.find('c:order', ns)
            assert identity is not None and order is not None
            result.append((identity.get('val'), order.get('val')))
        return result
    assert identities(saved) == identities(native)


def test_default_name_evidence_and_macrofree_inputs_remain_exact() -> None:
    assert hashlib.sha256((ROOT / 'series-default-persistence-native.json').read_bytes()).hexdigest() == MANIFEST['defaults_raw_sha256']
    assert DEFAULTS['completed'] and DEFAULTS['session_closed_utc']
    for control in DEFAULTS['probes']:
        assert control['status']['ok'] and control['reopen_status']['ok']
        assert hashlib.sha256(control['source'].encode()).hexdigest() == control['source_sha256']
    for name, expected in MANIFEST['defaults_inputs'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected


@pytest.mark.parametrize('kind,value', [('idx', '-1'), ('idx', '4294967296'), ('idx', '9' * 10000),
                                      ('order', 'NaN'), ('order', '-1')])
def test_invalid_series_identity_or_order_refuses_before_integer_expansion(kind: str, value: str) -> None:
    from pyopenvba.apps.excel._charts_file import read_chart
    from pyopenvba.exceptions import VBAUnsupportedError
    control = DEFAULTS['probes'][0]
    xml = control['chart_xml']['xl/charts/chart1.xml']
    xml, count = re.subn(r'(<c:' + kind + r' val=")[^"]*', lambda match: match[1] + value, xml, count=1)
    assert count == 1
    with pytest.raises(VBAUnsupportedError, match='series identity/order'):
        read_chart(xml, 'xl/charts/chart1.xml')


@pytest.mark.parametrize('kind', ['idx', 'order'])
def test_ambiguous_series_ids_or_orders_refuse_on_public_workbook_import(kind: str, tmp_path: Path) -> None:
    from zipfile import ZipFile

    from pyopenvba.exceptions import VBAUnsupportedError
    source = ROOT / 'unnamed-reorder-delete-save.xlsx'
    target = tmp_path / 'ambiguous.xlsx'
    with ZipFile(source) as original, ZipFile(target, 'w') as saved:
        for info in original.infolist():
            data = original.read(info.filename)
            if info.filename == 'xl/charts/chart1.xml':
                data = re.sub(rb'(<c:' + kind.encode() + rb' val=")[^"]*', rb'\g<1>0', data)
            saved.writestr(info, data)
    with pytest.raises(VBAUnsupportedError, match='Ambiguous chart series'):
        ExcelApplication.open(target)
