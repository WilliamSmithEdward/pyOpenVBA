"""Chart mutations prepare source state before publishing it or allocating caches."""
from __future__ import annotations

from pathlib import Path
from typing import cast
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication, _charts, _charts_file
from pyopenvba.apps.excel._charts import Chart, ChartObject, ChartObjects, ChartTitle, Series, SeriesState
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._model import Range
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL, VBAArray


def chart_with_prior_state() -> Chart:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    item = cast('ChartObject', ChartObjects(sheet).Add(10, 20, 300, 180))
    chart = cast('Chart', item.Chart())
    chart.state.series.append(SeriesState(name='prior', values=[2, 4]))
    chart.state.title = 'Prior title'
    chart.state.plot_by = 1
    chart.state.dirty = False
    sheet.drawing_dirty = False
    sheet.book.saved = True
    return chart


def snapshot(chart: Chart) -> object:
    return (id(chart.state), id(chart.state.series),
            [(id(entry), entry.name, tuple(entry.values), tuple(entry.categories),
              entry.name_reference, entry.value_reference, entry.category_reference) for entry in chart.state.series],
            chart.state.plot_by, chart.state.title, chart.state.legend, chart.state.original,
            chart.state.dirty, chart.sheet.drawing_dirty, chart.sheet.book.saved)


def test_failed_source_inference_preserves_series_identity_and_dirty_state(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    before = snapshot(chart)
    calls = 0

    def fail_during_header_inference(_range: Range) -> object:
        nonlocal calls
        calls += 1
        if calls == 3:
            raise VBAUnsupportedError('Source calculation unavailable')
        return 'header'

    monkeypatch.setattr(Range, 'Value', fail_during_header_inference)
    with pytest.raises(VBAUnsupportedError, match='Source calculation'):
        chart.SetSourceData(chart.sheet.Range('A1:C3'), 2)
    assert calls == 3
    assert snapshot(chart) == before


@pytest.mark.parametrize('member', ['HasTitle', 'HasLegend', 'TitleText', 'SeriesName', 'Values', 'XValues'])
def test_failed_member_conversion_preserves_chart_state(member: str) -> None:
    chart = chart_with_prior_state()
    series = Series(chart, chart.state.series[0])
    before = snapshot(chart)
    with pytest.raises(VBARuntimeError):
        if member == 'TitleText':
            ChartTitle(chart).vba_set('Text', NULL)
        elif member in ('SeriesName', 'Values', 'XValues'):
            series.vba_set('Name' if member == 'SeriesName' else member, NULL)
        else:
            chart.vba_set(member, NULL)
    assert snapshot(chart) == before


def test_source_budget_refuses_before_reading_any_cell(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    before = snapshot(chart)
    monkeypatch.setattr(_charts, 'MAX_CHART_POINTS', 4)

    def forbidden(_range: Range) -> object:
        pytest.fail('source cells were read before extent admission')

    monkeypatch.setattr(Range, 'Value', forbidden)
    with pytest.raises(VBAUnsupportedError, match='cell budget'):
        chart.SetSourceData(chart.sheet.Range('A1:C2'))
    assert snapshot(chart) == before


def test_series_budget_preserves_source_state(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    before = snapshot(chart)
    monkeypatch.setattr(_charts, 'MAX_CHART_SERIES', 2)
    with pytest.raises(VBAUnsupportedError, match='series budget'):
        chart.SetSourceData(chart.sheet.Range('A1:C2'), 2)
    assert snapshot(chart) == before


@pytest.mark.parametrize('operation', ['source', 'values', 'categories'])
def test_cross_workbook_sources_cannot_bind_silently_to_local_sheet(operation: str) -> None:
    chart = chart_with_prior_state()
    other = ExcelApplication()
    other.add_workbook()
    selected = other.workbook.sheets_[0].Range('A1:A2')
    before = snapshot(chart)
    with pytest.raises(VBAUnsupportedError, match='Cross-workbook'):
        if operation == 'source':
            chart.SetSourceData(selected)
        else:
            Series(chart, chart.state.series[0]).vba_set('Values' if operation == 'values' else 'XValues', selected)
    assert snapshot(chart) == before


def test_reference_union_budget_precedes_all_reads(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    monkeypatch.setattr(_charts, 'MAX_CHART_POINTS', 5)

    def forbidden(_range: Range) -> object:
        pytest.fail('a partial chart reference was read before aggregate admission')

    monkeypatch.setattr(Range, 'Value2', forbidden)
    with pytest.raises(VBAUnsupportedError, match='cell budget'):
        _charts.referenced_values(chart.sheet, 'A1:A3,B1:B3')


def test_small_reference_union_preserves_order_and_duplicates(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    monkeypatch.setattr(_charts, 'MAX_CHART_POINTS', 4)
    chart.sheet.Range('A1').vba_set('Value2', 3)
    chart.sheet.Range('A2').vba_set('Value2', 7)
    assert _charts.referenced_values(chart.sheet, 'A1:A2,A2:A2') == [3, 7, 7]


@pytest.mark.parametrize('index', ['-1', '4', '1000000000000', 'invalid', '1.5'])
def test_imported_sparse_point_index_is_admitted_before_dense_allocation(index: str,
                                                                       monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_charts_file, 'MAX_CHART_POINTS', 4)
    node = xml_tree('<c:val xmlns:c="' + _charts_file.NS['c'] + '"><c:numLit><c:pt idx="' + index
                    + '"><c:v>2</c:v></c:pt></c:numLit></c:val>')
    with pytest.raises(VBAUnsupportedError, match=r'(point index|point indexes)'):
        getattr(_charts_file, '_data')(node)


def test_imported_sparse_point_exact_boundary_retains_empty_slots(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_charts_file, 'MAX_CHART_POINTS', 4)
    node = xml_tree('<c:val xmlns:c="' + _charts_file.NS['c']
                    + '"><c:numLit><c:pt idx="3"><c:v>2</c:v></c:pt></c:numLit></c:val>')
    assert getattr(_charts_file, '_data')(node) == ('', [EMPTY, EMPTY, EMPTY, 2.0])


def test_series_array_budget_precedes_copy(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    monkeypatch.setattr(_charts, 'MAX_CHART_POINTS', 2)
    values = VBAArray([(0, 2)], items=[1, 2, 3])
    before = snapshot(chart)
    with pytest.raises(VBAUnsupportedError, match='point budget'):
        Series(chart, chart.state.series[0]).vba_set('Values', values)
    assert snapshot(chart) == before


def test_workbook_import_rejects_sparse_point_outside_admission_bound(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    item = cast('ChartObject', ChartObjects(app.workbook.sheets_[0]).Add(10, 20, 300, 180))
    chart = cast('Chart', item.Chart())
    chart.state.series.append(SeriesState(values=[2.0]))
    original = app.save(tmp_path / 'ordinary.xlsx')
    derived = tmp_path / 'sparse.xlsx'
    with ZipFile(original) as source, ZipFile(derived, 'w') as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == 'xl/charts/chart1.xml':
                assert b'idx="0"' in data
                data = data.replace(b'idx="0"', b'idx="1000000000000"', 1)
            target.writestr(entry, data)
    before = derived.read_bytes()
    with pytest.raises(VBAUnsupportedError, match='point index'):
        ExcelApplication.open(derived, with_vba=False)
    assert derived.read_bytes() == before


def test_chart_reference_text_bound_precedes_reference_parser(monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    monkeypatch.setattr(_charts, 'MAX_CHART_REFERENCE_TEXT', 8)

    def forbidden(_text: str) -> list[object]:
        pytest.fail('oversized chart reference reached the parser')

    monkeypatch.setattr(_charts, 'parse_reference', forbidden)
    with pytest.raises(VBAUnsupportedError, match='text budget'):
        _charts.referenced_values(chart.sheet, 'A1,A2,A3,A4')


@pytest.mark.parametrize('member', ['Name', 'Values', 'XValues'])
def test_reference_setter_text_bound_preserves_chart_state(member: str, monkeypatch: pytest.MonkeyPatch) -> None:
    chart = chart_with_prior_state()
    before = snapshot(chart)
    monkeypatch.setattr(_charts, 'MAX_CHART_REFERENCE_TEXT', 8)
    with pytest.raises(VBAUnsupportedError, match='text budget'):
        Series(chart, chart.state.series[0]).vba_set(member, '=Data!A1:A4')
    assert snapshot(chart) == before
