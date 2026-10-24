"""Staged local Pivot refresh identity and failed-admission regressions."""
from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pathlib import Path


SOURCE = '''Sub Prepare()
Dim ws As Worksheet,c As PivotCache,u As PivotCache,p As PivotTable,q As PivotTable,z As PivotTable,f As PivotField
Set ws=ActiveSheet
ws.Name="Data"
ws.Range("A1:B1").Value=Array("Region","Amount")
ws.Range("A2:B2").Value=Array("East",10)
ws.Range("A3:B3").Value=Array("West",20)
ws.Range("A4:B4").Value=Array("East",30)
ws.Range("J1:K1").Value=Array("Other","Amount")
ws.Range("J2:K2").Value=Array("X",7)
ws.Range("J3:K3").Value=Array("Y",8)
Set c=ThisWorkbook.PivotCaches.Create(xlDatabase,ws.Range("A1:B4"))
Set p=c.CreatePivotTable(ws.Range("D3"),"First")
Set f=p.AddDataField(p.PivotFields("Amount"),"Total",xlSum)
Set q=c.CreatePivotTable(ws.Range("G3"),"Shared")
Set f=q.AddDataField(q.PivotFields("Amount"),"Mean",xlAverage)
Set u=ThisWorkbook.PivotCaches.Create(xlDatabase,ws.Range("J1:K3"))
Set z=u.CreatePivotTable(ws.Range("M3"),"Unrelated")
Set f=z.AddDataField(z.PivotFields("Amount"),"OtherTotal",xlSum)
End Sub
Sub EditSource()
Range("B2").Value2=100
Range("B4").Value2=300
Range("K2").Value2=70
End Sub
'''


def prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='RefreshSetup')
    app.run('Prepare')
    return app


def state(app: ExcelApplication) -> object:
    book = app.workbook
    return (book.saved, tuple((id(cache), cache.index, cache.cache_id, cache.part, cache.changed,
                              xml_bytes(cache.root), xml_bytes(cache.records) if cache.records is not None else None)
                             for cache in book.pivot_caches),
            tuple((sheet.dirty, tuple((key, cell.value, cell.formula, cell.style, cell.stale) for key, cell in sheet.cells_.items()),
                   tuple((id(table), id(table.root), xml_bytes(table.root), table.changed, table.report_loaded,
                          table.report_metadata) for table in sheet.pivot_tables)) for sheet in book.sheets_),
            book.stylesheet.dxf_count)


def test_refresh_retains_cache_field_data_roots_and_unrelated_snapshot(tmp_path: Path) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    cache, other = app.workbook.pivot_caches
    first = next(table for table in sheet.pivot_tables if table.root.get('name') == 'First')
    shared = next(table for table in sheet.pivot_tables if table.root.get('name') == 'Shared')
    field = first.PivotFields('Amount')
    roots = (cache.root, first.root, field.root, first.root.find(f'{NS}dataFields/{NS}dataField'))
    other_before = xml_bytes(other.root), xml_bytes(other.records) if other.records is not None else None
    app.run('EditSource')
    assert first.RefreshTable() is True
    assert app.workbook.pivot_caches == [other, cache]
    assert (cache.index, other.index) == (2, 1)
    assert (cache.root, first.root, field.root, first.root.find(f'{NS}dataFields/{NS}dataField')) == roots
    assert field.Name() == 'Amount'
    assert sheet.cells_[4, 4].value == 420
    assert sheet.cells_[4, 7].value == 140
    assert sheet.cells_[4, 13].value == 15
    assert shared.cache is cache and first.cache is cache
    assert (xml_bytes(other.root), xml_bytes(other.records) if other.records is not None else None) == other_before
    result = app.save(tmp_path / 'refreshed.xlsm')
    reopened = ExcelApplication.open(result)
    again = reopened.workbook.sheets_[0]
    assert [again.cells_[4, column].value for column in (4, 7, 13)] == [420, 140, 15]
    # Cache ordinal changes never replace package cache identifiers/relationships.
    ids = [(entry.cache_id, entry.part) for entry in reopened.workbook.pivot_caches]
    assert len({entry[0] for entry in ids}) == 2
    assert all(entry[1] for entry in ids)
    origins = [entry.root.find(f'{NS}cacheSource/{NS}worksheetSource') for entry in reopened.workbook.pivot_caches]
    assert all(origin is not None for origin in origins)
    assert [origin.get('ref') for origin in origins if origin is not None] == ['J1:K3', 'A1:B4']


@pytest.mark.parametrize('failure', ['later-merge', 'unknown-cache', 'unknown-report', 'schema', 'bad-axis', 'invalid-location', 'hierarchy', 'field-compact'])
def test_later_report_or_metadata_failure_publishes_nothing(failure: str) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    cache = app.workbook.pivot_caches[0]
    first = next(table for table in sheet.pivot_tables if table.root.get('name') == 'First')
    app.run('EditSource')
    if failure == 'later-merge':
        from pyopenvba._a1 import Area

        sheet.merged_areas.append(Area(3, 4, 4, 4, sheet.name))
    elif failure == 'unknown-cache':
        cache.root.append(XmlElement('{urn:opaque}metadata'))
    elif failure == 'unknown-report':
        first.root.append(XmlElement('{urn:opaque}metadata'))
    elif failure == 'schema':
        sheet.cells_[1, 2].value = 'Renamed'
    elif failure == 'invalid-location':
        location = first.root.find(NS + 'location')
        assert location is not None
        location.set('ref', 'A0')
    elif failure == 'hierarchy':
        group = XmlElement(NS + 'pageFields', {'count': '1'})
        group.append(XmlElement(NS + 'pageField', {'fld': '0', 'hier': '0'}))
        first.root.append(group)
    elif failure == 'field-compact':
        field = first.root.find(f'{NS}pivotFields/{NS}pivotField')
        assert field is not None
        field.set('compact', '0')
    else:
        fields = XmlElement(NS + 'rowFields', {'count': '1'})
        fields.append(XmlElement(NS + 'field', {'x': '999999'}))
        first.root.append(fields)
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        cache.Refresh()
    assert state(app) == before


def test_refresh_disabled_and_pending_preserve_state() -> None:
    from pyopenvba.exceptions import VBARuntimeError

    app = prepared()
    cache = app.workbook.pivot_caches[0]
    cache.set_enable_refresh(False)
    app.run('EditSource')
    before = state(app)
    with pytest.raises(VBARuntimeError, match='Refresh has been disabled'):
        cache.Refresh()
    assert state(app) == before
    cache.set_enable_refresh(True)
    cache.Refresh()
    assert app.workbook.sheets_[0].cells_[4, 4].value == 420


def test_imported_public_enable_refresh_preserves_lazy_record_bytes(tmp_path: Path) -> None:
    from zipfile import ZipFile

    app = prepared()
    first = app.save(tmp_path / 'before.xlsm')
    with ZipFile(first) as package:
        before = {name: package.read(name) for name in package.namelist() if 'pivotCacheRecords' in name}
    imported = ExcelApplication.open(first)
    cache = imported.workbook.pivot_caches[0]
    assert cache.records is None
    cache.set_enable_refresh(False)
    second = imported.save(tmp_path / 'disabled.xlsm')
    with ZipFile(second) as package:
        assert {name: package.read(name) for name in before} == before
    reopened = ExcelApplication.open(second)
    assert reopened.workbook.pivot_caches[0].EnableRefresh() is False
    reopened.workbook.pivot_caches[0].set_enable_refresh(True)
    reopened.run('EditSource')
    reopened.workbook.pivot_caches[0].Refresh()
    third = reopened.save(tmp_path / 'refreshed-import.xlsm')
    final = ExcelApplication.open(third)
    assert [final.workbook.sheets_[0].cells_[4, column].value for column in (4, 7, 13)] == [420, 140, 15]


def test_worksheet_function_context_cannot_publish_cache_or_report_mutations() -> None:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBARuntimeError

    app = prepared()
    cache = app.workbook.pivot_caches[0]
    caller = app.workbook.sheets_[0].Range('Q1')
    assert isinstance(caller, Range)
    before = state(app)
    app.workbook.application.udf_callers.append(caller)
    try:
        with pytest.raises(VBARuntimeError, match='worksheet function cannot change cells'):
            cache.Refresh()
        assert state(app) == before
        with pytest.raises(VBAUnsupportedError, match='worksheet functions'):
            cache.set_enable_refresh(False)
        assert state(app) == before
    finally:
        app.workbook.application.udf_callers.pop()


def test_xml_false_and_zero_totals_have_identical_staged_meaning() -> None:
    observed: list[object] = []
    for flag in ('false', '0'):
        app = prepared()
        app.add_module('''Sub AddRows()
Worksheets("Data").PivotTables("Shared").PivotFields("Region").Orientation=xlRowField
End Sub''', name='RowsSetup')
        app.run('AddRows')
        sheet = app.workbook.sheets_[0]
        shared = next(table for table in sheet.pivot_tables if table.root.get('name') == 'Shared')
        shared.root.set('rowGrandTotals', flag)
        shared.root.set('colGrandTotals', flag)
        app.run('EditSource')
        app.workbook.pivot_caches[0].Refresh()
        location = shared.root.find(NS + 'location')
        assert location is not None
        assert location.get('ref') == 'G3:H5'
        assert shared.root.get('colGrandTotals') == '0'
        observed.append(tuple((key, cell.value) for key, cell in sheet.cells_.items()))
    assert observed[0] == observed[1]


def test_refreshed_formula_looking_row_labels_stay_literal() -> None:
    app = prepared()
    app.add_module('''Sub ChangeLabels()
Worksheets("Data").PivotTables("Shared").PivotFields("Region").Orientation=xlRowField
Worksheets("Data").Range("A2").Value="'=1+1"
End Sub''', name='LiteralLabels')
    app.run('ChangeLabels')
    app.workbook.pivot_caches[0].Refresh()
    labels = [cell for (row, col), cell in app.workbook.sheets_[0].cells_.items() if col == 7 and row >= 3]
    assert any(cell.value == '=1+1' and cell.formula == '' for cell in labels)


@pytest.mark.parametrize('failure', ['nonxml', 'dtd', 'unknown-record-attribute', 'aliased-definition', 'missing-record-link', 'bad-record-link'])
def test_imported_record_metadata_failure_is_atomic(failure: str, tmp_path: Path) -> None:
    from pyopenvba.apps.excel._pivot_refresh import existing_record_part

    app = prepared()
    imported = ExcelApplication.open(app.save(tmp_path / 'before-malformed.xlsm'))
    cache, other = imported.workbook.pivot_caches
    part = existing_record_part(cache)
    assert cache.records is None
    if failure == 'missing-record-link':
        cache.root.attrib.pop('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    elif failure == 'bad-record-link':
        cache.root.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'missing-record-reference')
    elif failure == 'aliased-definition':
        other.part = cache.part
    elif failure == 'nonxml':
        imported.workbook.package.write(part, b'not XML')
    elif failure == 'dtd':
        imported.workbook.package.write(part, b'<!DOCTYPE pivotCacheRecords><pivotCacheRecords/>')
    else:
        raw = imported.workbook.package.read(part)
        root = xml_tree(raw)
        root.set('unknownMetadata', 'retain')
        imported.workbook.package.write(part, xml_bytes(root))
    before = state(imported)
    with pytest.raises(VBAUnsupportedError):
        cache.Refresh()
    assert state(imported) == before


GROWTH_SOURCE = """Sub PrepareGrowth()
Dim ws As Worksheet,c As PivotCache,p As PivotTable,q As PivotTable,f As PivotField,r As Long
Set ws=ActiveSheet
ws.Name="Data"
ws.Range("A1:C1").Value=Array("Key","N","Tag")
For r=2 To 9
ws.Cells(r,1).Value2="A"
ws.Cells(r,2).Value2=r-1
ws.Cells(r,3).Value2="x"
Next
Set c=ThisWorkbook.PivotCaches.Create(xlDatabase,ws.Range("A1:C9"))
Set p=c.CreatePivotTable(ws.Range("E8"),"ColumnReport")
p.PivotFields("Key").Orientation=xlColumnField
Set f=p.AddDataField(p.PivotFields("N"),"Total",xlSum)
End Sub
Sub AddRowReport()
Dim c As PivotCache,q As PivotTable,f As PivotField
Set c=ThisWorkbook.PivotCaches(1)
Set q=c.CreatePivotTable(Range("H3"),"RowReport")
q.PivotFields("Key").Orientation=xlRowField
Set f=q.AddDataField(q.PivotFields("N"),"TotalRows",xlSum)
End Sub
Sub GrowKeys()
Dim r As Long
For r=2 To 9
Cells(r,1).Value2="K" & CStr(r-1)
Next
End Sub
"""


def growth_prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(GROWTH_SOURCE, name='GrowthSetup')
    app.run('PrepareGrowth')
    return app


def test_refresh_admits_empty_report_expansion() -> None:
    app = growth_prepared()
    app.run('GrowKeys')
    app.workbook.pivot_caches[0].Refresh()
    table = app.workbook.sheets_[0].pivot_tables[0]
    location = table.root.find(NS + 'location')
    assert location is not None and location.get('ref') == 'E8:N10'


@pytest.mark.parametrize('collision', ['planned-siblings', 'unrelated-page', 'populated-growth'])
def test_refresh_growth_footprint_failure_is_atomic(collision: str) -> None:
    from pyopenvba.apps.excel._model import Cell

    app = growth_prepared()
    sheet = app.workbook.sheets_[0]
    if collision == 'planned-siblings':
        app.run('AddRowReport')
    elif collision == 'unrelated-page':
        # A distinct-cache report body is outside the expanded report, but its
        # page selector occupies J9:K9 inside the proposed E8:N10 output.
        app.add_module("""Sub OtherPage()
Dim c As PivotCache,p As PivotTable,f As PivotField
Set c=ThisWorkbook.PivotCaches.Create(xlDatabase,Range("A1:C9"))
Set p=c.CreatePivotTable(Range("J11"),"OtherPageReport")
Set f=p.AddDataField(p.PivotFields("N"),"OtherTotal",xlSum)
End Sub""", name='OtherPageSetup')
        app.run('OtherPage')
        other = next(table for table in sheet.pivot_tables if table.root.get('name') == 'OtherPageReport')
        pages = XmlElement(NS + 'pageFields', {'count': '1'})
        pages.append(XmlElement(NS + 'pageField', {'fld': '2', 'hier': '-1'}))
        other.root.append(pages)
        location = other.root.find(NS + 'location')
        assert location is not None
        location.set('rowPageCount', '1')
        location.set('colPageCount', '1')
        sheet.cells_[9, 10] = Cell(value='Tag')
        sheet.cells_[9, 11] = Cell(value='(All)')
    else:
        sheet.cells_[8, 10] = Cell(value=777)
    app.run('GrowKeys')
    cache = app.workbook.pivot_caches[0]
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        cache.Refresh()
    assert state(app) == before


@pytest.mark.parametrize('limit', ['reports', 'work', 'source-text', 'metadata-text'])
def test_transaction_budget_refuses_before_materialization(limit: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _pivot_create, _pivot_refresh

    app = prepared()
    cache = app.workbook.pivot_caches[0]
    if limit == 'reports':
        monkeypatch.setattr(_pivot_refresh, 'MAX_REPORTS', 2)
    elif limit == 'work':
        monkeypatch.setattr(_pivot_refresh, 'MAX_WORK', 5)
    elif limit == 'source-text':
        monkeypatch.setattr(_pivot_refresh, 'MAX_TEXT', 50_000)
        for row in range(2, 5):
            app.workbook.sheets_[0].cells_[row, 1].value = 'long-category-' * 1_000
    else:
        monkeypatch.setattr(_pivot_refresh, 'MAX_TEXT', 50_000)
        cache.root.set('refreshedBy', 'metadata' * 10_000)

    def forbidden(*args: object, **kwargs: object) -> None:
        pytest.fail('transaction admission must precede source materialization')

    monkeypatch.setattr(_pivot_create, 'plan_materialize', forbidden)
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        cache.Refresh()
    assert state(app) == before


def test_transaction_staged_cell_limit_refuses_before_second_report_aggregation(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _pivot_aggregate, _pivot_refresh

    app = prepared()
    cache = app.workbook.pivot_caches[0]
    # Both linked reports individually occupy two cells. The transaction limit
    # admits the first plan but must stop the second before building its output.
    monkeypatch.setattr(_pivot_refresh, 'MAX_STAGED_CELLS', 3)
    aggregate = _pivot_aggregate.aggregate
    calls: list[object] = []

    def counted(entries: list[XmlElement], operation: str) -> object:
        calls.append(None)
        return aggregate(entries, operation)

    monkeypatch.setattr(_pivot_aggregate, 'aggregate', counted)
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        cache.Refresh()
    assert len(calls) == 1
    assert state(app) == before


def test_materialized_text_is_charged_before_record_encoding(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _pivot_create, _pivot_refresh

    app = prepared()
    budget = _pivot_refresh.TextBudget(2)
    budget.remaining = 100
    app.workbook.sheets_[0].cells_[2, 1].value = 'long-category' * 100

    def forbidden(value: object, format_code: str) -> XmlElement:
        pytest.fail('the first over-budget record must be refused before encoding')

    monkeypatch.setattr(_pivot_create, 'record_value', forbidden)
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        _pivot_create.plan_materialize(app.workbook.pivot_caches[0], charge_text=budget.source_value)
    assert state(app) == before


@pytest.mark.parametrize('source_ref', ['A1:B4', 'ValidNamedSource'])
def test_imported_source_binding_cannot_be_redirected_by_defined_names(source_ref: str, tmp_path: Path) -> None:
    from zipfile import ZIP_DEFLATED, ZipFile

    app = prepared()
    app.add_module("""Sub OtherSource()
Worksheets.Add(After:=Worksheets(1)).Name="Other"
Range("A1:B1").Value=Array("Region","Amount")
Range("A2:B2").Value=Array("East",9000)
Range("A3:B3").Value=Array("West",9001)
Range("A4:B4").Value=Array("East",9002)
End Sub""", name='OtherSourceSetup')
    app.run('OtherSource')
    saved = app.save(tmp_path / 'source-before.xlsm')
    changed = tmp_path / 'source-with-name.xlsm'
    with ZipFile(saved) as original, ZipFile(changed, 'w', ZIP_DEFLATED) as output:
        for entry in original.infolist():
            data = original.read(entry.filename)
            if entry.filename == 'xl/workbook.xml':
                name = f'<definedNames><definedName name="{source_ref}">Other!$A$1:$B$4</definedName></definedNames>'
                data = data.replace(b'</sheets>', b'</sheets>' + name.encode(), 1)
            output.writestr(entry, data)
    imported = ExcelApplication.open(changed, with_vba=False)
    cache = imported.workbook.pivot_caches[0]
    assert any(entry.name == source_ref for entry in imported.workbook.names_.entries)
    if source_ref == 'ValidNamedSource':
        origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
        assert origin is not None
        origin.set('ref', source_ref)
        before = state(imported)
        with pytest.raises(VBAUnsupportedError):
            cache.Refresh()
        assert state(imported) == before
    else:
        cache.Refresh()
        assert imported.workbook.sheets_[0].cells_[4, 4].value == 60



def test_repeated_values_captions_decode_once_per_plan(monkeypatch: pytest.MonkeyPatch) -> None:
    from copy import deepcopy

    from pyopenvba.apps.excel import _pivot_layout
    from pyopenvba.apps.excel._pivot_aggregate import records
    from pyopenvba.apps.excel._pivot_create import plan_materialize

    app = growth_prepared()
    app.run('GrowKeys')
    table = app.workbook.sheets_[0].pivot_tables[0]
    staged = plan_materialize(table.cache)
    normalized = _pivot_layout.cache_layout(staged, records(staged))
    definition = deepcopy(table.root)
    columns = definition.find(NS + 'colFields')
    assert columns is not None
    definition.remove(columns)
    rows = XmlElement(NS + 'rowFields', {'count': '2'})
    rows.append(XmlElement(NS + 'field', {'x': '0'}))
    rows.append(XmlElement(NS + 'field', {'x': '-2'}))
    definition.append(rows)
    data = definition.find(NS + 'dataFields')
    assert data is not None
    data[0].set('name', 'caption-one')
    data.append(deepcopy(data[0]))
    data[1].set('name', 'caption-two')
    data.set('count', '2')
    seen: list[str] = []
    decode = _pivot_layout.decode_text

    def counted(value: str) -> str:
        seen.append(value)
        return decode(value)

    monkeypatch.setattr(_pivot_layout, 'decode_text', counted)
    plan = _pivot_layout.plan_render(table, definition, data=list(data), cache=staged, normalized=normalized)
    assert len(plan.cells) > 20
    assert seen.count('caption-one') == seen.count('caption-two') == 1
    # K1 is read once more to find the report's own item, A, which the refreshed values no longer spell.
    assert seen.count('K1') == 2
