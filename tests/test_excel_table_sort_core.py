"""Transactional boundaries of the measured local table sort subset."""
from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication, _table_sort
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._table_sort import TableSortObject
from pyopenvba.exceptions import VBAUnsupportedError


def prepared() -> tuple[ExcelApplication, TableSortObject]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Prepare()
Dim w As Worksheet, t As ListObject
Set w=ActiveSheet
w.Range("A1").Value="Key": w.Range("B1").Value="Payload"
w.Range("A2").Value=3: w.Range("A3").Value=1: w.Range("A4").Value=2
w.Range("B2").Formula="=A2*10": w.Range("B3").Formula="=A3*10": w.Range("B4").Formula="=A4*10"
Set t=w.ListObjects.Add(xlSrcRange,w.Range("A1:B4"),,xlYes)
t.Sort.SortFields.Add t.ListColumns(1).DataBodyRange
End Sub''', name="Setup")
    app.run("Prepare")
    return app, TableSortObject(app.workbook.sheets_[0].tables[0])


def snapshot(sort: TableSortObject) -> object:
    sheet, table = sort.sheet, sort.table
    return ([(key, id(cell), repr(cell)) for key, cell in sheet.cells_.items()],
            table.xml, table.changed, repr(_table_sort.state_for(table).applied), _table_sort.state_for(table).dirty,
            sheet.dirty, tuple(sort.state.fields))


@pytest.mark.parametrize("limit", ["MAX_CELLS", "MAX_WORK", "MAX_TEXT"])
def test_work_limits_refuse_before_publishing_cells_or_sort_metadata(monkeypatch: pytest.MonkeyPatch, limit: str) -> None:
    _app, sort = prepared()
    before = snapshot(sort)
    monkeypatch.setattr(_table_sort, limit, 1)
    with pytest.raises(VBAUnsupportedError, match="budget"):
        sort.Apply()
    assert snapshot(sort) == before


def test_later_moved_formula_failure_leaves_original_cell_identities_and_metadata() -> None:
    _app, sort = prepared()
    sort.sheet.cells_[4, 2].formula = 'A4+"unterminated'
    before = snapshot(sort)
    with pytest.raises(VBAUnsupportedError, match="formula"):
        sort.Apply()
    assert snapshot(sort) == before


@pytest.mark.parametrize("markup", [
    '<extLst><extension xmlns="urn:foreign"/></extLst>',
    '<autoFilter ref="A1:B4"><filterColumn colId="0"/></autoFilter>',
    '<sortState ref="A2:B4" custom="opaque"><sortCondition ref="A2:A4"/></sortState>',
    '<sortState ref="A2:B4"/><sortState ref="A2:B4"/>',
])
def test_unmeasured_imported_metadata_refuses_without_losing_original_xml(markup: str) -> None:
    _app, sort = prepared()
    sort.table.xml = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + markup + '</table>'
    before = snapshot(sort)
    with pytest.raises(VBAUnsupportedError):
        sort.Apply()
    assert snapshot(sort) == before
    xml_tree(sort.table.xml)


@pytest.mark.parametrize("operation", [
    'ActiveSheet.Rows(1).Insert',
    'ActiveSheet.Rows(3).Delete',
    'ActiveSheet.Columns(1).Insert',
    'ActiveSheet.ListObjects(1).Resize ActiveSheet.Range("A1:C4")',
    'ActiveSheet.ListObjects(1).ListRows(1).Delete',
    'ActiveSheet.Range("A1:B4").Cut ActiveSheet.Range("F1")',
])
def test_unmeasured_structural_paths_are_atomic(operation: str) -> None:
    app, sort = prepared()
    before = snapshot(sort)
    app.add_module("Sub Change()\n" + operation + "\nEnd Sub", name="Change")
    with pytest.raises(VBAUnsupportedError):
        app.run("Change")
    assert snapshot(sort) == before


@pytest.mark.parametrize("operation", [
    'ActiveSheet.ListObjects(1).Resize ActiveSheet.Range("A1:D5")',
    'ActiveSheet.ListObjects(1).ListColumns(1).Delete',
    'ActiveSheet.ListObjects(1).Delete',
    'ActiveSheet.Range("A1:C5").Cut ActiveSheet.Range("F1")',
    'ActiveSheet.Range("F1:H5").Cut ActiveSheet.Range("A1")',
])
def test_imported_sort_metadata_guards_edits_before_first_sort_getter(operation: str) -> None:
    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/table_sort/independent-table-sheet-settings.xlsx',
                                with_vba=False)
    table = app.workbook.sheets_[0].tables[0]
    assert table.sort_state is None
    before = (table.area, table.xml, table.changed,
              [(key, id(cell), repr(cell)) for key, cell in table.sheet.cells_.items()])
    app.add_module('Sub Change()\n' + operation + '\nEnd Sub', name='Change')
    with pytest.raises(VBAUnsupportedError):
        app.run('Change')
    assert table.sort_state is None
    assert (table.area, table.xml, table.changed,
            [(key, id(cell), repr(cell)) for key, cell in table.sheet.cells_.items()]) == before


def test_fixed_width_growth_preserves_held_sort_and_key_span() -> None:
    app, sort = prepared()
    field = sort.state.fields[0]
    app.add_module('''Sub Grow()
Dim t As ListObject
Set t=ActiveSheet.ListObjects(1)
t.Resize ActiveSheet.Range("A1:B5")
t.ListRows.Add
End Sub''', name='Grow')
    app.run('Grow')
    assert TableSortObject(sort.table).state is sort.state
    assert sort.state.fields[0] is field
    assert field.key.first.address() == '$A$2:$A$4'
    assert sort.table.area.address() == '$A$1:$B$6'
    sort.Apply()


def test_same_column_key_outside_table_rows_is_refused_atomically() -> None:
    app, sort = prepared()
    app.add_module('''Sub OutsideKey()
With ActiveSheet.ListObjects(1).Sort
.SortFields.Clear
.SortFields.Add ActiveSheet.Range("A100")
End With
End Sub''', name='Outside')
    app.run('OutsideKey')
    before = snapshot(sort)
    with pytest.raises(VBAUnsupportedError, match='outside the table rows'):
        sort.Apply()
    assert snapshot(sort) == before


@pytest.mark.parametrize('xml', [
    '<s:table xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ref="A1:B4"/>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ref = "A1:B4"/>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="unmeasured>name"/>',
    """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xr:uid='opaque ref="keep-opaque"'/>""",
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><!-- <table ref="phantom"/> --></table>',
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns><![CDATA[<tableColumn name="phantom"/>]]></tableColumns></table>',
])
def test_unmeasured_lexical_table_metadata_is_preserved_on_refusal(xml: str) -> None:
    _app, sort = prepared()
    sort.table.xml = xml
    before = snapshot(sort)
    with pytest.raises(VBAUnsupportedError):
        sort.Apply()
    assert snapshot(sort) == before


def test_structural_edit_after_configured_table_keeps_keys_and_rows() -> None:
    app, sort = prepared()
    before = snapshot(sort)
    app.add_module('Sub InsertAfter()\nActiveSheet.Rows(20).Insert\nEnd Sub', name='After')
    app.run('InsertAfter')
    assert snapshot(sort) == before


@pytest.mark.parametrize('acquire_sort_view', [False, True])
def test_wrapped_sort_metadata_is_guarded_before_and_after_view_acquisition(acquire_sort_view: bool) -> None:
    app, sort = prepared()
    table, sheet = sort.table, sort.sheet
    table.xml = ('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                 '<extLst><sortState ref="A2:B4"><sortCondition ref="A2:A4"/></sortState></extLst></table>')
    table.sort_state = None
    if acquire_sort_view:
        TableSortObject(table)
    before = (table.area, table.xml, [(key, id(cell), repr(cell)) for key, cell in sheet.cells_.items()])
    app.add_module('Sub MoveRows()\nActiveSheet.Rows(1).Insert\nEnd Sub', name='MoveRows')
    with pytest.raises(VBAUnsupportedError):
        app.run('MoveRows')
    assert (table.area, table.xml, [(key, id(cell), repr(cell)) for key, cell in sheet.cells_.items()]) == before


@pytest.mark.parametrize('body', [
    'ActiveSheet.Sort.SortFields(1).Order=3',
    'ActiveSheet.Sort.SetRange Worksheets.Add.Range("D1:E4")',
])
def test_unmeasured_worksheet_sort_persistence_refuses_wrong_binding_or_enum(body: str, tmp_path: Path) -> None:
    app, _sort = prepared()
    app.add_module('''Sub Configure()
ActiveSheet.Sort.SetRange ActiveSheet.Range("D1:E4")
ActiveSheet.Sort.SortFields.Add ActiveSheet.Range("D2:D4")
''' + body + '\nEnd Sub', name='Config')
    app.run('Configure')
    with pytest.raises(VBAUnsupportedError):
        app.save(tmp_path / 'unsupported-binding.xlsm')


@pytest.mark.parametrize('boundary', ['table-clear-add', 'worksheet-orientation', 'worksheet-foreign-same-address'])
def test_sort_save_admission_precedes_package_writes(boundary: str, tmp_path: Path) -> None:
    app, sort = prepared()
    sort.Apply()
    if boundary == 'worksheet-foreign-same-address':
        app.add_module('''Sub ConfigureSheetSort()
With ActiveSheet.Sort
.SetRange ActiveSheet.Range("D1:E4")
.SortFields.Add ActiveSheet.Range("D2:D4")
End With
End Sub''', name='SheetSort')
        app.run('ConfigureSheetSort')
    app.save(tmp_path / 'baseline.xlsm')
    if boundary == 'table-clear-add':
        fields = sort.SortFields()
        fields.Clear()
        fields.Add(sort.sheet.Range('A2:A4'))
    else:
        body = ('ActiveSheet.Sort.SetRange Worksheets.Add.Range("D1:E4")'
                if boundary == 'worksheet-foreign-same-address' else
                'With ActiveSheet.Sort\n.SetRange ActiveSheet.Range("D1:E4")\n.SortFields.Add ActiveSheet.Range("D2:D4")\n.Orientation=2\nEnd With')
        app.add_module('Sub ChangeSort()\n' + body + '\nEnd Sub', name='ChangeSort')
        app.run('ChangeSort')
    app.add_module('Sub DirtySheet()\nRange("Z1").Value="pending change"\nEnd Sub', name='DirtySheet')
    app.run('DirtySheet')
    package = app.workbook.package
    assert package is not None
    before = package.serialize()
    rejected = tmp_path / 'rejected.xlsm'
    with pytest.raises(VBAUnsupportedError):
        app.save(rejected)
    assert package.serialize() == before
    assert not rejected.exists()


@pytest.mark.parametrize('position', ['group-text', 'condition-text', 'condition-tail'])
def test_nonwhitespace_sort_metadata_refuses_apply_and_save_without_losing_input(
        position: str, tmp_path: Path) -> None:
    native = Path(__file__).parent / 'fixtures/table_sort/ascending-formulas-styles-totals.xlsx'
    derived = tmp_path / 'nonwhitespace.xlsx'
    with ZipFile(native) as source, ZipFile(derived, 'w') as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == 'xl/tables/table1.xml':
                xml = data.decode('utf8')
                if position == 'group-text':
                    xml, count = re.subn(r'(<sortState\b[^>]*>)', r'\1unmodeled-text', xml, count=1)
                elif position == 'condition-text':
                    xml, count = re.subn(r'<sortCondition\b([^>]*)/>',
                                         r'<sortCondition\1>unmodeled-text</sortCondition>', xml, count=1)
                else:
                    xml, count = re.subn(r'(<sortCondition\b[^>]*/>)', r'\1unmodeled-text', xml, count=1)
                assert count == 1
                data = xml.encode()
                xml_tree(data)
            target.writestr(info, data)
    app = ExcelApplication.open(derived, with_vba=False)
    sort = TableSortObject(app.workbook.sheets_[0].tables[0])
    package = app.workbook.package
    assert package is not None
    before, package_before = snapshot(sort), package.serialize()
    with pytest.raises(VBAUnsupportedError):
        sort.Apply()
    assert snapshot(sort) == before
    assert package.serialize() == package_before
    sort.vba_set('MatchCase', True)
    sort.sheet.Range('Z1').vba_set('Value', 'pending change')
    package_before = package.serialize()
    format_before = app.workbook.format_code
    rejected = tmp_path / 'rejected-nonwhitespace.xlsm'
    with pytest.raises(VBAUnsupportedError):
        app.save(rejected)
    assert package.serialize() == package_before
    assert app.workbook.format_code == format_before
    assert not rejected.exists()
