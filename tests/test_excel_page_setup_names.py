"""Scoped print-name ownership and quote-safe preserved reserved-name remapping."""
from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication, _io
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._page_setup import PageSetup
from pyopenvba.exceptions import VBAUnsupportedError

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
SPELLINGS = ["localSheetId='1' x:opaque='keep localSheetId=\"0\" > tail'",
             "x:opaque='keep localSheetId=\"0\" > tail' localSheetId=\"1\"",
             "localSheetId='1' x:localSheetId='0' x:opaque='keep localSheetId=\"0\" > tail'"]


def reserved_book(tmp_path: Path, spelling: str, bare: str) -> ExcelApplication:
    app = ExcelApplication()
    book = app.add_workbook()
    book.sheets_[0].name = 'First'
    book.add_sheet('Owner')
    book.add_sheet('Last')
    app.save(tmp_path / 'base.xlsm')
    package = book.package
    assert package is not None
    xml = package.read('xl/workbook.xml').decode('utf8')
    entry = f'<definedName xmlns:x="urn:opaque" name="{bare}" {spelling}>Owner!$A$1:$B$3</definedName>'
    xml = xml.replace('</sheets>', '</sheets><definedNames>' + entry + '</definedNames>')
    package.write('xl/workbook.xml', xml.encode('utf8'))
    path = tmp_path / 'controlled-input.xlsm'
    path.write_bytes(package.serialize())
    return ExcelApplication.open(path, with_vba=False)


@pytest.mark.parametrize('bare', ['_xlnm.Print_Area', '_xlnm.UnmodeledPrint'])
@pytest.mark.parametrize('spelling', SPELLINGS)
@pytest.mark.parametrize('operation', ['insert-first', 'move-owner', 'delete-first', 'delete-owner'])
def test_reserved_scope_and_opaque_metadata_survive_sheet_lifecycle(
        tmp_path: Path, bare: str, spelling: str, operation: str) -> None:
    app = reserved_book(tmp_path, spelling, bare)
    book = app.workbook
    first, owner, _last = book.sheets_
    if operation == 'insert-first':
        book.add_sheet('Inserted', at=0)
    elif operation == 'move-owner':
        owner.Move(Before=first)
    elif operation == 'delete-first':
        first.Delete()
    else:
        owner.Delete()
    path = app.save(tmp_path / 'lifecycle.xlsm')
    with ZipFile(path) as package:
        root = xml_tree(package.read('xl/workbook.xml'))
    names = [entry for entry in root.findall(NS + 'definedNames/' + NS + 'definedName') if entry.get('name') == bare]
    if operation == 'delete-owner':
        assert names == []
    else:
        assert len(names) == 1 and names[0].get('localSheetId') == str(book.sheets_.index(owner))
        assert names[0].get('{urn:opaque}opaque') == 'keep localSheetId="0" > tail'
        assert names[0].text == 'Owner!$A$1:$B$3'
        if 'x:localSheetId' in spelling:
            assert names[0].get('{urn:opaque}localSheetId') == '0'
        reopened = ExcelApplication.open(path, with_vba=False)
        if bare == '_xlnm.Print_Area':
            assert PageSetup(reopened.workbook.sheet_named('Owner')).PrintArea() == '$A$1:$B$3'


@pytest.mark.parametrize('scope', ['-1', 'bad', '999', '1' * 1000])
def test_reserved_scope_validation_precedes_any_package_or_sheet_part_mutation(tmp_path: Path, scope: str) -> None:
    app = reserved_book(tmp_path, f"localSheetId='{scope}'", '_xlnm.UnmodeledPrint')
    book = app.workbook
    book.add_sheet('Unpublished', at=0)
    package = book.package
    assert package is not None
    before = package.serialize(), [(sheet.part_name, sheet.name) for sheet in book.sheets_]
    sync = getattr(_io, '_sync_sheets')
    with pytest.raises(VBAUnsupportedError, match='invalid sheet scope'):
        sync(book, package)
    assert (package.serialize(), [(sheet.part_name, sheet.name) for sheet in book.sheets_]) == before


def test_owned_titles_keep_each_axis_and_follow_rename_move_delete(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    owner = book.sheets_[0]
    other = book.add_sheet('Other')
    page = PageSetup(owner)
    page.vba_set('PrintArea', 'A1:B9,D2:E4')
    page.vba_set('PrintTitleRows', '2:3')
    page.vba_set('PrintTitleColumns', 'B:C')
    page.vba_set('PrintTitleRows', '')
    assert page.PrintTitleRows() == '' and page.PrintTitleColumns() == '$B:$C'
    owner.vba_set('Name', "O'Brien Data")
    owner.Move(After=other)
    saved = app.save(tmp_path / 'moved.xlsm')
    reopened = ExcelApplication.open(saved)
    moved = reopened.workbook.sheet_named("O'Brien Data")
    assert PageSetup(moved).PrintArea() == '$A$1:$B$9,$D$2:$E$4'
    assert PageSetup(moved).PrintTitleColumns() == '$B:$C'
    moved.Delete()
    assert all("O''Brien Data" not in entry.name for entry in reopened.workbook.names_.entries)
    final = reopened.save(tmp_path / 'deleted.xlsm')
    with ZipFile(final) as package:
        assert '_xlnm.Print_Area' not in package.read('xl/workbook.xml').decode()
        assert '_xlnm.Print_Titles' not in package.read('xl/workbook.xml').decode()


@pytest.mark.parametrize('imported', [False, True])
@pytest.mark.parametrize('property_name,value', [('LeftMargin', 18), ('Orientation', 2), ('PrintArea', 'A1:B9')])
def test_configured_copy_is_refused_before_destination_publication(
        tmp_path: Path, imported: bool, property_name: str, value: object) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    source = book.sheets_[0]
    PageSetup(source).vba_set(property_name, value)
    if imported:
        app = ExcelApplication.open(app.save(tmp_path / 'configured.xlsm'))
        book = app.workbook
        source = book.sheets_[0]
    PageSetup(source).LeftMargin()
    package_bytes = book.package.serialize() if book.package is not None else None
    before = list(book.sheets_), list(app.application.workbooks_.books), tuple(book.names_.entries), package_bytes
    with pytest.raises(VBAUnsupportedError, match='PageSetup'):
        source.Copy(After=source)
    after = (list(book.sheets_), list(app.application.workbooks_.books), tuple(book.names_.entries),
             book.package.serialize() if book.package is not None else None)
    assert after == before
    with pytest.raises(VBAUnsupportedError, match='PageSetup'):
        source.Copy()
    assert len(app.application.workbooks_.books) == 1


def test_default_equivalent_imported_configuration_copy_remains_available(tmp_path: Path) -> None:
    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/page_setup/print-names-clear.xlsx', with_vba=False)
    source = app.workbook.sheets_[0]
    assert PageSetup(source).LeftMargin() == 50.4
    source.Copy(After=source)
    assert len(app.workbook.sheets_) == 2
    assert PageSetup(app.workbook.sheets_[1]).LeftMargin() == 50.4


def test_removed_sheet_page_proxy_cannot_recreate_global_print_names() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    owner = book.sheets_[0]
    book.add_sheet('Survivor')
    page = PageSetup(owner)
    page.vba_set('PrintArea', 'A1:B9')
    owner.Delete()
    before = list(book.names_.entries), book.saved
    with pytest.raises(VBAUnsupportedError, match='removed worksheet'):
        page.vba_set('PrintArea', 'A1:C9')
    assert (list(book.names_.entries), book.saved) == before


def test_generic_delete_of_owned_print_name_is_not_resurrected_on_save(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._model import DefinedName

    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/page_setup/print-names-quoted.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    entry = next(entry for entry in app.workbook.names_.entries if entry.name.endswith('!Print_Area'))
    DefinedName(entry).Delete()
    assert PageSetup(sheet).PrintArea() == ''
    saved = app.save(tmp_path / 'deleted-print-name.xlsm')
    with ZipFile(saved) as package:
        xml = package.read('xl/workbook.xml').decode()
        assert '_xlnm.Print_Area' not in xml
        assert '_xlnm.Print_Titles' in xml
    reopened = ExcelApplication.open(saved)
    assert PageSetup(reopened.workbook.sheets_[0]).PrintArea() == ''
