"""Sheet copying must rewrite decoded formulas and preserve extended rules atomically."""
from __future__ import annotations

import os
from pathlib import Path
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._conditional import NS
from pyopenvba.apps.excel._model import Workbook
from pyopenvba.apps.excel._xstring import decode
from pyopenvba.exceptions import VBAUnsupportedError


@pytest.mark.parametrize('external', [False, True])
def test_sheet_copy_rewrites_literal_escape_in_self_sheet_name(external: bool, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    copy = 'ActiveSheet.Copy' if external else 'ActiveSheet.Copy After:=ActiveSheet'
    app.add_module('''Sub Probe()
ActiveSheet.Name="_x0041_"
Range("A1").Value=10
Range("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=_x0041_!A1>0"
''' + copy + '\nEnd Sub', name='Harness')
    app.run('Probe')
    book = app.workbooks()[-1] if external else app.workbooks()[0]
    copied = book.sheets_[-1]
    node = copied.conditional_formats[0].root.find(NS + 'formula')
    assert node is not None
    expected = "'_x0041_ (2)'!A1>0" if not external else '_x0041_!A1>0'
    assert decode(node.text or '') == expected
    reopened = ExcelApplication.open(app.save(tmp_path / 'copied.xlsm', workbook=book))
    saved = reopened.workbooks()[0].sheets_[-1].conditional_formats[0].root.find(NS + 'formula')
    assert saved is not None and decode(saved.text or '') == expected


def test_cross_book_copy_discovers_escaped_name_in_conditional_formula() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Probe()
ActiveWorkbook.Names.Add Name:="_x0041_",RefersTo:="=42"
Range("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=_x0041_>0"
ActiveSheet.Copy
End Sub''', name='Harness')
    app.run('Probe')
    assert any(entry.name == '_x0041_' and entry.refers_to == '=42'
               for entry in app.workbooks()[-1].names_.entries)


def test_copy_rejects_x14_only_rules_before_creating_sheet_or_workbook(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    original = app.save(tmp_path / 'ordinary.xlsx')
    extended = tmp_path / 'extended.xlsx'
    extension = '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormatting/></x14:conditionalFormattings></ext></extLst>'
    with ZipFile(original) as source, ZipFile(extended, 'w') as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == 'xl/worksheets/sheet1.xml':
                data = data.replace(b'</worksheet>', extension.encode() + b'</worksheet>')
            target.writestr(entry, data)
    reopened = ExcelApplication.open(extended)
    book = reopened.workbooks()[0]
    before = (len(reopened.workbooks()), len(book.sheets_))
    reopened.add_module('Sub Probe()\nActiveSheet.Copy\nEnd Sub', name='Harness')
    with pytest.raises(VBAUnsupportedError, match='extended conditional formatting'):
        reopened.run('Probe')
    assert (len(reopened.workbooks()), len(book.sheets_)) == before


def custom_format_copy() -> tuple[ExcelApplication, Workbook, Workbook]:
    app = ExcelApplication()
    source = app.add_workbook()
    app.add_module('''Sub SourceSetup()
ActiveSheet.Name="Widgets"
Range("A1").Value=20
Columns("A").ColumnWidth=30
Dim f As FormatCondition
Set f=Range("A1").FormatConditions.Add(Type:=xlExpression,Formula1:="=TRUE")
f.NumberFormat="0.00 ""widgets"""
End Sub
Sub DestinationSetup()
ActiveSheet.Name="Items"
Range("A1").Value=20
Columns("A").ColumnWidth=30
Dim f As FormatCondition
Set f=Range("A1").FormatConditions.Add(Type:=xlExpression,Formula1:="=TRUE")
f.NumberFormat="0.00 ""items"""
End Sub''', name='SetupModule')
    app.run('SourceSetup')
    destination = app.add_workbook()
    app.run('DestinationSetup')
    source.sheets_[0].Copy(After=destination.sheets_[0])
    source.sheets_[0].Copy(After=destination.sheets_[-1])
    return app, source, destination


READ_CUSTOM = '''Function ReadCustom() As String
Dim s As Worksheet
For Each s In Worksheets
ReadCustom=ReadCustom & s.Name & ":" & s.Range("A1").Text & ":" & s.Range("A1").FormatConditions(1).NumberFormat & "|"
Next
End Function'''
EXPECTED_CUSTOM = 'Items:20.00 items:0.00 "items"|Widgets:20.00 widgets:0.00 "widgets"|Widgets (2):20.00 widgets:0.00 "widgets"|'


def test_cross_book_copy_allocates_custom_differential_number_format_ids(tmp_path: Path) -> None:
    app, source, destination = custom_format_copy()
    identifiers: list[str] = []
    for sheet in [source.sheets_[0], *destination.sheets_]:
        rule = sheet.conditional_formats[0]
        assert rule.dxf is not None
        number = rule.dxf.find(NS + 'numFmt')
        assert number is not None
        identifiers.append(number.get('numFmtId', ''))
    assert identifiers == ['164', '164', '165', '165']
    reopened = ExcelApplication.open(app.save(tmp_path / 'custom-copy.xlsm', workbook=destination))
    reopened.add_module(READ_CUSTOM, name='Readback')
    assert reopened.run('ReadCustom') == EXPECTED_CUSTOM


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
def test_native_reopens_distinct_custom_formats_after_cross_book_copy(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    app, _, destination = custom_format_copy()
    path = app.save(tmp_path / 'custom-copy-native.xlsm', workbook=destination)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        result = excel.run_vba(READ_CUSTOM, 'ReadCustom', line_numbers=False)
        assert result.ok and not result.dialogs and result.value == EXPECTED_CUSTOM, result
