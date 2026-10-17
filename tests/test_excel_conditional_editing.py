"""Unimplemented CF transfers and structural moves fail before changing a workbook."""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes
from pyopenvba.apps.excel._model import Workbook
from pyopenvba.exceptions import VBAUnsupportedError


def prepare(rules: str, action: str) -> tuple[ExcelApplication, Workbook]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\nRange("A1").Value=10\nRange("D1").Value=30\n' + rules
                   + '\nEnd Sub\nSub Action()\n' + action + '\nEnd Sub', name='Harness')
    app.run('Setup')
    return app, book


def state(book: Workbook) -> object:
    return [(sheet.name, {position: (cell.value, cell.formula, cell.style) for position, cell in sheet.cells_.items()},
             [(tuple(rule.areas), xml_bytes(rule.root), rule.deleted) for rule in sheet.conditional_formats])
            for sheet in book.sheets_]


RULE = 'Range("A1:A5").FormatConditions.Add Type:=xlExpression,Formula1:="=A1>0"'


@pytest.mark.parametrize('action', [
    'Rows(3).Insert', 'Rows(1).Delete', 'Columns(1).Insert', 'Columns(1).Delete',
    'Range("A3").Insert xlShiftDown', 'Range("A3").Delete xlShiftUp',
])
def test_structural_rule_movement_is_atomic(action: str) -> None:
    app, book = prepare(RULE, action)
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before


def test_structural_reference_on_another_sheet_is_guarded() -> None:
    setup = 'ActiveSheet.Name="Data"\nWorksheets.Add.Name="Report"\nRange("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=Data!A3>0"'
    app, book = prepare(setup, 'Worksheets("Data").Rows(2).Insert')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting formulas'):
        app.run('Action')
    assert state(book) == before


@pytest.mark.parametrize('action', ['Rows(9).Insert', 'Range("D3").Delete xlShiftUp'])
def test_unrelated_structural_edits_remain_available(action: str) -> None:
    app, book = prepare(RULE, action)
    rule = book.sheets_[0].conditional_formats[0]
    before = (tuple(rule.areas), xml_bytes(rule.root))
    app.run('Action')
    assert (tuple(rule.areas), xml_bytes(rule.root)) == before


@pytest.mark.parametrize('action', [
    'Range("A1").Copy Range("D1")',
    'Range("A1").Copy\nRange("D1").PasteSpecial xlPasteFormats',
    'Range("A1").Copy\nRange("D1").PasteSpecial xlPasteAll',
    'Range("A1").Cut Range("D1")',
    'Range("A1").Cut\nActiveSheet.Paste Range("D1")',
])
def test_format_bearing_transfer_cannot_drop_source_rules(action: str) -> None:
    app, book = prepare(RULE, action)
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before


def test_copy_cannot_overwrite_destination_rule_silently() -> None:
    app, book = prepare('Range("D1").FormatConditions.Add xlCellValue,xlGreater,5',
                        'Range("A1").Copy Range("D1")')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before


def test_value_paste_keeps_conditional_rules() -> None:
    app, book = prepare(RULE, 'Range("D1").Copy\nRange("A1").PasteSpecial xlPasteValues')
    rule = book.sheets_[0].conditional_formats[0]
    before = (tuple(rule.areas), xml_bytes(rule.root))
    app.run('Action')
    assert book.sheets_[0].cells_[(1, 1)].value == 30
    assert (tuple(rule.areas), xml_bytes(rule.root)) == before


def test_unrelated_copy_on_a_sheet_with_rules_remains_available() -> None:
    app, book = prepare(RULE, 'Range("D1").Copy Range("F1")')
    app.run('Action')
    assert book.sheets_[0].cells_[(1, 6)].value == 30


def test_cut_guards_rule_references_on_another_sheet() -> None:
    setup = 'ActiveSheet.Name="Data"\nWorksheets.Add.Name="Report"\nRange("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=Data!A1>0"'
    app, book = prepare(setup, 'Worksheets("Data").Range("A1").Cut Worksheets("Data").Range("D1")')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting formulas'):
        app.run('Action')
    assert state(book) == before


def test_multi_area_copy_checks_every_area_before_pasting_first() -> None:
    app, book = prepare('Range("D3").FormatConditions.Add xlCellValue,xlGreater,5',
                        'Union(Range("D1"),Range("D3")).Copy Range("F1")')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before
