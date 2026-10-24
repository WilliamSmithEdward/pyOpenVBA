"""Native macro-free and macro-enabled conditional load/conversion provenance."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, RANGE_RULES, FormatConditions, read_rules, with_rules
from pyopenvba.apps.excel._io import workbook_format
from pyopenvba.apps.excel._model import Range, Workbook
from pyopenvba.interpreter._objects import VBAObject

ROOT = Path(__file__).parent / 'fixtures/conditional_transfer'
CASES: list[dict[str, Any]] = [case for case in json.loads((ROOT / 'controls.json').read_text(encoding='utf-8'))['probes']
                              if case['name'] in ('copy-interior-clip', 'duplicates-copy-other-sheet')]
CONVERSIONS: list[dict[str, Any]] = json.loads((ROOT / 'native52_provenance.json').read_text(encoding='utf-8'))['conversions']


def untouched(book: Workbook) -> object:
    return [[(id(rule), list(rule.areas), xml_bytes(rule.root), rule.imported)
             for rule in sheet.conditional_formats] for sheet in book.sheets_]


def geometry(path: Path) -> object:
    with ZipFile(path) as archive:
        return [[(group.get('sqref'), [(rule.get('type'), [child.text for child in rule.findall(NS + 'formula')])
                                       for rule in group.findall(NS + 'cfRule')])
                 for group in xml_tree(archive.read(name)).findall(NS + 'conditionalFormatting')]
                for name in sorted(archive.namelist()) if name.startswith('xl/worksheets/sheet') and name.endswith('.xml')]


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
@pytest.mark.parametrize('file_format', [51, 52])
def test_native_format_loading_projects_only_macro_free_areas_without_dirtying(
        case: dict[str, Any], file_format: int) -> None:
    app = ExcelApplication.open(ROOT / case['file' if file_format == 51 else 'native52_file'], with_vba=False)
    book = app.workbook
    assert workbook_format(book) == file_format and book.saved
    before = untouched(book)
    for sheet in book.sheets_:
        assert not sheet.dirty and not sheet.conditional_changed
        for rule in sheet.conditional_formats:
            assert rule.imported is (file_format == 51)
            collection = FormatConditions(Range(sheet, list(rule.areas)))
            wrapper = next(item for item in collection.vba_items()
                           if isinstance(item, VBAObject) and item.vba_get('Priority') == int(rule.root.get('priority', '0')))
            assert isinstance(wrapper, VBAObject)
            applies = wrapper.vba_get('AppliesTo')
            assert isinstance(applies, Range)
            generic = rule.root.get('type') not in RANGE_RULES
            expected = (sorted(rule.areas, key=lambda area: (area.top, area.bottom, area.left, area.right))
                        if file_format == 51 and generic else rule.areas)
            assert [(area.top, area.left, area.bottom, area.right) for area in applies.areas] == [
                (area.top, area.left, area.bottom, area.right) for area in expected]
    assert untouched(book) == before and book.saved
    assert all(not sheet.dirty and not sheet.conditional_changed for sheet in book.sheets_)


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
@pytest.mark.parametrize('inspection_vba', [False, True], ids=['without-vba-project', 'with-inspection-vba'])
def test_native_macro_free_to_enabled_conversion_serializes_loaded_order_only(
        case: dict[str, Any], inspection_vba: bool, tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    book = app.workbook
    before = untouched(book)
    native = next(record for record in CONVERSIONS if record['name'] == case['name'])
    # The native conversion contained its inspection project before SaveAs52.
    # Independently measured no-VBA52 controls retain the plain load mode; the
    # extension/format code alone does not establish a VBA-enabled load.
    with ZipFile(ROOT / native['file']) as archive:
        assert 'xl/vbaProject.bin' in archive.namelist()
    if inspection_vba:
        app.add_module(native['inspection_source'], name='NativeInspection')
    result = app.save(tmp_path / 'converted.xlsm')
    with ZipFile(result) as archive:
        assert ('xl/vbaProject.bin' in archive.namelist()) is inspection_vba
    assert geometry(result) == geometry(ROOT / native['file'])
    assert untouched(book) == before
    assert workbook_format(book) == 52
    reopened = ExcelApplication.open(result, with_vba=False)
    assert all(rule.imported is (not inspection_vba) and rule.plain_load is (not inspection_vba)
               for sheet in reopened.workbook.sheets_ for rule in sheet.conditional_formats)
    reopened.add_module(native['inspection_source'], name='NativeInspection')
    assert reopened.run('Snapshot') == native['reopened']


@pytest.mark.parametrize('projected', [True, False], ids=['macro-free-load', 'macro-enabled-load'])
def test_generic_area_and_group_orders_remain_independent_without_reanchoring_storage(projected: bool) -> None:
    xml = f'''<worksheet xmlns="{NS[1:-1]}"><sheetData/>
<conditionalFormatting sqref="E3:E4"><cfRule type="expression" priority="1"><formula>E3&gt;0</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="K3:K5 B9:B10"><cfRule type="expression" priority="2"><formula>B3&gt;0</formula></cfRule></conditionalFormatting>
</worksheet>'''
    rules = read_rules(xml, imported=projected)
    before = [(list(rule.areas), xml_bytes(rule.root)) for rule in rules]
    rewritten = xml_tree(with_rules(xml, rules))
    actual = [group.get('sqref') for group in rewritten.findall(NS + 'conditionalFormatting')]
    assert actual == ['E3:E4', 'K3:K5 B9:B10']
    assert [(list(rule.areas), xml_bytes(rule.root)) for rule in rules] == before


def test_loaded_statistical_group_uses_original_first_population_area_for_column_identity() -> None:
    xml = f'''<worksheet xmlns="{NS[1:-1]}"><sheetData/>
<conditionalFormatting sqref="M9:M10 K3:K5"><cfRule type="top10" priority="1" rank="2"/></conditionalFormatting>
<conditionalFormatting sqref="L1:L2"><cfRule type="expression" priority="2"><formula>L1&gt;0</formula></cfRule></conditionalFormatting>
</worksheet>'''
    rules = read_rules(xml)
    rewritten = xml_tree(with_rules(xml, rules))
    assert [group.get('sqref') for group in rewritten.findall(NS + 'conditionalFormatting')] == ['L1:L2', 'M9:M10 K3:K5']


@pytest.mark.parametrize('areas', ['B9:B10 K3:K5', 'K3:K5 B9:B10'])
def test_loaded_generic_group_identity_uses_first_row_projected_area_column(areas: str) -> None:
    xml = f'''<worksheet xmlns="{NS[1:-1]}"><sheetData/>
<conditionalFormatting sqref="{areas}"><cfRule type="expression" priority="1"><formula>B3&gt;0</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="E3:E4"><cfRule type="expression" priority="2"><formula>E3&gt;0</formula></cfRule></conditionalFormatting>
</worksheet>'''
    rules = read_rules(xml)
    before = [(list(rule.areas), xml_bytes(rule.root)) for rule in rules]
    rewritten = xml_tree(with_rules(xml, rules))
    assert [group.get('sqref') for group in rewritten.findall(NS + 'conditionalFormatting')] == ['E3:E4', 'K3:K5 B9:B10']
    assert [(list(rule.areas), xml_bytes(rule.root)) for rule in rules] == before


@pytest.mark.parametrize('type_args,mutation,formula', [
    ('9,String:="x",TextOperator:=0', '', 'NOT(ISERROR(SEARCH("x",B3)))'),
    ('9,String:="x",TextOperator:=0', 'f.Text="y"', 'NOT(ISERROR(SEARCH("y",B3)))'),
    ('9,String:="x",TextOperator:=0', 'f.TextOperator=2', 'LEFT(B3,LEN("x"))="x"'),
    ('9,String:="x",TextOperator:=0', 'f.Modify Type:=9,String:="z",Operator2:=3', 'RIGHT(B3,LEN("z"))="z"'),
    ('9,String:="x",TextOperator:=0', 'f.ModifyAppliesToRange Union(Range("D8:D9"),Range("F4:F5"))',
     'NOT(ISERROR(SEARCH("x",D4)))'),
    ('9,String:="x",TextOperator:=0', 'Range("B9:B10").FormatConditions.Delete', 'NOT(ISERROR(SEARCH("x",K3)))'),
    ('10', '', 'LEN(TRIM(B3))=0'),
    ('16', '', 'ISERROR(B3)'),
    ('11,DateOperator:=0', '', 'FLOOR(B3,1)=TODAY()'),
    ('11,DateOperator:=0', 'f.DateOperator=6', 'FLOOR(B3,1)=TODAY()+1'),
])
def test_predicate_mutations_regenerate_against_bounding_corner(
        type_args: str, mutation: str, formula: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Probe() As String
Dim f As Object
Set f=Union(Range("B9:B10"),Range("K3:K5")).FormatConditions.Add(Type:=''' + type_args + ''')
''' + mutation + '''
Probe=f.Formula1
End Function''', name='PredicateAnchor')
    assert app.run('Probe') == '=' + formula
