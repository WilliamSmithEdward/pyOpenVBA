"""Native conditional API spelling and area order without altering stored rules."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, FormatCondition
from pyopenvba.apps.excel._conditional_display import DisplayFormat
from pyopenvba.apps.excel._differential import number_format
from pyopenvba.apps.excel._model import Range
from pyopenvba.interpreter._values import EMPTY, to_text

ROOT = Path(__file__).parent / 'fixtures/conditional_api_projection'
CASES: list[dict[str, Any]] = json.loads((ROOT / 'native.json').read_text(encoding='utf-8'))['probes']
ANCHOR_CASES: list[dict[str, Any]] = json.loads((ROOT / 'anchors.json').read_text(encoding='utf-8'))['probes']
FAMILY_CASES: list[dict[str, Any]] = json.loads((ROOT / 'family_area_order.json').read_text(encoding='utf-8'))['probes']
PREDICATE_CASES: list[dict[str, Any]] = json.loads((ROOT / 'predicate_anchors.json').read_text(encoding='utf-8'))['probes']


def snapshot(app: ExcelApplication) -> str:
    sheet = app.workbooks()[0].sheets_[0]
    answer = ''
    displayed = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)])).NumberFormat()
    for rule in sheet.conditional_formats:
        wrapped = FormatCondition(sheet, rule)
        answer += (f'nf=[{to_text(wrapped.NumberFormat())}];'
                   f'applies=[{wrapped.AppliesTo().vba_get("Address")}];'
                   f'formula=[{wrapped.vba_get("Formula1")}];display=[{displayed}]|')
    return answer


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_import_getters_preserve_raw_rules_and_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / (case['name'] + '.xlsx'))
    sheet = app.workbooks()[0].sheets_[0]
    rule = sheet.conditional_formats[0]
    before = xml_bytes(rule.root), list(rule.areas)
    identifier = rule.root.get('dxfId')
    dxf = sheet.book.stylesheet.dxf(int(identifier)) if identifier is not None else None
    raw = None if dxf is None else xml_bytes(dxf)
    assert snapshot(app) == case['reopened']
    assert (xml_bytes(rule.root), rule.areas) == before
    assert dxf is None or xml_bytes(dxf) == raw
    assert not sheet.conditional_changed
    saved = app.save(tmp_path / 'projected.xlsx')
    reopened = ExcelApplication.open(saved)
    assert snapshot(reopened) == case['reopened']
    if dxf is not None:
        saved_rule = reopened.workbooks()[0].sheets_[0].conditional_formats[0]
        saved_id = saved_rule.root.get('dxfId')
        assert saved_id is not None
        saved_dxf = reopened.workbooks()[0].stylesheet.dxf(int(saved_id))
        assert number_format(saved_dxf) == number_format(dxf)


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_runtime_getter_projection_preserves_creation_area_order(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    literal = str(case['input']).replace('"', '""')
    if case['name'].startswith('number-format'):
        setup = 'Set f=Range("A1:A4").FormatConditions.Add(2,Formula1:="=A1>0")\nf.NumberFormat="' + literal + '"'
    else:
        setup = 'Set f=Range("' + literal + '").FormatConditions.Add(2,Formula1:="=M9>0")'
    app.add_module('Sub Setup()\nDim f As Object\n' + setup + '\nEnd Sub', name='Harness')
    app.run('Setup')
    actual = snapshot(app)
    # Formula anchoring is independent of getter ordering and covered elsewhere.
    assert actual.split(';formula=')[0] == case['before_save'].split(';formula=')[0]


def test_number_format_projection_retains_absence_and_the_files_spelling() -> None:
    app = ExcelApplication.open(ROOT / 'number-format-2.xlsx')
    sheet = app.workbooks()[0].sheets_[0]
    identifier = sheet.conditional_formats[0].root.get('dxfId')
    assert identifier is not None
    root = sheet.book.stylesheet.dxf(int(identifier))
    entry = root.find(NS + 'numFmt')
    assert entry is not None and entry.get('formatCode') == '0.00\\-\\ "source"'
    assert number_format(root) == '0.00- "source"'
    root.remove(entry)
    assert number_format(root) is EMPTY


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_samples_are_macro_free(case: dict[str, Any]) -> None:
    with ZipFile(ROOT / (case['name'] + '.xlsx')) as package:
        assert not any('vbaproject' in name.lower() or 'externallinks/' in name.lower() for name in package.namelist())


@pytest.mark.parametrize('case', ANCHOR_CASES, ids=lambda case: case['name'])
def test_native_union_formula_anchor_runtime_and_import(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Harness')
    app.run('Setup')
    assert app.run('Probe') == case['before']
    imported = ExcelApplication.open(ROOT / case['file'])
    imported.add_module(case['inspection_source'], name='Harness')
    assert imported.run('Probe') == case['reopened']
    reopened = ExcelApplication.open(imported.save(tmp_path / 'anchor.xlsm'))
    assert reopened.run('Probe') == case['reopened']


@pytest.mark.parametrize('case', FAMILY_CASES, ids=lambda case: case['name'])
def test_native_family_area_runtime_import_and_conversion(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Harness')
    app.run('Setup')
    assert app.run('Probe') == case['before']
    imported = ExcelApplication.open(ROOT / case['file'])
    imported.add_module(case['inspection_source'], name='Readback')
    assert imported.run('Readback.Probe') == case['reopened51']
    converted_file = imported.save(tmp_path / 'family-converted.xlsm')
    converted = ExcelApplication.open(converted_file)
    assert converted.run('Readback.Probe') == case['converted52']
    with ZipFile(converted_file) as package:
        generated = xml_tree(package.read('xl/worksheets/sheet1.xml').decode('utf-8'))
    native = xml_tree(case['converted_sheet_xml'])
    assert [group.get('sqref') for group in generated.findall(NS + 'conditionalFormatting')] == [
        group.get('sqref') for group in native.findall(NS + 'conditionalFormatting')]


@pytest.mark.parametrize('case', PREDICATE_CASES, ids=lambda case: case['name'])
def test_native_predicate_union_bounding_anchor(case: dict[str, Any], tmp_path: Path,
                                              monkeypatch: pytest.MonkeyPatch) -> None:
    moment = dt.datetime(2026, 10, 20, 12)
    app = ExcelApplication()
    app.add_workbook()
    monkeypatch.setattr(app.interpreter, 'clock', lambda: moment)
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Harness')
    app.run('Setup')
    assert app.run('Probe') == case['before']
    imported = ExcelApplication.open(ROOT / case['file'])
    monkeypatch.setattr(imported.interpreter, 'clock', lambda: moment)
    imported.add_module(case['inspection_source'], name='Readback')
    assert imported.run('Readback.Probe') == case['reopened']
    converted = ExcelApplication.open(imported.save(tmp_path / 'predicate-union.xlsm'))
    monkeypatch.setattr(converted.interpreter, 'clock', lambda: moment)
    assert converted.run('Readback.Probe') == case['reopened']
