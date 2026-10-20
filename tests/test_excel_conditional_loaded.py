"""Native loaded group identities, interval ordering and bounded multiple transfers."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from _conditional_loaded_package import assert_witness_part

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, read_rules, with_rules
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/conditional_loaded'
WITNESSES: list[dict[str, Any]] = [case for name in ('same-column-native', 'interval-key-native', 'tail-tie-native', 'prefix-native')
                                 for case in json.loads((ROOT / (name + '.json')).read_text(encoding='utf-8'))['probes']]
COPIES: list[dict[str, Any]] = json.loads((ROOT / 'loaded_copy.json').read_text(encoding='utf-8'))['probes']


def node(entry: XmlElement, *, omit_dxf: bool = False) -> object:
    attributes = tuple(sorted((key, value) for key, value in entry.attrib.items()
                              if key != 'xmlns' and not key.startswith('xmlns:') and not (omit_dxf and key == 'dxfId')))
    return entry.tag, attributes, entry.text or '', tuple(node(child) for child in entry)


def groups(xml: str) -> list[object]:
    # Preserve the independent native group and child sequences. Only the dxf
    # identifier is an implementation allocation; API witnesses compare color.
    return [(tuple(sorted((key, value) for key, value in group.attrib.items()
                          if key != 'xmlns' and not key.startswith('xmlns:'))),
             tuple(node(rule, omit_dxf=True) for rule in group))
            for group in xml_tree(xml).findall(NS + 'conditionalFormatting')]


@pytest.mark.parametrize('case', WITNESSES, ids=lambda case: case['name'])
def test_native_loaded_group_witness_and_saved_order(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    assert app.workbook.FileFormat() == case['format']
    app.add_module(case['inspection_source'], name='SlotWitness')
    assert app.run('Probe') == case['loaded_witness']
    path = app.save(tmp_path / 'witness.xlsm')
    assert_witness_part(path)
    with ZipFile(path) as package:
        assert groups(package.read('xl/worksheets/sheet1.xml').decode('utf-8')) == groups(case['converted_xml'])
    reopened = ExcelApplication.open(path)
    assert reopened.run('ReadWitness') == case['reopened_witness']


@pytest.mark.parametrize('case', COPIES, ids=lambda case: case['name'])
def test_native_loaded_multiple_recopy_and_saved_order(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    assert app.workbook.FileFormat() == case['input_format']
    app.add_module(case['source'], name='LoadedControl')
    assert app.run('LoadedCopy') == case['result']
    assert app.run('Snapshot') == case['runtime_final']
    path = app.save(tmp_path / 'loaded-recopy.xlsm')
    with ZipFile(path) as package:
        for part, native in case['xml52'].items():
            if part.startswith('xl/worksheets/sheet'):
                assert groups(package.read(part).decode('utf-8')) == groups(native)
    # The loaded51 cross-sheet control has a distinct, independently measured
    # post-reopen result; it must not be replaced with the runtime expectation.
    assert ExcelApplication.open(path).run('Snapshot') == case['reopened52']


@pytest.mark.parametrize('format_number', [51, 52])
def test_shared_population_multiple_copy_remains_atomic(format_number: int) -> None:
    case = next(case for case in WITNESSES if case['name'] == 'same-complete-statistical-population-' + str(format_number))
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    app.add_module('Sub CopyShared()\nWorksheets("Data").Range("A2:A5").Copy Destination:=Worksheets("Data").Range("H2:H5")\nEnd Sub', name='Atomic')
    sheet = app.workbook.sheets_[0]
    assert len(sheet.conditional_formats) == 2
    assert sheet.conditional_formats[0].color_slot == sheet.conditional_formats[1].color_slot

    def state() -> object:
        return ([(key, cell.value, cell.formula, cell.style) for key, cell in sheet.cells_.items()],
                [(id(rule), tuple(rule.areas), xml_bytes(rule.root), rule.color_slot, rule.deleted,
                  None if rule.dxf is None else xml_bytes(rule.dxf)) for rule in sheet.conditional_formats],
                tuple(sheet.conditional_color_slots.items()), sheet.conditional_color_next,
                app.workbook.stylesheet.dxf_count)

    before = state()
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyShared')
    assert state() == before


@pytest.mark.parametrize('opaque', ['attribute', 'extension', 'nested-rule', 'local-alias'])
def test_equal_populations_preserve_distinct_opaque_containers(opaque: str) -> None:
    extra = {'attribute': ' flag="first"', 'extension': '', 'nested-rule': '',
             'local-alias': ' xmlns:q="urn:first"'}[opaque]
    other = {'attribute': ' flag="second"', 'extension': '', 'nested-rule': '',
             'local-alias': ' xmlns:q="urn:second"'}[opaque]
    children = '<extLst><ext uri="first"><q:data xmlns:q="urn:first">payload</q:data></ext></extLst>' if opaque == 'extension' else ''
    second_children = children.replace('first', 'second')
    nested = '<extLst><ext uri="nested"><q:data xmlns:q="urn:rule">keep</q:data></ext></extLst>' if opaque == 'nested-rule' else ''
    xml = ('<worksheet xmlns="' + NS[1:-1] + '"><sheetData/>'
           '<conditionalFormatting sqref="A2:A5 C1:C4"' + extra + '><cfRule type="top10" priority="9">' + nested + '</cfRule>' + children + '</conditionalFormatting>'
           '<conditionalFormatting sqref="A2:A5 C1:C4"' + other + '><cfRule type="top10" priority="2">' + nested + '</cfRule>' + second_children + '</conditionalFormatting></worksheet>')
    rules = read_rules(xml)
    original = [(xml_bytes(rule.root), None if rule.container is None else xml_bytes(rule.container)) for rule in rules]
    saved = with_rules(xml, rules)
    actual = xml_tree(saved).findall(NS + 'conditionalFormatting')
    if opaque == 'nested-rule':
        # Rule-local fields may share a plain population container, but each
        # original child payload remains attached to its own priority/root.
        assert sum(len(group.findall(NS + 'cfRule')) for group in actual) == 2
        assert [entry.find(NS + 'extLst') is not None for group in actual for entry in group.findall(NS + 'cfRule')] == [True, True]
    else:
        assert len(actual) == 2
        if opaque == 'attribute':
            assert [group.get('flag') for group in actual] == ['first', 'second']
        elif opaque == 'extension':
            extensions = [group.find(NS + 'extLst') for group in actual]
            assert all(extension is not None for extension in extensions)
            assert [extension[0].get('uri') for extension in extensions if extension is not None] == ['first', 'second']
        else:
            assert [group.get('xmlns:q') for group in actual] == ['urn:first', 'urn:second']
    assert [(xml_bytes(rule.root), None if rule.container is None else xml_bytes(rule.container)) for rule in rules] == original


def test_native_input_package_integrity() -> None:
    manifest = json.loads((ROOT / 'sample_provenance.json').read_text(encoding='utf-8'))
    assert len(WITNESSES) == 36 and len(COPIES) == 6
    for name, metadata in manifest.items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == metadata.get('fixture_sha256', metadata['native_sha256'])
        with ZipFile(ROOT / name) as package:
            if name.endswith('.xlsx') or metadata.get('separate_native_mode'):
                assert not any('vbaproject' in part.lower() for part in package.namelist())
            else:
                # Removing VBA changes Excel's loaded group identities even
                # with a52 content type. Preserve the exact native input; every
                # model open above explicitly disables imported VBA execution.
                assert 'xl/vbaProject.bin' in package.namelist()
