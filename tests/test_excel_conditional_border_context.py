"""Native fresh-rule color identities across activation, reordering and deletion."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS

ROOT = Path(__file__).parent / 'fixtures/conditional_border_context'
NATIVE: dict[str, Any] = json.loads((ROOT / 'native.json').read_text(encoding='utf-8'))
SETTER_CASES: list[dict[str, Any]] = json.loads((ROOT / 'setter_style.json').read_text(encoding='utf-8'))['probes']
LOADED: dict[str, Any] = json.loads((ROOT / 'loaded_slots.json').read_text(encoding='utf-8'))


def activate(app: ExcelApplication, sheet: str) -> None:
    app.add_module('Sub ActivateSheet()\nWorksheets("' + sheet + '").Activate\nEnd Sub', name='Activation')
    app.run('ActivateSheet')


@pytest.mark.parametrize('case', NATIVE['fresh'], ids=lambda case: case['name'])
def test_fresh_native_color_context(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Harness')
    result = app.run('Setup')
    if 'setup_result' in case:
        assert result == case['setup_result']
    roots = [(rule, xml_bytes(rule.root), xml_bytes(XmlElement(NS + 'dxf') if rule.dxf is None else rule.dxf))
             for sheet in app.workbooks()[0].sheets_ for rule in sheet.conditional_formats]
    for phase in case['phases']:
        activate(app, 'Landing' if phase['phase'] == 'landing' else 'Data')
        assert app.run('Probe') == phase['result']
    for rule, root, dxf in roots:
        assert xml_bytes(rule.root) == root
        assert xml_bytes(XmlElement(NS + 'dxf') if rule.dxf is None else rule.dxf) == dxf


@pytest.mark.parametrize('case', NATIVE['cross_sheet'], ids=lambda case: case['name'])
def test_native_transfer_context_and_reopened_owner_binding(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='SetupHarness')
    assert app.run('Probe') == case['result']
    app.add_module(case['inspection_source'], name='Inspection')
    for phase in case['phases']:
        if phase['phase'] == 'reopened':
            continue
        activate(app, 'Landing' if phase['phase'] == 'active-landing' else 'Data')
        assert app.run('BorderInspect') == phase['result']
    imported = ExcelApplication.open(ROOT / case['file'])
    imported.add_module(case['inspection_source'], name='Inspection')
    expected = next(phase['result'] for phase in case['phases'] if phase['phase'] == 'reopened')
    assert imported.run('BorderInspect') == expected
    saved = tmp_path / 'context.xlsx'
    assert app.workbooks()[0].save_to(saved, file_format=51)
    reopened = ExcelApplication.open(saved)
    reopened.add_module(case['inspection_source'], name='Inspection')
    assert reopened.run('BorderInspect') == expected


@pytest.mark.parametrize('assignment', ['Color=vbBlue', 'ColorIndex=5'])
def test_setting_border_color_writes_native_implicit_thin_line(assignment: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Setup()\nDim f As Object\n'
                   'Set f=Range("A1:A4").FormatConditions.Add(2,Formula1:="=TRUE")\n'
                   'f.Borders(xlBottom).' + assignment + '\nEnd Sub', name='Harness')
    app.run('Setup')
    with ZipFile(app.save(tmp_path / 'colored.xlsm')) as package:
        root = xml_tree(package.read('xl/styles.xml').decode('utf-8'))
    bottom = root.find(NS + 'dxfs/' + NS + 'dxf/' + NS + 'border/' + NS + 'bottom')
    assert bottom is not None and bottom.get('style') == 'thin'
    assert bottom.find(NS + 'color') is not None


@pytest.mark.parametrize('case', SETTER_CASES, ids=lambda case: case['name'])
def test_native_color_setter_line_fidelity(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Harness')
    app.run('Setup')
    assert app.run('Probe') == case['result']
    generated = ExcelApplication.open(app.save(tmp_path / 'setter.xlsm'))
    assert generated.run('Probe') == case['reopened']
    imported = ExcelApplication.open(ROOT / case['file'])
    imported.add_module(case['inspection_source'], name='Harness')
    assert imported.run('Probe') == case['reopened']
    imported_again = ExcelApplication.open(imported.save(tmp_path / 'imported-setter.xlsm'))
    assert imported_again.run('Probe') == case['reopened']


@pytest.mark.parametrize('case', LOADED['loaded'], ids=lambda case: case['name'])
def test_native_loaded_slots_and_new_allocation(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'])
    app.add_module(case['inspection_source'], name='Readback')
    for phase in case['phases']:
        activate(app, 'Landing' if phase['phase'] == 'landing' else 'Data')
        assert app.run('Readback.Probe') == phase['result']
    app.add_module('Sub Fresh()\nDim f As Object\nWorksheets("Data").Activate\n'
                   'Set f=Worksheets("Landing").Range("E1:E4").FormatConditions.AddTop10\n'
                   'f.Borders(xlBottom).Color=vbMagenta\nEnd Sub', name='FreshControl')
    app.run('Fresh')
    assert app.run('Readback.Probe') == case['after_fresh']


@pytest.mark.parametrize('case', LOADED['groups'], ids=lambda case: case['name'])
def test_native_column_first_loaded_group_slots_and_conversion(case: dict[str, Any], tmp_path: Path) -> None:
    for filename in (case['file'], case['converted_file']):
        app = ExcelApplication.open(ROOT / filename)
        app.add_module(case['inspection_source'], name='Readback')
        assert app.run('Readback.Probe') == case['result']
        if 'active_source' in case:
            activate(app, case['active_source']['active_sheet'])
            assert app.run('Readback.Probe') == case['active_source']['result']
    original = ExcelApplication.open(ROOT / case['file'])
    original.add_module(case['inspection_source'], name='Readback')
    converted = ExcelApplication.open(original.save(tmp_path / 'converted.xlsm'))
    assert converted.run('Readback.Probe') == case['result']


@pytest.mark.parametrize('case', LOADED['deleted'], ids=lambda case: case['name'])
def test_native_loaded_slot_reindexes_live_rules_without_tombstones(case: dict[str, Any], tmp_path: Path) -> None:
    imported = ExcelApplication.open(ROOT / case['file'])
    inspection = 'Function P(' + case['source'].split('Function P(', 1)[1]
    imported.add_module(inspection, name='Readback')
    assert imported.run('Readback.BorderInspect') == case['reopened']
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeScenario')
    app.run('Setup')
    reopened = ExcelApplication.open(app.save(tmp_path / 'deleted.xlsm'))
    assert reopened.run('NativeScenario.BorderInspect') == case['reopened']


@pytest.mark.parametrize('case', NATIVE['imported_supported'], ids=lambda case: case['name'])
def test_native_imported_rule_mutation_and_creation_context(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / 'duplicate-copy.xlsx')
    app.add_module(case['source'] + '\n' + case['inspection_source'], name='Readback')
    app.run('Setup')
    for phase in case['phases']:
        activate(app, 'Landing' if phase['phase'] == 'landing' else 'Data')
        assert app.run('Readback.Probe') == phase['result']
