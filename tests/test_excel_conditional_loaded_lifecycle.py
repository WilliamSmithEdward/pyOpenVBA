"""Native conditional groups across no-VBA52 loads and loaded group mutations."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from _conditional_loaded_package import assert_witness_part

from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, FormatCondition
from pyopenvba.apps.excel._model import Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import to_text

ROOT = Path(__file__).parent / 'fixtures/conditional_loaded'
LIFECYCLES: list[dict[str, Any]] = [case for name in ('shared_lifecycle', 'materialization')
                                  for case in json.loads((ROOT / (name + '.json')).read_text(encoding='utf-8'))['probes']]
MODE_RECORD: dict[str, Any] = json.loads((ROOT / 'novba52.json').read_text(encoding='utf-8'))
MODES: list[dict[str, Any]] = MODE_RECORD['probes']
COPIES: list[dict[str, Any]] = json.loads((ROOT / 'novba52_copy.json').read_text(encoding='utf-8'))['probes']


def node(entry: XmlElement) -> object:
    return (entry.tag, tuple(sorted((key, value) for key, value in entry.attrib.items()
                                   if key != 'dxfId' and key != 'xmlns' and not key.startswith('xmlns:'))),
            entry.text or '', tuple(node(child) for child in entry))


def groups(xml: str) -> list[object]:
    return [node(group) for group in xml_tree(xml).findall(NS + 'conditionalFormatting')]


def saved_groups(path: Path, expected: dict[str, str]) -> None:
    with ZipFile(path) as package:
        for part, native in expected.items():
            assert groups(package.read(part).decode('utf-8')) == groups(native)


@pytest.mark.parametrize('case', LIFECYCLES, ids=lambda case: case['name'] + ('-three' if 'afterCyan' in case['result'] else '-two'))
def test_native_loaded_group_mutation_runtime_and_serialized_identity(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    app.add_module(case['source'], name='GroupMutation')
    assert app.run('Probe') == case['result']
    assert app.run('Snapshot') == case['runtime_final']
    path = app.save(tmp_path / 'mutated.xlsm')
    assert_witness_part(path)
    saved_groups(path, case['xml52'])
    assert ExcelApplication.open(path).run('Snapshot') == case['reopened52']


@pytest.mark.parametrize('case', MODES, ids=lambda case: case['input_file'])
def test_native_no_vba52_loading_projection_and_persistence(case: dict[str, Any], tmp_path: Path) -> None:
    # These are intentionally separate native controls, not replacements for
    # the macro-bearing52 originals: VBA removal changes the loaded API mode.
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    assert app.workbook.FileFormat() == 52
    assert case['preprobe_safe_properties'] == '52|False'
    source = case['source'].split('\nFunction Capture()', 1)[0]
    app.add_module(source + '\nSub ActivateData()\nWorksheets("Data").Activate\nEnd Sub', name='ModeRead')
    app.run('ActivateData')
    witness = 'Function Probe()' in source
    expected = case.get('explicit_data_snapshot', case['loaded_witness'])
    assert app.run('Probe' if witness else 'Snapshot') == expected
    path = app.save(tmp_path / 'mode.xlsm')
    if witness:
        assert_witness_part(path)
    saved_groups(path, case['xml52'])
    reopened = ExcelApplication.open(path)
    reopened.run('ActivateData')
    # The controller observed the original target with HasVBProject=False.
    # Adding this inspection module changes the saved package, so its formerly
    # plain load projection must already have been materialized into the XML.
    assert reopened.run('ReadWitness' if witness else 'Snapshot') == (expected if 'explicit_data_snapshot' in case else case['reopened_witness'])


@pytest.mark.parametrize('case', COPIES, ids=lambda case: case['name'])
def test_native_no_vba52_generic_union_copy_maps_projected_population(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    source = case['inspection_source'] + '\nSub CopyNow()\nWorksheets("Data").Activate\nWorksheets("Data").Range("' + case['source_range'] + '").Copy Destination:=Worksheets("Data").Range("' + case['destination_range'] + '")\nApplication.CutCopyMode=False\nEnd Sub'
    app.add_module(source, name='UnionCopy')
    assert app.run('Snapshot') == case['before']
    app.run('CopyNow')
    assert app.run('Snapshot') == case['runtime_final']
    path = app.save(tmp_path / 'copy.xlsm')
    saved_groups(path, case['xml52'])
    assert ExcelApplication.open(path).run('Snapshot') == case['reopened52']


def test_derived_no_vba52_inputs_preserve_native_workbook_and_sheet_bytes() -> None:
    for case in MODES:
        with ZipFile(ROOT / case['input_file']) as package:
            assert not any('vba' in part.lower() for part in package.namelist())
            for part, native_sha in case['unchanged_parts'].items():
                assert hashlib.sha256(package.read(part)).hexdigest() == native_sha
            assert 'xl/workbook.xml' in case['unchanged_parts']
            assert 'xl/styles.xml' in case['unchanged_parts']
            assert any(part.startswith('xl/worksheets/sheet') for part in case['unchanged_parts'])


def test_native_initial_activation_context_is_retained_as_separate_evidence() -> None:
    # Both no-VBA formats initially report active Data but need an explicit
    # activation before these specialized foreign getters initialize. Preserve
    # the raw observation instead of substituting the functional test oracle.
    contexts = MODE_RECORD['activation_context_evidence']['probes']
    assert len(contexts) == 2
    for case in contexts:
        assert case['initial_properties'].endswith('|False|active=Data')
        assert case['initial_snapshot'] != case['data_snapshot']
        assert all(item[1:] == ['ERR:440', 'ERR:440'] for item in case['initial_unique_signature'])
        assert all(item[1:] == ['1(3)', '0(5)'] for item in case['data_unique_signature'])
    assert contexts[0]['data_snapshot'] == contexts[1]['data_snapshot']


PLAIN_MODIFY: list[dict[str, Any]] = json.loads((ROOT / 'plain_modify.json').read_text(encoding='utf-8'))['probes']
FRESH_MODIFY: dict[str, Any] = json.loads((ROOT / 'fresh_modify.json').read_text(encoding='utf-8'))['probes'][0]


def first_condition(app: ExcelApplication) -> FormatCondition:
    sheet = app.workbook.sheets_[0]
    return FormatCondition(sheet, min(sheet.conditional_formats, key=lambda rule: int(rule.root.get('priority', '0'))))


def describe(condition: FormatCondition) -> str:
    target = condition.AppliesTo()
    assert isinstance(target, Range)
    return 'range=' + str(target.Address()) + ';formula=' + str(condition.Formula1()) + ';priority=' + to_text(condition.Priority())


@pytest.mark.parametrize('case', PLAIN_MODIFY, ids=lambda case: case['name'])
def test_native_loaded_plain_modify_uses_fresh_area_order_until_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    # Direct Python API calls keep this input free of VBA, matching the native
    # controller. Native held proxies rebind to the sibling; use the separately
    # recorded fresh collection getter instead of claiming proxy conformance.
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    assert describe(first_condition(app)) == case['before']
    sheet = app.workbook.sheets_[0]
    first_condition(app).ModifyAppliesToRange(Range(sheet, parse_reference('M8:M9,K3:K5')))
    assert describe(first_condition(app)) == case['after_fresh']
    assert first_condition(app).rule.imported is False
    assert first_condition(app).rule.plain_load is True
    path = app.save(tmp_path / 'plain-modified.xlsm')
    with ZipFile(path) as package:
        assert 'xl/vbaProject.bin' not in package.namelist()
    assert describe(first_condition(app)) == case['after_save_fresh']
    saved_groups(path, case['xml52'])
    assert describe(first_condition(ExcelApplication.open(path, with_vba=False))) == case['reopened52']


def test_native_fresh_modify_rebases_formula_to_the_new_bounding_anchor(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FRESH_MODIFY['source'], name='FreshModify')
    assert app.run('Probe') == FRESH_MODIFY['result']
    assert app.run('Snapshot') == FRESH_MODIFY['runtime_final']
    path = app.save(tmp_path / 'fresh-modified.xlsm')
    saved_groups(path, FRESH_MODIFY['xml52'])
    assert ExcelApplication.open(path).run('Snapshot') == FRESH_MODIFY['reopened52']


def test_modify_rebase_failure_keeps_loaded_shared_group_atomic() -> None:
    case = PLAIN_MODIFY[0]
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    condition = first_condition(app)
    sheet = app.workbook.sheets_[0]
    condition.rule.root.set('type', 'cellIs')
    condition.rule.root.set('operator', 'between')
    first = condition.rule.root.find(NS + 'formula')
    assert first is not None
    first.text = 'A1+$A$1+A$1+$A1'
    second = XmlElement(NS + 'formula')
    second.text = 'A1+"unterminated'
    condition.rule.root.append(second)

    def state() -> object:
        return ([(id(rule), tuple(rule.areas), xml_bytes(rule.root), rule.imported, rule.plain_load,
                  rule.child_order, rule.color_slot, id(rule.container),
                  None if rule.dxf is None else xml_bytes(rule.dxf)) for rule in sheet.conditional_formats],
                tuple(sheet.conditional_color_slots.items()), sheet.conditional_color_next,
                sheet.conditional_changed, sheet.dirty, app.workbook.stylesheet.dxf_count)

    before = state()
    with pytest.raises(VBAUnsupportedError, match='Rebasing'):
        condition.ModifyAppliesToRange(Range(sheet, parse_reference('M8:M9,K3:K5')))
    assert state() == before
