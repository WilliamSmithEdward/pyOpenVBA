"""Native worksheet-copy identities, activated proxies and loaded group children."""
from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Iterator
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, FormatCondition, Rule
from pyopenvba.apps.excel._conditional_border_context import resolve
from pyopenvba.apps.excel._model import Range
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/conditional_sheet_copy'
FRESH: list[dict[str, Any]] = json.loads((ROOT / 'fresh.json').read_text(encoding='utf-8'))['probes']
LOADED: list[dict[str, Any]] = json.loads((ROOT / 'loaded.json').read_text(encoding='utf-8'))['probes']
INITIAL_GAPS = {'deleted-leading-snapshot-first', 'deleted-middle-genuine-snapshot-first'}
PHASE = re.compile(r'(beforeH-copyActive|beforeH-originalActive|afterH-copyActive|afterH-originalActive|afterCopyEnumeration-copyActive|afterOriginalEnumeration-copyActive|afterSnapshot-copyActive|afterSnapshot-originalActive)\[activeBefore=([^|]+)\|(.*?)\|activeAfter=([^\]]+)\]')
LOADED_PHASE = re.compile(r'(beforeH-copyActive|beforeH-originalActive|afterH-copyActive|afterH-originalActive)=\[(.*?)\](?=;|$)')
PHASE_CASES = [(case, match.group(1)) for case in FRESH for match in PHASE.finditer(case['result'])]
LOADED_PHASE_CASES = [(case, match.group(1)) for case in LOADED for match in LOADED_PHASE.finditer(case['result'])]


@pytest.fixture(scope='module')
def fresh_runs() -> dict[str, tuple[ExcelApplication, str]]:
    runs: dict[str, tuple[ExcelApplication, str]] = {}
    for case in FRESH:
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(case['source'], name='ProxyDiagnostics')
        actual = app.run('ProxyProbe')
        assert isinstance(actual, str)
        runs[case['name']] = app, actual
    return runs


@pytest.fixture(scope='module')
def loaded_runs() -> dict[str, tuple[ExcelApplication, str]]:
    runs: dict[str, tuple[ExcelApplication, str]] = {}
    for case in LOADED:
        app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
        app.add_module(case['source'], name='LoadedSheetCopy')
        actual = app.run('Probe')
        assert isinstance(actual, str)
        runs[case['name']] = app, actual
    return runs


@pytest.mark.parametrize('case,phase', PHASE_CASES, ids=[case['name'] + '-' + phase for case, phase in PHASE_CASES])
def test_native_fresh_copy_activated_proxy_phase(case: dict[str, Any], phase: str,
                                               fresh_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    actual = {match.group(1): match.group(0) for match in PHASE.finditer(fresh_runs[case['name']][1])}
    expected = {match.group(1): match.group(0) for match in PHASE.finditer(case['result'])}
    assert len(actual) == len(expected) == 8
    assert actual[phase] == expected[phase]


@pytest.mark.parametrize('case', [case for case in FRESH if case['name'] not in INITIAL_GAPS], ids=lambda case: case['name'])
def test_native_fresh_copy_complete_controls(case: dict[str, Any], fresh_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    assert fresh_runs[case['name']][1] == case['result']


def test_initial_copy_context_evidence_is_preserved() -> None:
    assert len(FRESH) == 8 and len(PHASE_CASES) == 64
    partial = [case for case in FRESH if case['name'] in INITIAL_GAPS]
    assert len(partial) == 2
    for case in partial:
        assert '|initialSnapshotActiveBefore=Copied|initialSnapshot{' in case['result']
        assert '|initialSnapshotActiveAfter=Copied|' in case['result']
    # Native raw snapshots are retained; explicit activation is a measured
    # boundary. Do not replace the two initial snapshots with model output.
    assert len(LOADED_PHASE_CASES) == 16


@pytest.mark.parametrize('case,phase', LOADED_PHASE_CASES, ids=[case['name'] + '-' + phase for case, phase in LOADED_PHASE_CASES])
def test_native_loaded_copy_activated_group_children(case: dict[str, Any], phase: str,
                                                   loaded_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    expected = {match.group(1): match.group(0) for match in LOADED_PHASE.finditer(case['result'])}
    actual = {match.group(1): match.group(0) for match in LOADED_PHASE.finditer(loaded_runs[case['name']][1])}
    assert len(actual) == len(expected) == 4
    assert actual[phase] == expected[phase]


@pytest.mark.parametrize('case', [case for case in LOADED if case['name'] != 'shared-52'], ids=lambda case: case['name'])
def test_native_loaded_copy_complete_controls(case: dict[str, Any], loaded_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    assert loaded_runs[case['name']][1] == case['result']


def group_sequence(xml: str) -> object:
    return [(group.get('sqref'), [(rule.get('type'), rule.get('priority'), rule.get('rank'),
                                 tuple(child.text or '' for child in rule.findall(NS + 'formula')))
                                for rule in group.findall(NS + 'cfRule')])
            for group in xml_tree(xml).findall(NS + 'conditionalFormatting')]


@pytest.mark.parametrize('case', LOADED, ids=lambda case: case['name'])
def test_native_loaded_copy_save_and_explicit_reopen_phases(case: dict[str, Any], tmp_path: Path,
                                                         loaded_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    app = loaded_runs[case['name']][0]
    assert app.run('ReadOriginal') == case['runtime_final']
    path = app.save(tmp_path / 'sheet-copied.xlsm')
    with ZipFile(path) as package:
        for part, native in case['xml52'].items():
            if part.startswith('xl/worksheets/sheet'):
                assert group_sequence(package.read(part).decode('utf-8')) == group_sequence(native)
    reopened = ExcelApplication.open(path)
    assert reopened.run('ReadCopied') == case['reopened_copied']
    assert reopened.run('ReadOriginal') == case['reopened_original']


def test_shared52_initial_context_remains_separate_raw_evidence() -> None:
    case = next(case for case in LOADED if case['name'] == 'shared-52')
    initial = case['result'].split(';initial=[', 1)[1].split('];beforeH-copyActive=[', 1)[0]
    activated = LOADED_PHASE.search(case['result'])
    assert activated is not None
    assert initial != activated.group(2)


@pytest.mark.parametrize('case', FRESH, ids=lambda case: case['name'])
def test_copied_dense_registry_does_not_reassign_source_tombstones(case: dict[str, Any],
                                                               fresh_runs: dict[str, tuple[ExcelApplication, str]]) -> None:
    app = fresh_runs[case['name']][0]
    original, copied = app.workbook.sheets_[:2]
    assert copied.conditional_color_next == 3
    assert sorted(rule.color_slot for rule in copied.conditional_formats if rule.color_slot is not None) == [1, 2, 3]
    for rule in copied.conditional_formats:
        assert rule.color_slot is not None
        assert copied.conditional_color_slots[rule.color_slot] is rule
    assert len({id(rule) for rule in original.conditional_formats + copied.conditional_formats}) == 5
    assert all(rule not in copied.conditional_color_slots.values() for rule in original.conditional_formats)
    if case['name'].startswith('deleted-'):
        assert original.conditional_color_next == 3
        assert any(rule.deleted for rule in original.conditional_color_slots.values())


class CountedRules(list[Rule]):
    iterations: int = 0

    def __iter__(self) -> Iterator[Rule]:
        self.iterations += 1
        return super().__iter__()


def test_shared_foreign_reads_index_each_sheet_once_and_invalidate_after_mutation() -> None:
    case = next(case for case in LOADED if case['name'] == 'shared-52')
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    app.add_module(case['source'], name='CounterSetup')
    app.run('Probe')
    source, copied = app.workbook.sheets_[:2]
    copied.Activate()
    rule = next(rule for rule in source.conditional_formats if rule.root.get('priority') == '2')
    source_rules, target_rules = CountedRules(source.conditional_formats), CountedRules(copied.conditional_formats)
    source.conditional_formats, copied.conditional_formats = source_rules, target_rules
    source.conditional_color_groups = copied.conditional_color_groups = None
    first = xml_bytes(resolve(source, rule))
    for _ in range(100):
        assert xml_bytes(resolve(source, rule)) == first
    assert source_rules.iterations == target_rules.iterations == 1
    FormatCondition(source, rule).vba_set('Priority', 1)
    assert source.conditional_color_groups is None
    before = source_rules.iterations, target_rules.iterations
    for _ in range(100):
        assert xml_bytes(resolve(source, rule)) == first
    assert source_rules.iterations == before[0] + 1
    assert target_rules.iterations == before[1]


def test_failed_rebase_preserves_warm_group_index_and_all_rule_state() -> None:
    app = ExcelApplication.open(ROOT / '../conditional_loaded/generic-ordered-tail-permutation-51.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    rule = min(sheet.conditional_formats, key=lambda item: int(item.root.get('priority', '0')))
    formula = rule.root.find(NS + 'formula')
    assert formula is not None
    formula.text = 'A1+"unterminated'
    sheet.conditional_color_groups = {1: (list(sheet.conditional_formats), {item: index for index, item in enumerate(sheet.conditional_formats)})}
    index = sheet.conditional_color_groups

    def state() -> object:
        return ([(id(item), tuple(item.areas), xml_bytes(item.root), item.color_slot, item.child_order,
                  None if item.dxf is None else xml_bytes(item.dxf)) for item in sheet.conditional_formats],
                tuple(sheet.conditional_color_slots.items()), sheet.conditional_color_next,
                sheet.conditional_changed, sheet.dirty)

    before = state()
    with pytest.raises(VBAUnsupportedError, match='Rebasing'):
        FormatCondition(sheet, rule).ModifyAppliesToRange(Range(sheet, parse_reference('M8:M9,K3:K5')))
    assert state() == before and sheet.conditional_color_groups is index


def test_preserved_native_records_and_input_package_hashes() -> None:
    manifest = json.loads((ROOT / 'provenance.json').read_text(encoding='utf-8'))
    assert hashlib.sha256((ROOT / 'fresh.json').read_bytes()).hexdigest() == manifest['fresh_native_record_sha256']
    for name, expected in manifest['loaded_input_packages'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected

UNIQUE: dict[str, Any] = json.loads((ROOT / 'unique.json').read_text(encoding='utf-8'))['probe']


def valid_unique_frames(snapshot: str) -> dict[str, str]:
    # The native diagnostic also queried out-of-range collection slots. Keep
    # those raw results in the fixture; this contract covers actual rules.
    keys = ('heldCopySecond', 'Data1', 'Data2', 'Copied1', 'Copied2', 'Copied3')
    result: dict[str, str] = {}
    for key in keys:
        match = re.search(re.escape(key) + r'\{([^}]*)\}', snapshot)
        assert match is not None
        result[key] = match.group(1)
    return result


def test_native_unique_valid_group_children_and_saved_contexts(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / UNIQUE['input_file'], with_vba=False)
    app.add_module(UNIQUE['source'], name='UniqueChildControl')
    actual = app.run('Cx5Probe', app.workbook)
    assert isinstance(actual, str)
    for phase in ('before-copy', 'before-original', 'after-copy', 'after-original'):
        native_match = re.search(re.escape(phase) + r'=\[(.*?)\](?=;|$)', UNIQUE['result'])
        model_match = re.search(re.escape(phase) + r'=\[(.*?)\](?=;|$)', actual)
        assert native_match is not None and model_match is not None
        assert model_match.group(1).split(';heldCopySecond', 1)[0] == native_match.group(1).split(';heldCopySecond', 1)[0]
        assert valid_unique_frames(model_match.group(1)) == valid_unique_frames(native_match.group(1))
    path = app.save(tmp_path / UNIQUE['file52'])
    with ZipFile(path) as package:
        for part, native in UNIQUE['xml52'].items():
            if part.startswith('xl/worksheets/sheet'):
                assert group_sequence(package.read(part).decode('utf-8')) == group_sequence(native)
    reopened = ExcelApplication.open(path)
    copied = reopened.workbook.sheets_[1]
    from pyopenvba.apps.excel._conditional import FormatConditions

    held = FormatConditions(Range(copied, parse_reference('A1:D22'))).Item(1)
    for active, key in (('Copied', 'reopened_copied'), ('Data', 'reopened_original')):
        model = reopened.run('Cx5Read', reopened.workbook, active, held)
        assert isinstance(model, str)
        assert valid_unique_frames(model) == valid_unique_frames(UNIQUE[key])


def test_unsupported_missing_foreign_child_read_preserves_workbook_state() -> None:
    case = next(case for case in LOADED if case['name'] == 'shared-52')
    app = ExcelApplication.open(ROOT / case['input_file'], with_vba=False)
    app.add_module(case['source'], name='MissingMemberSetup')
    app.run('Probe')
    source, copied = app.workbook.sheets_[:2]
    removed = next(rule for rule in source.conditional_formats if rule.root.get('priority') == '9')
    FormatCondition(source, removed).Delete()
    source.Activate()
    rule = next(rule for rule in copied.conditional_formats if rule.root.get('priority') == '2')

    def state() -> object:
        return (app.workbook.saved, app.workbook.stylesheet.xml, app.workbook.stylesheet.dxf_count,
                tuple(xml_bytes(app.workbook.stylesheet.dxf(index)) for index in range(app.workbook.stylesheet.dxf_count)),
                [(sheet.dirty, sheet.conditional_changed, sheet.conditional_color_next,
                  tuple(sheet.conditional_color_slots.items()),
                  [(id(item), item.deleted, tuple(item.areas), xml_bytes(item.root), item.child_order, item.color_slot,
                    None if item.dxf is None else xml_bytes(item.dxf)) for item in sheet.conditional_formats])
                 for sheet in app.workbook.sheets_])

    before = state()
    with pytest.raises(VBAUnsupportedError, match='missing group member'):
        resolve(copied, rule)
    assert state() == before


def test_unique_native_input_and_source_inventory_are_pinned() -> None:
    from pyopenvba import ExcelFile

    manifest = json.loads((ROOT / 'provenance.json').read_text(encoding='utf-8'))
    item = manifest['unique_input_package']
    path = ROOT / item['file']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == item['sha256']
    inventory = json.loads((ROOT / item['module_inventory']).read_text(encoding='utf-8'))
    with ExcelFile(path) as workbook:
        assert workbook.vba_modules() == inventory['modules']
    assert inventory['external_relationships'] == []
    for name, expected in manifest['additional_raw_evidence'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected


ADMISSION: list[dict[str, Any]] = json.loads((ROOT / 'fresh_admission.json').read_text(encoding='utf-8'))['controls']


@pytest.mark.parametrize('case', ADMISSION, ids=lambda case: case['name'])
def test_native_fresh_copy_safe_identity_controls(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='SafeIdentityControl')
    assert app.run('Probe') == case['result']
