"""Native text, presence and time conditional APIs and generated formulas."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule, remove
from pyopenvba.apps.excel._conditional_predicates import regenerate
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/conditional/predicates_api.json').read_text(encoding='utf-8'))
API_CASES: list[dict[str, Any]] = list(RECORD['probes'])
API_CASES.extend(json.loads((Path(__file__).parent / 'fixtures/conditional/predicates_api_edges.json').read_text(encoding='utf-8'))['probes'])
STRING_CASES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/conditional/predicates_strings.json').read_text(encoding='utf-8'))['probes']
STRING_ROUNDTRIP: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/conditional/predicates_strings_roundtrip.json').read_text(encoding='utf-8'))['probes']
STRING_BEFORE = [case for case in STRING_ROUNDTRIP if case.get('phase') != 'after-native-save-reopen']
STRING_AFTER = [case for case in STRING_ROUNDTRIP
                if case.get('phase') == 'after-native-save-reopen' and case['result'] != 'NATIVE-OPEN-ERROR']


@pytest.mark.parametrize('case', API_CASES + STRING_CASES + STRING_BEFORE, ids=lambda case: case['name'])
def test_native_predicate_api(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', STRING_AFTER, ids=lambda case: case['name'])
def test_native_saved_string_import_and_resave(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/conditional' / case['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    # Force a rules rewrite, rather than validating only untouched package copying.
    app.workbook.sheets_[0].conditional_changed = True
    app.workbook.sheets_[0].touched()
    reopened = ExcelApplication.open(app.save(tmp_path / 'native-resaved.xlsx'), with_vba=False)
    reopened.add_module(case['source'], name='Harness')
    assert reopened.run('Probe') == case['result']


@pytest.mark.parametrize('case', STRING_AFTER, ids=lambda case: case['name'])
def test_model_saved_strings_match_native_reopened_snapshot(case: dict[str, Any], tmp_path: Path) -> None:
    native_name = Path(case['file']).stem.removeprefix('predicate-string-')
    creation = next((item for item in STRING_BEFORE if item.get('file') == case['file']), None)
    if creation is None:
        creation = next(item for item in STRING_CASES if item['name'] == native_name)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(creation['source'], name='Harness')
    assert app.run('Probe') == creation['result']
    reopened = ExcelApplication.open(app.save(tmp_path / 'model-saved.xlsx'), with_vba=False)
    reopened.add_module(case['source'], name='Harness')
    assert reopened.run('Probe') == case['result']


@pytest.mark.parametrize('case', [case for case in API_CASES if 'class=' in case['result']],
                         ids=lambda case: case['name'])
def test_predicate_snapshot_after_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    expected = 'class=' + case['result'].split('class=', 1)[1].split('|bold=', 1)[0].split('|count=', 1)[0]
    inspection = RECORD['inspection_source'].replace('Inspect', 'SavedInspect').replace('Prop', 'SavedProp')
    app.add_module('''Function SavedRule() As String
SavedRule=SavedInspect(Cells.FormatConditions(1))
End Function
''' + inspection, name='Saved', workbook=app.workbook)
    assert app.run('SavedRule') == expected
    reopened = ExcelApplication.open(app.save(tmp_path / 'predicate.xlsm', workbook=app.workbook))
    assert reopened.run('SavedRule') == expected


def test_regeneration_preserves_unknown_xml_and_ordinary_conditions() -> None:
    root = xml_tree(f'<cfRule xmlns="{NS[1:-1]}" xmlns:x="urn:unknown" type="beginsWith" text="a&amp;b" custom="kept">'
                    '<formula>obsolete</formula><extLst><x:opaque value="unchanged"/></extLst></cfRule>')
    regenerate(root, Area(9, 3, 12, 3))
    assert root.findtext(NS + 'formula') == 'LEFT(C9,LEN("a&b"))="a&b"'
    assert root.get('custom') == 'kept' and root.find(NS + 'extLst/{urn:unknown}opaque') is not None
    root.set('type', 'expression')
    regenerate(root, Area(20, 7, 25, 7))
    assert root.findtext(NS + 'formula') == 'LEFT(C9,LEN("a&b"))="a&b"'


def test_range_change_regenerates_missing_imported_compatibility_formula() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    root = XmlElement(NS + 'cfRule', {'type': 'containsBlanks', 'priority': '1'})
    book.sheets_[0].conditional_formats.append(Rule([Area(1, 1, 5, 1)], root))
    remove(book.sheets_[0], [Area(1, 1, 1, 1)])
    assert root.findtext(NS + 'formula') == 'LEN(TRIM(A2))=0'


@pytest.mark.parametrize('text,after', [('_x0041_', '_x0041_'), ('a\r\nb', 'a\nb'), ('\U0001f600', '\U0001f600'),
                                     ('\ud83d', ''), ('\ude00', ''), ('a' * 253 + '\ud83d', 'a' * 253)])
def test_plain_text_attribute_has_independent_save_normalization(text: str, after: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    root = XmlElement(NS + 'cfRule', {'type': 'containsText', 'text': text, 'priority': '1'})
    regenerate(root, Area(1, 1, 5, 1))
    book.sheets_[0].conditional_formats.append(Rule([Area(1, 1, 5, 1)], root))
    book.sheets_[0].conditional_changed = True
    book.sheets_[0].touched()
    reopened = ExcelApplication.open(app.save(tmp_path / 'text.xlsm'))
    imported = reopened.workbook.sheets_[0].conditional_formats[0].root
    assert imported.get('text') == after
    assert imported.findtext(NS + 'formula') == root.findtext(NS + 'formula')
    assert root.get('text') == text


def test_nul_text_runtime_supported_but_save_refuses_native_corruption(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Probe() As String
Dim f As Object
Set f=Range("A1:A5").FormatConditions.Add(Type:=9,String:="a" & ChrW(0) & "b",TextOperator:=0)
Probe=CStr(Len(f.Text))
End Function''', name='Harness')
    assert app.run('Probe') == '3'
    rule = app.workbook.sheets_[0].conditional_formats[0]
    before = (dict(rule.root.attrib), rule.root.findtext(NS + 'formula'), list(rule.areas))
    path = tmp_path / 'nul.xlsm'
    with pytest.raises(VBAUnsupportedError, match='NUL'):
        app.save(path)
    assert not path.exists()
    assert (dict(rule.root.attrib), rule.root.findtext(NS + 'formula'), list(rule.areas)) == before
