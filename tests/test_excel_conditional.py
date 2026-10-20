"""Independent native conditional-formatting controls and package preservation."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule, read_rules, with_rules

FIXTURE = Path(__file__).parent / 'fixtures/conditional/initial.json'
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding='utf-8'))
SAVED_SOURCE = '''Function SavedRules() As String
Dim f As Object,key As Variant
SavedRules=CStr(Cells.FormatConditions.Count)
For Each f In Cells.FormatConditions
For Each key In Array("Type","Operator","Formula1","Formula2","Priority","StopIfTrue")
SavedRules=SavedRules & "|" & SavedProp(f,CStr(key))
Next
SavedRules=SavedRules & "|" & f.AppliesTo.Address
Next
End Function
Function SavedProp(f As Object,key As String) As String
On Error GoTo Failed
SavedProp=CStr(CallByName(f,key,VbGet))
Exit Function
Failed:
SavedProp="ERR:" & CStr(Err.Number)
End Function'''


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_conditional_controls(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_conditional_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module(SAVED_SOURCE, name='Saved')
    expected = case.get('saved_result52', case.get('saved_result', app.run('SavedRules')))
    reopened = ExcelApplication.open(app.save(tmp_path / 'rules.xlsm'))
    assert reopened.run('SavedRules') == expected


def test_file_edit_preserves_other_bytes_and_namespace_only_values() -> None:
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="{NS[1:-1]}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:q="urn:value" xmlns:e="urn:extension" mc:Ignorable="q e">
<!-- untouched --><sheetData strange="x > y"/><mergeCells count="0"/>
<conditionalFormatting sqref="A1:A5" e:setting="q:token"><cfRule type="expression" priority="1"
 mc:PreserveElements="q:keep"><formula>A1&gt;0</formula><extLst><ext uri="z"><e:opaque/></ext></extLst></cfRule></conditionalFormatting>
<dataValidations count="0"/><?preserve instruction?><extLst><ext uri="q"><e:opaque a="q:token"/></ext></extLst></worksheet>'''
    rules = read_rules(xml)
    rules[0].areas = [Area(2, 2, 7, 2)]
    output = with_rules(xml, rules)
    assert '<!-- untouched --><sheetData strange="x > y"/><mergeCells count="0"/>' in output
    assert '<dataValidations count="0"/><?preserve instruction?><extLst><ext uri="q"><e:opaque a="q:token"/></ext></extLst>' in output
    root = xml_tree(output)
    group = root.find(NS + 'conditionalFormatting')
    assert group is not None and group.get('sqref') == 'B2:B7'
    assert group.get('{urn:extension}setting') == 'q:token'
    assert group.get('xmlns:q') == 'urn:value'
    rule = group.find(NS + 'cfRule')
    assert rule is not None and rule.get('{http://schemas.openxmlformats.org/markup-compatibility/2006}PreserveElements') == 'q:keep'
    assert rule.find('.//{urn:extension}opaque') is not None


def test_file_changes_only_standard_rules_and_places_before_validation() -> None:
    xml = f'<worksheet xmlns="{NS[1:-1]}" xmlns:e="urn:extension"><sheetData/>' \
          '<mergeCells count="0"/><dataValidations count="0"/>' \
          '<extLst><ext uri="q"><e:conditionalFormatting value="untouched"/></ext></extLst></worksheet>'
    root = XmlElement(NS + 'cfRule', {'type': 'expression', 'priority': '1'})
    formula = XmlElement(NS + 'formula')
    formula.text = 'A1>0'
    root.append(formula)
    output = with_rules(xml, [Rule([Area(1, 1, 2, 1)], root)])
    parsed = xml_tree(output)
    assert [entry.tag.rsplit('}', 1)[-1] for entry in parsed] == [
        'sheetData', 'mergeCells', 'conditionalFormatting', 'dataValidations', 'extLst']
    assert '<e:conditionalFormatting value="untouched"/>' in output
    assert len(read_rules(output)) == 1
    stripped = with_rules(output, [])
    assert read_rules(stripped) == []
    assert '<e:conditionalFormatting value="untouched"/>' in stripped


def test_file_edit_groups_original_rules_without_losing_container_metadata() -> None:
    xml = f'<worksheet xmlns="{NS[1:-1]}"><sheetData/><conditionalFormatting pivot="1" sqref="A1">' \
          '<cfRule type="expression" priority="1"><formula>1</formula></cfRule>' \
          '<cfRule type="expression" priority="2"><formula>2</formula></cfRule>' \
          '</conditionalFormatting></worksheet>'
    rules = read_rules(xml)
    rules[0].root.set('stopIfTrue', '1')
    output = xml_tree(with_rules(xml, rules))
    groups = output.findall(NS + 'conditionalFormatting')
    assert len(groups) == 1 and groups[0].get('pivot') == '1'
    assert len(groups[0].findall(NS + 'cfRule')) == 2


def test_empty_self_closing_worksheet_gets_rules_inside_its_root() -> None:
    xml = f'<?xml version="1.0"?><s:worksheet xmlns:s="{NS[1:-1]}"/><!--after-->'
    rule = Rule([Area(1, 1, 1, 1)], XmlElement(NS + 'cfRule', {'type': 'expression', 'priority': '1'}))
    output = with_rules(xml, [rule])
    assert len(read_rules(output)) == 1
    assert output.endswith('</s:worksheet><!--after-->')
    assert with_rules(xml, []) == xml


@pytest.mark.parametrize('xml', ['<!DOCTYPE worksheet><worksheet/>',
                                '<!DOCTYPE worksheet [<!ENTITY x "boom">]><worksheet>&x;</worksheet>',
                                '<!DOCTYPE worksheet SYSTEM "file:///missing"><worksheet/>'])
def test_rule_reader_and_writer_refuse_dtds(xml: str) -> None:
    with pytest.raises(ValueError, match='DTD'):
        read_rules(xml)
    with pytest.raises(ValueError, match='DTD'):
        with_rules(xml, [])


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_saved_rules_read_by_native_excel(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module(SAVED_SOURCE, name='Saved')
    expected = case.get('saved_result52', case.get('saved_result', app.run('SavedRules')))
    path = app.save(tmp_path / 'rules.xlsm')
    native_excel.open_document(path)
    assert native_excel.run_raw('Saved.SavedRules') == expected
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel
