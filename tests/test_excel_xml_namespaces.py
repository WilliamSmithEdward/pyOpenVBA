"""Namespace-sensitive Office attribute values survive XML tree edits."""
from __future__ import annotations

from copy import deepcopy
import xml.etree.ElementTree as ET

from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree

MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'


def bindings(entry: XmlElement, inherited: dict[str, str] | None = None) -> dict[str, str]:
    result = dict(inherited or {})
    result.update({key[6:]: value for key, value in entry.attrib.items() if key.startswith('xmlns:')})
    if 'xmlns' in entry.attrib:
        result[''] = entry.attrib['xmlns']
    return result


def test_compatibility_attributes_keep_ancestor_aliases_and_nested_rebindings() -> None:
    source = f'''<root xmlns:mc="{MC}" xmlns:x="urn:first" xmlns:alias="urn:first"
      mc:Ignorable="x alias" mc:MustUnderstand="alias" mc:PreserveElements="x:keep alias:*"
      mc:PreserveAttributes="alias:flag" mc:ProcessContent="x:content">
      <mc:AlternateContent><mc:Choice Requires="x alias"><x:kept/></mc:Choice></mc:AlternateContent>
      <child xmlns:x="urn:second" mc:Ignorable="x" custom="x:arbitrary">
        <mc:Choice Requires="x alias"><x:second alias:flag="yes"/></mc:Choice>
      </child>
      <x:first/>
    </root>'''
    original = xml_tree(source)
    output = xml_tree(xml_bytes(deepcopy(original)))
    scope = bindings(output)
    assert scope['x'] == scope['alias'] == 'urn:first'
    for attribute in ('Ignorable', 'MustUnderstand', 'PreserveElements', 'PreserveAttributes', 'ProcessContent'):
        assert output.get('{' + MC + '}' + attribute) == original.get('{' + MC + '}' + attribute)
    choice = output[0][0]
    assert choice.get('Requires') == 'x alias'
    assert choice[0].tag == '{urn:first}kept'
    inner_scope = bindings(output[1], scope)
    assert inner_scope['x'] == 'urn:second'
    assert inner_scope['alias'] == 'urn:first'
    assert output[1].get('custom') == 'x:arbitrary'
    assert output[1][0].get('Requires') == 'x alias'
    assert output[1][0][0].tag == '{urn:second}second'
    assert output[1][0][0].get('{urn:first}flag') == 'yes'
    assert output[2].tag == '{urn:first}first'


def test_generated_namespace_cannot_capture_alias_used_only_in_values() -> None:
    root = xml_tree(f'<root xmlns:mc="{MC}" xmlns:ns0="urn:value-only" mc:Ignorable="ns0">'
                    '<child xmlns:ns1="urn:nested-value-only" custom="ns1:item"/></root>')
    root.append(XmlElement('{urn:new}added', {'{urn:attribute}flag': 'yes'}))
    output = xml_tree(xml_bytes(root))
    assert bindings(output)['ns0'] == 'urn:value-only'
    assert bindings(output[0], bindings(output))['ns1'] == 'urn:nested-value-only'
    assert output[0].get('custom') == 'ns1:item'
    assert output[1].tag == '{urn:new}added'
    assert output[1].get('{urn:attribute}flag') == 'yes'


def test_default_namespace_and_qualified_attributes_stay_distinct() -> None:
    root = xml_tree('<root xmlns="urn:default"><child xmlns=""><plain/></child>'
                    '<other xmlns="urn:other"/></root>')
    root.append(XmlElement('added'))
    root.set('{urn:default}flag', 'yes')
    output = xml_tree(xml_bytes(root))
    assert output.tag == '{urn:default}root'
    assert output.get('{urn:default}flag') == 'yes'
    assert output[0].tag == 'child'
    assert output[0][0].tag == 'plain'
    assert output[1].tag == '{urn:other}other'
    assert output[2].tag == 'added'


def test_serializer_does_not_modify_tree_or_global_namespace_choices() -> None:
    root = xml_tree('<p:root xmlns:p="urn:original" xmlns:q="urn:value" custom="q:item" xml:space="preserve"/>')
    before = deepcopy(root)
    control = XmlElement('{urn:original}control')
    default_output = ET.tostring(control)
    xml_bytes(root)
    assert ET.tostring(control) == default_output
    assert root.tag == before.tag and root.attrib == before.attrib
    output = xml_tree(xml_bytes(root))
    assert output.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'
    assert bindings(output)['q'] == 'urn:value'
