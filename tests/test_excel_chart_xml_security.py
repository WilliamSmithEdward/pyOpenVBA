"""The chart tree builder never processes DTDs or declared entities."""
import pytest

from pyopenvba.apps.excel._charts_file import xml_parser, xml_tree


@pytest.mark.parametrize('payload', [
    '<!DOCTYPE r [<!ENTITY x "expanded">]><r>&x;</r>',
    '<!DOCTYPE r SYSTEM "file:///not-read.dtd"><r/>',
    '<!DOCTYPE r [<!ENTITY x SYSTEM "https://example.invalid/not-fetched">]><r>&x;</r>',
    '<!DOCTYPE r [<!ENTITY % p SYSTEM "file:///not-read.dtd">%p;]><r/>',
])
@pytest.mark.parametrize('encoding', ['utf-8', 'utf-16'])
@pytest.mark.parametrize('reader', ['tree', 'offset'])
def test_chart_parser_refuses_dtd_before_processing_entities(payload: str, encoding: str, reader: str) -> None:
    declaration = f'<?xml version="1.0" encoding="{encoding}"?>'
    with pytest.raises(ValueError, match='DTD'):
        data = (declaration + payload).encode(encoding)
        if reader == 'tree':
            xml_tree(data)
        else:
            xml_parser().Parse(data, True)

def test_chart_parser_keeps_namespaces_and_predefined_entities() -> None:
    root = xml_tree('<c:r xmlns:c="urn:chart"><c:t>A &amp; B</c:t></c:r>')
    assert root.tag == '{urn:chart}r'
    assert root.findtext('{urn:chart}t') == 'A & B'
