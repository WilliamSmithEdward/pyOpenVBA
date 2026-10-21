"""Bounded direct-child XML scans retain namespaces and reject unsafe structure."""
from __future__ import annotations

import pytest

from pyopenvba._opc_xml import bounded_child

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def read(xml: str, *, max_bytes: int = 4096, max_elements: int = 65) -> object:
    return bounded_child(xml.encode(), NS, 'worksheet', 'sortState',
                         max_bytes=max_bytes, max_elements=max_elements)


def test_selected_subtree_uses_ancestor_namespaces_and_preserves_owned_text_flags() -> None:
    root = bounded_child((f'<s:worksheet xmlns:s="{NS}"><s:sortState ref="A1:B4">'
                          '<s:sortCondition ref="A2:A4">text</s:sortCondition>tail'
                          '</s:sortState></s:worksheet>').encode(), NS, 'worksheet', 'sortState',
                         max_bytes=4096, max_elements=65)
    assert root is not None
    assert root.name == NS + '}sortState' and root.get('ref') == 'A1:B4'
    assert root.has_text and root.children[0].has_text
    assert root.children[0].name == NS + '}sortCondition'


@pytest.mark.parametrize('inside', [
    '<sortState/><sortState/>',
    '<extLst><sortState/></extLst>',
    '<x:sortState xmlns:x="urn:foreign"/>',
])
def test_duplicate_wrapped_or_foreign_metadata_is_not_treated_as_absent(inside: str) -> None:
    with pytest.raises(ValueError):
        read(f'<worksheet xmlns="{NS}">{inside}</worksheet>')


@pytest.mark.parametrize('declaration', [
    '<!DOCTYPE worksheet>',
    '<!DOCTYPE worksheet [<!ENTITY value "text">]>',
    '<!DOCTYPE worksheet SYSTEM "file:///not-read">',
])
def test_selected_reader_preserves_dtd_entity_refusal(declaration: str) -> None:
    with pytest.raises(ValueError, match='DTD'):
        read(declaration + f'<worksheet xmlns="{NS}"><sortState/></worksheet>')


@pytest.mark.parametrize('budget', ['bytes', 'elements'])
def test_selected_limits_bound_only_the_retained_metadata(budget: str) -> None:
    xml = f'<worksheet xmlns="{NS}"><sortState ref="A1:B4"><sortCondition ref="A2:A4"/></sortState></worksheet>'
    with pytest.raises(ValueError, match='budget'):
        read(xml, max_bytes=8 if budget == 'bytes' else 4096, max_elements=1 if budget == 'elements' else 65)


def test_selected_reader_accepts_absence_but_rejects_invalid_document() -> None:
    assert read(f'<worksheet xmlns="{NS}"><sheetData/></worksheet>') is None
    with pytest.raises(ValueError):
        read(f'<worksheet xmlns="{NS}"><sheetData></worksheet>')


@pytest.mark.parametrize('opaque', ['<!--retained-->', '<?custom retained?>', '<![CDATA[ ]]>'])
def test_only_owned_opaque_xml_is_flagged(opaque: str) -> None:
    root = bounded_child((f'<worksheet xmlns="{NS}"><!--outside--><sortState ref="A1:B4">'
                          + opaque + '</sortState></worksheet>').encode(), NS, 'worksheet', 'sortState',
                         max_bytes=4096, max_elements=65)
    assert root is not None and root.has_opaque


@pytest.mark.parametrize('encoding', ['ISO-8859-1', 'UTF-16'])
def test_selected_reader_refuses_non_utf8_encoding_declarations(encoding: str) -> None:
    with pytest.raises(ValueError, match='UTF-8'):
        read(f'<?xml version="1.0" encoding="{encoding}"?>'
             f'<worksheet xmlns="{NS}"><sortState/></worksheet>')


@pytest.mark.parametrize('encoding', ['UTF-8', 'utf8'])
def test_selected_reader_accepts_utf8_encoding_declarations(encoding: str) -> None:
    assert read(f'<?xml version="1.0" encoding="{encoding}"?>'
                f'<worksheet xmlns="{NS}"><sortState/></worksheet>') is not None


@pytest.mark.parametrize('selected', [
    '<sortState></sortState>',
    '<sortState/>',
    '<sortState><sortCondition></sortCondition></sortState>',
    '<sortState><sortCondition/></sortState>',
])
@pytest.mark.parametrize('following', ['', '<sheetData/>'])
def test_selected_byte_limit_includes_complete_end_tags(selected: str, following: str) -> None:
    xml = f'<worksheet xmlns="{NS}">{selected}{following}</worksheet>'
    length = len(selected.encode())
    assert read(xml, max_bytes=length) is not None
    with pytest.raises(ValueError, match='byte budget'):
        read(xml, max_bytes=length - 1)
