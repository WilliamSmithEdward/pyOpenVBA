"""Opt-in OPC span budgets preserve default namespace/DTD validation."""
import pytest

from pyopenvba._opc_xml import opc_part

NS = 'urn:budget'


def test_empty_exact_and_overflow_direct_children() -> None:
    empty = b'<root xmlns="urn:budget"/>'
    assert opc_part(empty, NS, 'root', max_children=1) == opc_part(empty, NS, 'root')
    one = b'<root xmlns="urn:budget"><child><grandchild/></child></root>'
    assert opc_part(one, NS, 'root', max_children=1) == opc_part(one, NS, 'root')
    with pytest.raises(ValueError, match='direct-child budget'):
        opc_part(b'<root xmlns="urn:budget"><a/><b/><unclosed', NS, 'root', max_children=1)
    with pytest.raises(ValueError, match='positive'):
        opc_part(empty, NS, 'root', max_children=0)


def test_nested_budget_counts_only_selected_container_children() -> None:
    xml = b'<root xmlns="urn:budget"><outer/><selected><a><nested/></a></selected><after/></root>'
    assert opc_part(xml, NS, 'root', container_path=('selected',), max_children=1) == opc_part(
        xml, NS, 'root', container_path=('selected',))
    with pytest.raises(ValueError, match='direct-child budget'):
        opc_part(xml.replace(b'</selected>', b'<b/></selected>'), NS, 'root',
                 container_path=('selected',), max_children=1)


@pytest.mark.parametrize('path', [None, ('selected',)])
def test_budget_does_not_relax_dtd_or_entity_rejection(path: tuple[str, ...] | None) -> None:
    xml = b'<!DOCTYPE root [<!ENTITY e "value">]><root xmlns="urn:budget"><selected>&e;</selected></root>'
    with pytest.raises(ValueError, match='DTDs'):
        if path is None:
            opc_part(xml, NS, 'root', max_children=1)
        else:
            opc_part(xml, NS, 'root', container_path=path, max_children=1)
