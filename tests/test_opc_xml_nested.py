"""Nested OPC byte spans retain inherited namespace scopes and parser policy."""
from __future__ import annotations

import pytest

from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel._pivot_refresh import cache_order_bytes
from pyopenvba.exceptions import VBAUnsupportedError

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


@pytest.mark.parametrize('prefix', ['', 's:'])
def test_nested_container_uses_original_scope_and_absolute_spans(prefix: str) -> None:
    xmlns = 'xmlns:s' if prefix else 'xmlns'
    raw = (f'<{prefix}workbook {xmlns}="{NS}" xmlns:x="urn:opaque">'
           f'<{prefix}pivotCaches x:keep="a&gt;b"><!--stay-->'
           f'<{prefix}pivotCache cacheId="7" x:q="a&gt;b"/>'
           f'<x:keep/><{prefix}pivotCache cacheId="3"></{prefix}pivotCache>'
           f'</{prefix}pivotCaches><{prefix}extLst/></{prefix}workbook>').encode()
    group = opc_part(raw, NS, 'workbook', container_path=('pivotCaches',))
    assert group is not None and group.data is raw
    assert group.qname == prefix + 'pivotCaches'
    assert [item[0] for item in group.children] == [NS + '}pivotCache', 'urn:opaque}keep', NS + '}pivotCache']
    one, foreign, two = [raw[start:end] for _, _, start, end in group.children]
    ordered = cache_order_bytes(raw, [3, 7])
    assert ordered == raw.replace(one, b'FIRST', 1).replace(two, one, 1).replace(b'FIRST', two, 1)
    assert foreign in ordered and b'<!--stay-->' in ordered


def test_missing_self_closing_and_deep_container_paths() -> None:
    raw = f'<worksheet xmlns="{NS}"><sheetPr/></worksheet>'.encode()
    group = opc_part(raw, NS, 'worksheet', container_path=('sheetPr',))
    assert group is not None and group.children == []
    inserted = group.insert(group.closing_start, group.child('pageSetUpPr', 'fitToPage="1"'))
    assert b'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' in inserted
    assert opc_part(raw, NS, 'worksheet', container_path=('extLst',)) is None
    nested = f'<worksheet xmlns="{NS}"><sheetPr><pageSetUpPr/></sheetPr></worksheet>'.encode()
    found = opc_part(nested, NS, 'worksheet', container_path=('sheetPr', 'pageSetUpPr'))
    assert found is not None and found.qname == 'pageSetUpPr'


def test_unrelated_deep_branches_are_not_collected() -> None:
    raw = (f'<workbook xmlns="{NS}"><extLst>' + '<x>' * 1000 + '</x>' * 1000 +
           '</extLst><pivotCaches><pivotCache cacheId="1"/></pivotCaches></workbook>').encode()
    group = opc_part(raw, NS, 'workbook', container_path=('pivotCaches',))
    assert group is not None and len(group.children) == 1


@pytest.mark.parametrize('body', [
    '<!DOCTYPE worksheet><worksheet xmlns="NS"><sheetPr/></worksheet>',
    '<!DOCTYPE worksheet [<!ENTITY x "text">]><worksheet xmlns="NS"><sheetPr/></worksheet>',
    '<!DOCTYPE worksheet [<!ENTITY x SYSTEM "file:///not-read">]><worksheet xmlns="NS"><sheetPr/></worksheet>',
])
def test_nested_query_keeps_dtd_and_entity_refusal(body: str) -> None:
    with pytest.raises(ValueError, match='DTDs or declared entities'):
        opc_part(body.replace('NS', NS).encode(), NS, 'worksheet', container_path=('sheetPr',))


def test_ambiguous_container_and_cache_identity_are_refused() -> None:
    raw = f'<worksheet xmlns="{NS}"><sheetPr/><sheetPr/></worksheet>'.encode()
    with pytest.raises(ValueError, match='Ambiguous'):
        opc_part(raw, NS, 'worksheet', container_path=('sheetPr',))
    workbook = f'<workbook xmlns="{NS}"><pivotCaches><pivotCache cacheId="1"/><pivotCache cacheId="1"/></pivotCaches></workbook>'.encode()
    with pytest.raises(VBAUnsupportedError, match='ambiguous package cache identities'):
        cache_order_bytes(workbook, [1, 2])
