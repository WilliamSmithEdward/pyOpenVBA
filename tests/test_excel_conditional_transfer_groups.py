"""Shared conditional groups stay linear during transfer admission and cloning."""
from __future__ import annotations

from collections.abc import Callable
from copy import deepcopy
from typing import cast

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _conditional_transfer as transfer
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._conditional import NS, with_rules


@pytest.mark.parametrize('count', [8, 48])
def test_shared_groups_are_validated_once_and_cloned_without_sibling_rules(
        count: int, monkeypatch: pytest.MonkeyPatch) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Sub Setup()
Dim n As Long,f As Object
For n=1 To ''' + str(count) + '''
Set f=Range("A1:A5").FormatConditions.Add(xlExpression,Formula1:="=A1>" & n)
Next
End Sub''', name='SharedGroups')
    app.run('Setup')
    sheet = book.sheets_[0]
    rules = sheet.conditional_formats
    container = XmlElement(NS + 'conditionalFormatting', {'sqref': 'A1:A5', 'xmlns:alias': 'urn:alias'})
    container.extend(rule.root for rule in rules)
    for rule in rules:
        rule.container = container
    root_ids = {id(rule.root) for rule in rules}
    validated_roots: list[int] = []
    validated_containers: list[XmlElement] = []
    original_attributes_known = cast('Callable[[XmlElement, set[str]], bool]', getattr(transfer, '_attributes_known'))
    original_validate_container = cast('Callable[[XmlElement], None]', getattr(transfer, '_validate_container'))

    def attributes_known(root: XmlElement, allowed: set[str]) -> bool:
        if id(root) in root_ids:
            validated_roots.append(id(root))
        return original_attributes_known(root, allowed)

    def validate_container(root: XmlElement) -> None:
        validated_containers.append(root)
        original_validate_container(root)

    def copy_without_siblings(root: XmlElement) -> XmlElement:
        assert root is not container, 'planning must not deepcopy shared sibling cfRule trees'
        return deepcopy(root)

    monkeypatch.setattr(transfer, '_attributes_known', attributes_known)
    monkeypatch.setattr(transfer, '_validate_container', validate_container)
    monkeypatch.setattr(transfer, 'deepcopy', copy_without_siblings)
    staged = transfer.plan(sheet, [Area(1, 1, 5, 1)], sheet, [Area(1, 7, 5, 7)])
    assert staged is not None
    assert validated_containers == [container]
    assert len(validated_roots) == count * 2
    assert set(validated_roots) == root_ids
    containers = [rule.container for rule in staged.incoming_rules]
    assert len(containers) == count and len({id(root) for root in containers}) == count
    for root in containers:
        assert root is not None and root is not container
        assert root.tag == container.tag and root.attrib == container.attrib and not list(root)
    assert len(container) == count
    staged.commit()
    rewritten = xml_tree(with_rules('<worksheet xmlns="' + NS[1:-1] + '"/>', rules))
    groups = rewritten.findall(NS + 'conditionalFormatting')
    assert len(groups) == count + 1
    assert len(rewritten.findall('.//' + NS + 'cfRule')) == count * 2
