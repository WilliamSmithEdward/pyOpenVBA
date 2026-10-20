"""Native color bindings for statistical conditions and their loaded identities."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._conditional import Rule
    from pyopenvba.apps.excel._model import Worksheet

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
STATISTICAL = frozenset(('top10', 'aboveAverage', 'uniqueValues', 'duplicateValues'))


def initialize(sheet: Worksheet) -> None:
    """Reindex live loaded rules in their format's native group-load order."""
    from pyopenvba.apps.excel._conditional import rule_groups

    if sheet.conditional_color_slots or sheet.conditional_color_next:
        return
    groups = rule_groups(sheet.conditional_formats)
    sheet.conditional_color_slots = {slot: group[0] for slot, group in enumerate(groups, 1)}
    for slot, group in enumerate(groups, 1):
        for order, rule in enumerate(group):
            rule.color_slot = slot
            rule.child_order = order
    sheet.conditional_color_next = len(groups)
    sheet.conditional_color_groups = None


def register(sheet: Worksheet, rule: Rule) -> None:
    """Assign a fresh identity at commit; deletion intentionally leaves its entry."""
    sheet.conditional_color_groups = None
    sheet.conditional_color_next += 1
    rule.color_slot = sheet.conditional_color_next
    sheet.conditional_color_slots[rule.color_slot] = rule


def _siblings(sheet: Worksheet, rule: Rule) -> list[Rule]:
    if rule.child_order is None or rule.color_slot is None:
        return []
    return sorted((other for other in sheet.conditional_formats
                   if not other.deleted and other.color_slot == rule.color_slot),
                  key=lambda other: other.child_order if other.child_order is not None else 0)


def delete_member(sheet: Worksheet, rule: Rule) -> None:
    """An individual deletion retains the target's slot and materializes survivors together."""
    siblings = _siblings(sheet, rule)
    if len(siblings) < 2:
        return
    assert rule.color_slot is not None
    sheet.conditional_color_slots[rule.color_slot] = rule
    survivors = [other for other in siblings if other is not rule]
    register(sheet, survivors[0])
    for other in survivors[1:]:
        other.color_slot = survivors[0].color_slot


def move_member(sheet: Worksheet, rule: Rule) -> None:
    """A shared-rule range change appends the moved member after the untouched group."""
    siblings = _siblings(sheet, rule)
    if len(siblings) < 2:
        return
    assert rule.color_slot is not None
    sheet.conditional_color_slots[rule.color_slot] = next(other for other in siblings if other is not rule)
    register(sheet, rule)
    rule.child_order = 0


def _root(sheet: Worksheet, rule: Rule) -> XmlElement:
    if rule.dxf is not None:
        return rule.dxf
    identifier = rule.root.get('dxfId')
    return sheet.book.stylesheet.dxf(int(identifier)) if identifier is not None else XmlElement(NS + 'dxf')


def _group(sheet: Worksheet, slot: int) -> tuple[list[Rule], dict[Rule, int]]:
    if sheet.conditional_color_groups is None:
        members: dict[int, list[Rule]] = {}
        for rule in sheet.conditional_formats:
            if not rule.deleted and rule.color_slot is not None:
                members.setdefault(rule.color_slot, []).append(rule)
        groups: dict[int, tuple[list[Rule], dict[Rule, int]]] = {}
        for number, representative in sheet.conditional_color_slots.items():
            entries = members.get(number, [])
            if representative.deleted or representative not in entries:
                entries = [representative]
            else:
                entries.sort(key=lambda rule: rule.child_order if rule.child_order is not None else 0)
            groups[number] = entries, {rule: index for index, rule in enumerate(entries)}
        sheet.conditional_color_groups = groups
    group = sheet.conditional_color_groups.get(slot)
    if group is None:
        raise error(440)
    return group


def bound_rule(sheet: Worksheet, rule: Rule) -> tuple[Worksheet, Rule]:
    """Resolve a foreign group and compact child position without changing either owner."""
    if rule.color_slot is None:
        return sheet, rule
    application = sheet.book.application
    active = application.active_book.active_sheet if application.active_book is not None else None
    if active is sheet:
        return sheet, rule
    bound = None if active is None else active.conditional_color_slots.get(rule.color_slot)
    if bound is None or active is None:
        raise error(440)
    position = 0
    if rule.child_order is not None and sheet.conditional_color_slots.get(rule.color_slot) is not rule:
        position = _group(sheet, rule.color_slot)[1].get(rule, 0)
    if position:
        targets = _group(active, rule.color_slot)[0]
        if position >= len(targets):
            raise VBAUnsupportedError('Foreign conditional formatting lookup of a missing group member is not implemented')
        bound = targets[position]
    return active, bound


def resolve(sheet: Worksheet, rule: Rule) -> XmlElement:
    """Resolve foreign member colors while owner reads and stored styles stay independent."""
    if rule.root.get('type') not in STATISTICAL:
        return _root(sheet, rule)
    owner, bound = bound_rule(sheet, rule)
    return _root(owner, bound)
