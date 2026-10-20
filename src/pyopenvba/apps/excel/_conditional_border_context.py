"""Native color bindings for statistical conditions and their loaded identities."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._conditional import Rule
    from pyopenvba.apps.excel._model import Worksheet

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
STATISTICAL = frozenset(('top10', 'aboveAverage', 'uniqueValues', 'duplicateValues'))


def initialize(sheet: Worksheet) -> None:
    """Reindex live loaded rules in their format's native group-load order."""
    from pyopenvba.apps.excel._conditional import group_order

    if sheet.conditional_color_slots or sheet.conditional_color_next:
        return
    rules = list(sheet.conditional_formats)
    if rules and rules[0].imported:
        rules.sort(key=group_order)
    sheet.conditional_color_slots = dict(enumerate(rules, 1))
    for slot, rule in sheet.conditional_color_slots.items():
        rule.color_slot = slot
    sheet.conditional_color_next = len(rules)


def register(sheet: Worksheet, rule: Rule) -> None:
    """Assign a fresh identity at commit; deletion intentionally leaves its entry."""
    sheet.conditional_color_next += 1
    rule.color_slot = sheet.conditional_color_next
    sheet.conditional_color_slots[rule.color_slot] = rule


def _root(sheet: Worksheet, rule: Rule) -> XmlElement:
    if rule.dxf is not None:
        return rule.dxf
    identifier = rule.root.get('dxfId')
    return sheet.book.stylesheet.dxf(int(identifier)) if identifier is not None else XmlElement(NS + 'dxf')


def resolve(sheet: Worksheet, rule: Rule) -> XmlElement:
    """Resolve color getters only; stored styles and setters retain their owner."""
    if rule.color_slot is None or rule.root.get('type') not in STATISTICAL:
        return _root(sheet, rule)
    application = sheet.book.application
    active = application.active_book.active_sheet if application.active_book is not None else None
    if active is sheet:
        return _root(sheet, rule)
    bound = None if active is None else active.conditional_color_slots.get(rule.color_slot)
    if bound is None or active is None:
        raise error(440)
    return _root(active, bound)
