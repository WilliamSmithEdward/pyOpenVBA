"""Conditional formats Excel joins as it reads a file.

Measured in live Excel (scripts/measure_conditional_joins.py, tests/fixtures/conditional_joins/):

- A workbook Excel reads without a VBA project -- an .xlsx, or an .xlsm holding no code -- has its rules joined
  where they are alike: an area of a rule a formula decides (a cell value or an expression, a text, blank, error or
  date rule) that makes one block with an area of a rule alike before it in priority goes to that rule, the two
  areas becoming one, and the rule goes once it has no area left. Rules drawn from their range -- top or bottom,
  average, unique or duplicate, colour scales, data bars and icon sets -- are never joined. A workbook holding a VBA
  project keeps its rules as the file has them.
- Rules are alike when they are of one kind with the same settings -- operator, text, period and StopIfTrue --
  and formulas that read the same cells from each rule's own corner, and their formats show the same fill, font
  and border; a number format counts only where both have one.
- An area moves only where no rule between the two in priority covers any of its cells, whatever that rule's kind;
  a rule before or after them both holds nothing back.
- The rule an area joins keeps its priority and its format, and the other rules keep theirs, gaps and all.

The model joins areas in the measured shapes: one block with one other. An area that could join either of two
areas joins the first, and the areas a rule holds after a join are not joined with each other.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes
from pyopenvba.apps.excel._conditional import NS, RANGE_RULES, Rule, formula_anchor, meets, projected_areas
from pyopenvba.apps.excel._conditional_predicates import PREDICATE_TYPES, regenerate
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.formula._parse import FormulaError, shift_text
from pyopenvba.formula._r1c1 import from_a1

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: The kinds of rule a formula decides, which Excel joins.
_JOINED = frozenset(('cellIs', 'expression', *PREDICATE_TYPES))
Likeness = tuple[tuple[object, ...], tuple[str, str] | None]


def _priority(rule: Rule) -> int:
    try:
        return int(rule.root.get('priority', '0'))
    except ValueError:
        return 0


def _format(sheet: Worksheet, rule: Rule) -> XmlElement | None:
    """The rule's format; an empty dxf for a rule naming none. Raises IndexError for a dxf the file lacks."""
    identifier = rule.root.get('dxfId')
    if rule.dxf is not None:
        return rule.dxf
    if identifier is None or not identifier.isdigit():
        return None
    return sheet.book.stylesheet.dxf(int(identifier))


def _likeness(sheet: Worksheet, rule: Rule) -> Likeness | None:
    """What two rules alike share, and the number format their formats may differ by having one only on one side;
    None for a rule Excel does not join."""
    kind = rule.root.get('type', '')
    if kind not in _JOINED or kind in RANGE_RULES or rule.extension is not None or rule.extension_only \
            or not rule.areas:
        return None
    settings = tuple(sorted((key, value) for key, value in rule.root.attrib.items()
                            if key not in ('priority', 'dxfId') and key != 'xmlns' and not key.startswith('xmlns:')))
    anchor = formula_anchor(rule.areas)
    formulas: list[str] = []
    for child in rule.root.findall(NS + 'formula'):
        try:
            formulas.append(from_a1('=' + decode(child.text or ''), anchor.top, anchor.left))
        except FormulaError:
            return None
    if any(child.tag != NS + 'formula' for child in rule.root):
        return None
    try:
        shown = _format(sheet, rule)
    except IndexError:
        return None
    number: tuple[str, str] | None = None
    rest: bytes = b''
    if shown is not None:
        others = XmlElement(shown.tag, dict(shown.attrib))
        for child in shown:
            if child.tag == NS + 'numFmt':
                number = (child.get('formatCode', ''), child.get('numFmtId', ''))
            else:
                others.append(child)
        rest = xml_bytes(others)
    return (kind, settings, tuple(formulas), rest), number


def _alike(first: Likeness, second: Likeness) -> bool:
    return first[0] == second[0] and (first[1] is None or second[1] is None or first[1][0] == second[1][0])


def _block(first: Area, second: Area) -> Area | None:
    """The block two areas make together, where they make one."""
    top, left = min(first.top, second.top), min(first.left, second.left)
    bottom, right = max(first.bottom, second.bottom), max(first.right, second.right)

    def cells(area: Area) -> int:
        return area.rows * area.columns

    rows = min(first.bottom, second.bottom) - max(first.top, second.top) + 1
    columns = min(first.right, second.right) - max(first.left, second.left) + 1
    shared = rows * columns if rows > 0 and columns > 0 else 0
    whole = Area(top, left, bottom, right)
    return whole if cells(first) + cells(second) - shared == cells(whole) else None


def _rebased(rule: Rule, old: Area) -> None:
    """A joined rule's formulas for the corner of its new range."""
    new = formula_anchor(rule.areas)
    if (new.top, new.left) == (old.top, old.left):
        return
    if rule.root.get('type') in PREDICATE_TYPES:
        regenerate(rule.root, new)
        return
    for child in rule.root.findall(NS + 'formula'):
        if child.text:
            child.text = encode_text(shift_text('=' + decode(child.text), new.top - old.top, new.left - old.left)[1:])


def join(sheet: Worksheet) -> bool:
    """Join a sheet's rules alike as Excel joins them reading a file without a VBA project, as the module docstring
    has it; whether any were."""
    order = sorted((rule for rule in sheet.conditional_formats if not rule.deleted), key=_priority)
    likenesses = {id(rule): _likeness(sheet, rule) for rule in order}
    joined = False
    moving = True
    while moving:
        moving = False
        for later_at, later in enumerate(order):
            later_like = likenesses[id(later)]
            if later_like is None or not later.areas:
                continue
            for earlier in order[:later_at]:
                earlier_like = likenesses[id(earlier)]
                if earlier_like is None or not earlier.areas or not _alike(earlier_like, later_like):
                    continue
                between = [rule for rule in order
                           if _priority(earlier) < _priority(rule) < _priority(later) and rule.areas]
                # Each rule's areas in the order Excel holds them once it has read the file.
                for area in projected_areas(later):
                    if any(meets(area, other) for rule in between for other in rule.areas):
                        continue
                    target = next(((held, block) for held in projected_areas(earlier)
                                   if (block := _block(held, area)) is not None), None)
                    if target is None:
                        continue
                    old = formula_anchor(earlier.areas)
                    earlier.areas[earlier.areas.index(target[0])] = target[1]
                    later.areas.remove(area)
                    _rebased(earlier, old)
                    moving = joined = True
                if not later.areas:
                    break
    for rule in order:
        if not rule.areas:
            rule.deleted = True
            sheet.conditional_formats.remove(rule)
    return joined
