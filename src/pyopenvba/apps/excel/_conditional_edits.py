"""Conditional formats through rows, columns and cells inserted and deleted.

Measured in live Excel (scripts/measure_conditional_edits.py, tests/fixtures/conditional_edits/):

- Through whole rows or columns a rule's range moves as a reference does: an area shifts, stretches for lines
  inserted inside it, shrinks for lines deleted from it and goes when all of its lines go; one spanning every line
  stays whole. Inserted lines take the rules of the line they copy their formats from, the one before them, or
  with CopyOrigin xlFormatFromRightOrBelow the one after. A delete puts the areas it cuts after the others, and
  joins areas it brings together into one, in the later one's place.
- Cells shifted down, up, right or left carry an area that lies in the band, and one spanning every row (or
  column) stays whole. A block the band runs through is cut into strips: its rows above the moved cells and below
  them at full width, and the parts beside them, then the cells that moved; a shift across puts the part before
  the moved cells back with them, in their place. Cells an insert makes take the rules of the cells before them
  (after, with xlFormatFromRightOrBelow), joining the moved cells or added after every area. A delete taking every
  cell a block has in the band keeps a rule a formula decides as a reference keeps it, whole or without that edge,
  and cuts a rule drawn from its range into strips like any other.
- A rule's formulas move as a cell's do, a colour scale's, data bar's or icon set's points among them, and a
  point naming the sheet from another: a reference to moved cells follows them and one to deleted cells becomes
  #REF!.
- Excel makes a rule anew when a delete turns one of its references to #REF!, or when it shrinks an area of a top
  or bottom, average, unique or duplicate rule, a colour scale, a data bar or an icon set as it shrinks whole
  lines, unless one of the rule's areas moves whole: then the first such area leads the range and the rule keeps
  its priority. A rule made anew takes the workbook's highest priority plus one, with its range's other rules after
  it in their order; before it, each rule the delete cuts on the sheet ahead of it uses up two more. Nothing else
  an edit does changes a priority.

What was not measured reports itself: a relative reference the edit leaves pointing somewhere else for some cells
than for others, which Excel answers by splitting the rule; one edit making several ranges' rules anew, or one
alongside other rules it cuts, removes, or changes on other sheets, but for single rules it cuts ahead on the same
sheet with a whole-row or column delete; a rule made anew beside two rules alike, which Excel joins; areas an
insert or a cell shift brings together, or a delete brings together in a range made anew or one already holding
areas that would join; a range holding rules of both kinds where they part; a reference pushed off the sheet;
and an Excel 2010 extension holding rules the model does not read that name the edited sheet.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_number
from pyopenvba.apps.excel._charts_file import xml_bytes
from pyopenvba.apps.excel._conditional import (
    NS,
    RANGE_RULES,
    Rule,
    changed,
    formula_anchor,
    projected_areas,
    rule_groups,
    sqref,
)
from pyopenvba.apps.excel._editing import CellShift, interval, rewrite, rewrite_shift, shifted_box
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, shift_text, split_sheet, tokenize

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._data_bars import Point
    from pyopenvba.apps.excel._icon_sets import Criterion
    from pyopenvba.apps.excel._model import Worksheet

_CORNER = re.compile(r'(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?')
Box = tuple[int, int, int, int]


@dataclass(frozen=True)
class LineEdit:
    """Whole rows (``rows``) or columns: ``count`` going in or out at ``start``."""

    rows: bool
    start: int
    count: int
    delete: bool


Edit = LineEdit | CellShift


def _unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError(f'Inserting or deleting cells {detail} is not implemented')


def _along(edit: Edit) -> bool:
    """Whether the edit moves cells along the rows, down or up, rather than across the columns."""
    return edit.rows if isinstance(edit, LineEdit) else edit.vertical


# --- where a range goes ---------------------------------------------------------------------------------------


@dataclass
class _Range:
    """A rule's range once the edit is made, in Excel's order: each area and what the edit did to it -- kept,
    moved whole, grown, cut as whole lines are, split into strips, or joined with another; whether an area shrank
    or went as whole lines do, and whether one went."""

    areas: list[Box] = field(default_factory=lambda: [])
    kinds: list[str] = field(default_factory=lambda: [])
    shrunk: bool = False
    removed: bool = False

    def add(self, box: Box, kind: str) -> None:
        self.areas.append(box)
        self.kinds.append(kind)


def _rectangle(first: Box, second: Box) -> Box | None:
    """Two blocks as one, where together they are one."""
    if first[1] == second[1] and first[3] == second[3] and (first[2] + 1 == second[0] or second[2] + 1 == first[0]):
        return min(first[0], second[0]), first[1], max(first[2], second[2]), first[3]
    if first[0] == second[0] and first[2] == second[2] and (first[3] + 1 == second[1] or second[3] + 1 == first[1]):
        return first[0], min(first[1], second[1]), first[2], max(first[3], second[3])
    return None


def _joins(boxes: list[Box]) -> bool:
    return any(_rectangle(one, other) is not None for index, one in enumerate(boxes) for other in boxes[index + 1:])


def _joined(result: _Range) -> _Range:
    """Areas brought together joined into one, in the later one's place, as a delete leaves them."""
    out = _Range(shrunk=result.shrunk, removed=result.removed)
    for box, kind in zip(result.areas, result.kinds, strict=True):
        while True:
            index = next((at for at, held in enumerate(out.areas) if _rectangle(held, box) is not None), None)
            if index is None:
                break
            joined = _rectangle(out.areas[index], box)
            assert joined is not None
            del out.areas[index], out.kinds[index]
            box, kind = joined, 'joined'
        out.add(box, kind)
    return out


def _line_range(boxes: list[Box], edit: LineEdit, below: bool) -> _Range:
    """A range through whole rows or columns, as the module docstring has it."""
    rows, start, count, delete = edit.rows, edit.start, edit.count, edit.delete
    limit = MAX_ROWS if rows else MAX_COLUMNS
    end = start + count - 1
    out, cut = _Range(), _Range()

    def placed(box: Box, low: int, high: int) -> Box:
        return (low, box[1], high, box[3]) if rows else (box[0], low, box[2], high)

    for box in boxes:
        low, high = (box[0], box[2]) if rows else (box[1], box[3])
        if (low, high) == (1, limit) or high < start - (0 if delete else 1):
            out.add(box, 'kept')
        elif not delete:
            if high + count > limit:
                raise _unsupported('pushing a conditional formatting range off the sheet')
            if high == start - 1:
                out.add(box if below else placed(box, low, high + count), 'kept' if below else 'grown')
            elif low >= start:
                out.add(placed(box, start if below and low == start else low + count, high + count),
                        'grown' if below and low == start else 'moved')
            else:
                out.add(placed(box, low, high + count), 'grown')
        elif low > end:
            out.add(placed(box, low - count, high - count), 'moved')
        else:
            moved = interval(low, high, start, count, delete, limit)
            if moved is None:
                out.removed = True
            else:
                cut.add(placed(box, *moved), 'cut')
            out.shrunk = True
    if not delete:
        return out
    whole = _Range(out.areas + cut.areas, out.kinds + cut.kinds, out.shrunk, out.removed)
    return _joined(whole) if whole.kinds.count('kept') < len(boxes) else whole


def _shift_range(boxes: list[Box], shift: CellShift, below: bool, whole_blocks: bool) -> _Range:
    """A range through a cell shift, as the module docstring has it; ``whole_blocks`` for a rule a formula
    decides, whose blocks a delete taking all their band keeps as references are kept."""
    vertical, band, start, count, delete = shift.vertical, shift.band, shift.start, shift.count, shift.delete
    limit = MAX_ROWS if vertical else MAX_COLUMNS
    end = start + count - 1
    out = _Range()
    tail: list[Box] = []

    def placed(along: tuple[int, int], across: tuple[int, int]) -> Box:
        return (along[0], across[0], along[1], across[1]) if vertical else (across[0], along[0], across[1], along[1])

    for box in boxes:
        top, left, bottom, right = box
        along, across = ((top, bottom), (left, right)) if vertical else ((left, right), (top, bottom))
        inside = (max(across[0], band[0]), min(across[1], band[1]))
        if inside[0] > inside[1] or along == (1, limit) or along[1] < start - (0 if delete else 1):
            out.add(box, 'kept')
            continue
        if not delete and along[1] + count > limit:
            raise _unsupported('pushing a conditional formatting range off the sheet')
        if along[1] < start:
            # The block ends just before the inserted cells, which take its rules from it.
            if below:
                out.add(box, 'kept')
                continue
            made = placed((start, end), inside)
            joined = _rectangle(box, made)
            if joined is None:
                out.add(box, 'kept')
                tail.append(made)
            else:
                out.add(joined, 'grown')
            continue
        if band[0] <= across[0] and across[1] <= band[1]:
            # The block lies in the band: it moves as whole lines move it.
            moved = shifted_box(box, shift)
            if moved is None:
                out.removed = out.shrunk = True
                continue
            if not delete and below and along[0] == start:
                out.add(placed((start, moved[2] if vertical else moved[3]), across), 'grown')
            elif delete:
                out.shrunk = out.shrunk or along[0] <= end
                out.add(moved, 'moved' if along[0] > end else 'cut')
            else:
                out.add(moved, 'moved' if along[0] >= start else 'grown')
            continue
        if whole_blocks and delete and start <= along[0] and along[1] <= end:
            # A delete takes every cell the block has in the band: it keeps the rest as a reference does.
            moved = shifted_box(box, shift)
            assert moved is not None
            out.add(moved, 'split')
            continue
        # The band runs through the block: strips round the moved cells, then the moved cells.
        pieces: list[Box] = []
        first = max(along[0], start)
        if vertical:
            if top < start:
                pieces.append((top, left, start - 1, right))
            if left < band[0]:
                pieces.append((first, left, bottom, band[0] - 1))
            if right > band[1]:
                pieces.append((first, band[1] + 1, bottom, right))
        else:
            if top < band[0]:
                pieces.append((top, left, band[0] - 1, right))
            if left < start:
                pieces.append((inside[0], left, inside[1], start - 1))
            if bottom > band[1]:
                pieces.append((band[1] + 1, left, bottom, right))
        if delete:
            low, high = max(first, start + count) - count, along[1] - count
        else:
            low, high = first + count, along[1] + count
            source = start if below else start - 1
            if along[0] <= source <= along[1] and first == start:
                low = start
        if low <= high:
            moved_box = placed((low, high), inside)
            for index, piece in enumerate(pieces):
                joined = _rectangle(piece, moved_box)
                if joined is not None:
                    del pieces[index]
                    moved_box = joined
                    break
            pieces.append(moved_box)
        for piece in pieces:
            out.add(piece, 'split')
    for made in tail:
        out.add(made, 'grown')
    return out


def _range(boxes: list[Box], edit: Edit, below: bool, whole_blocks: bool) -> _Range:
    if not boxes:
        return _Range()
    if isinstance(edit, LineEdit):
        result = _line_range(boxes, edit, below)
    else:
        result = _shift_range(boxes, edit, below, whole_blocks)
    if result.areas != boxes:
        # A whole-line delete joins what it brings together; areas that could already join, or that an insert or
        # a shift brings together, were not measured.
        joining = boxes if isinstance(edit, LineEdit) and edit.delete else result.areas
        if _joins(joining):
            raise _unsupported('bringing conditional formatting areas together')
    return result


# --- what formulas become -------------------------------------------------------------------------------------


def _position(edit: Edit, row: int, column: int) -> tuple[int, int] | None:
    """Where a cell of the edited sheet goes; None when it is deleted."""
    down = _along(edit)
    along, across = (row, column) if down else (column, row)
    if isinstance(edit, CellShift) and not edit.band[0] <= across <= edit.band[1]:
        return row, column
    moved = interval(along, along, edit.start, edit.count, edit.delete, MAX_ROWS if down else MAX_COLUMNS)
    if moved is None:
        return None
    return (moved[0], column) if down else (row, moved[0])


def _bounds(edit: Edit, text: str, anchor: Area, owner: str, edited: str, on_edited: bool) -> tuple[set[int], ...]:
    """The rows and the columns of a rule's cells where what the edit does to a cell, or to a reference the
    formula reads from it, changes."""
    down = _along(edit)
    lines = (edit.start, edit.start + edit.count)
    band = (edit.band[0], edit.band[1] + 1) if isinstance(edit, CellShift) else ()
    rows: set[int] = set()
    columns: set[int] = set()
    if on_edited:
        (rows if down else columns).update(lines)
        (columns if down else rows).update(band)
    for token in tokenize(text):
        if token.kind != 'ref':
            continue
        named, reference = split_sheet(token.text)
        if (named or owner).casefold() != edited.casefold():
            continue
        for part in reference.split(':'):
            found = _CORNER.fullmatch(part)
            if found is None:
                continue
            column_mark, letters, row_mark, digits = found.groups()
            # A relative reference reads the cell its distance away, crossing an edit's line at the cell that
            # distance short of it.
            if digits and not row_mark:
                rows.update(line - (int(digits) - anchor.top) for line in (lines if down else band))
            if letters and not column_mark:
                columns.update(line - (column_number(letters) - anchor.left) for line in (band if down else lines))
    return rows, columns


def _samples(areas: list[Area], bounds: tuple[set[int], ...]) -> list[tuple[int, int]]:
    """A cell at each end of every stretch of the range the edit treats alike."""
    rows, columns = bounds

    def ends(low: int, high: int, cuts: set[int]) -> list[int]:
        found = {low, high}
        for cut in cuts:
            if low < cut <= high:
                found.update((cut, cut - 1))
        return sorted(found)

    return [(row, column) for area in areas for row in ends(area.top, area.bottom, rows)
            for column in ends(area.left, area.right, columns)]


def _lost(formula: str) -> int:
    return sum(1 for token in tokenize(formula.removeprefix('=')) if token.kind == 'error'
               and token.text.upper().endswith('#REF!'))


def _moved_formula(text: str, areas: list[Area], new_areas: list[Area], owner: str, edited: str, edit: Edit,
                   on_edited: bool) -> tuple[str, bool]:
    """A rule's formula once the edit has moved its cells and those it reads, for the new range's corner, and
    whether a reference became #REF!."""
    anchor = formula_anchor(areas)
    target = formula_anchor(new_areas) if new_areas else anchor
    found: set[tuple[str, bool]] = set()
    try:
        before = _lost(text)
        for row, column in _samples(areas, _bounds(edit, text, anchor, owner, edited, on_edited)):
            position = _position(edit, row, column) if on_edited else (row, column)
            if position is None:
                continue
            here = shift_text('=' + text, row - anchor.top, column - anchor.left)
            if isinstance(edit, LineEdit):
                moved = rewrite(here, owner, edited, rows=edit.rows, start=edit.start, count=edit.count,
                                delete=edit.delete)
            else:
                moved = rewrite_shift(here, owner, edited, edit)
            out = shift_text(moved, target.top - position[0], target.left - position[1])
            if _lost(here) != before or _lost(out) != _lost(moved):
                raise _unsupported('where a conditional formula reaches past the sheet')
            found.add((out[1:], _lost(moved) > before))
    except FormulaError:
        raise _unsupported('where a conditional formula cannot be read') from None
    if len(found) > 1:
        raise _unsupported('where conditional formatting cells would need formulas of their own')
    return found.pop() if found else (text, False)


def _texts(rule: Rule) -> list[tuple[str, Callable[[str], None]]]:
    """Each formula a rule holds, its points' among them, with how to put its new text."""
    from pyopenvba.apps.excel._conditional_extension import NUMBER

    found: list[tuple[str, Callable[[str], None]]] = []

    def formula(child: XmlElement) -> Callable[[str], None]:
        def put(text: str) -> None:
            child.text = encode_text(text)
        return put

    for child in rule.root.findall(NS + 'formula'):
        found.append((decode(child.text or ''), formula(child)))
    kind = rule.root.get('type')
    if kind == 'colorScale':
        def point(element: XmlElement) -> Callable[[str], None]:
            return lambda text: element.set('val', text)

        for element in rule.root.findall(NS + 'colorScale/' + NS + 'cfvo'):
            value = element.get('val', '')
            if element.get('type') not in ('min', 'max') and value and not NUMBER.fullmatch(value):
                found.append((value, point(element)))
    elif kind in ('dataBar', 'iconSet'):
        from pyopenvba.apps.excel._data_bars import Bar
        from pyopenvba.apps.excel._data_bars import parse as parse_bar
        from pyopenvba.apps.excel._icon_sets import Icons
        from pyopenvba.apps.excel._icon_sets import parse as parse_icons

        if rule.state is None:
            rule.state = parse_bar(rule) if kind == 'dataBar' else parse_icons(rule)
        state = rule.state

        def held(holder: Point | Criterion) -> Callable[[str], None]:
            def put(text: str) -> None:
                holder.value = text
                state.dirty = True
            return put

        holders: list[Point] | list[Criterion] = (state.points if isinstance(state, Bar) else state.criteria
                                                  if isinstance(state, Icons) else [])
        for holder in holders:
            if isinstance(holder.value, str) and holder.value:
                found.append((holder.value, held(holder)))
    return found


# --- the plan -------------------------------------------------------------------------------------------------


@dataclass
class _Group:
    """What the edit does to one range's rules: its areas afterwards, each rule's new formulas, and whether Excel
    makes the rules anew."""

    owner: Worksheet
    rules: list[Rule]
    areas: list[Area]
    result: _Range
    texts: list[tuple[str, Callable[[str], None]]]
    triggered: bool
    remade: bool

    @property
    def cut(self) -> bool:
        return self.result.removed or any(kind in ('cut', 'split', 'joined') for kind in self.result.kinds)


def _priority(rule: Rule) -> int:
    try:
        return int(rule.root.get('priority', '0'))
    except ValueError:
        return 0


def _reads(text: str, owner: str, edited: str) -> bool:
    """Whether a formula on another sheet reads the edited one."""
    try:
        return any(token.kind == 'ref' and (split_sheet(token.text)[0] or owner).casefold() == edited.casefold()
                   for token in tokenize(text))
    except FormulaError:
        return True


def _group(owner: Worksheet, rules: list[Rule], sheet: Worksheet, edit: Edit, below: bool) -> _Group:
    on_edited = owner is sheet
    old = projected_areas(rules[0])
    boxes = [(area.top, area.left, area.bottom, area.right) for area in old]
    drawn = [rule.root.get('type') in RANGE_RULES for rule in rules]
    if not on_edited:
        result = _Range(list(boxes), ['kept'] * len(boxes))
        if not any(_reads(text, owner.name, sheet.name) for rule in rules for text, _ in _texts(rule)):
            return _Group(owner, rules, [Area(*box) for box in boxes], result, [], False, False)
    else:
        result = _range(boxes, edit, below, not drawn[0])
        if any(one != drawn[0] for one in drawn) and _range(boxes, edit, below, drawn[0]) != result:
            raise _unsupported('parting a conditional formatting range whose rules part differently')
    areas = [Area(*box) for box in result.areas]
    texts: list[tuple[str, Callable[[str], None]]] = []
    lost = False
    if areas:
        for rule in rules:
            for text, put in _texts(rule):
                new, gone = _moved_formula(text, old, areas, owner.name, sheet.name, edit, on_edited)
                lost = lost or gone
                if new != text:
                    texts.append((new, put))
    triggered = bool(areas) and (lost or (on_edited and edit.delete and result.shrunk and any(drawn)))
    remade = triggered and 'moved' not in result.kinds
    if triggered and not remade:
        if 'joined' in result.kinds:
            raise _unsupported('bringing conditional formatting areas together in a rule Excel would make anew')
        # The rule stays by the first of its areas that moved whole, which leads its range.
        lead = result.kinds.index('moved')
        areas.insert(0, areas.pop(lead))
        result.kinds.insert(0, result.kinds.pop(lead))
    return _Group(owner, rules, areas, result, texts, triggered, remade)


def _names(owner: Worksheet, edited: Worksheet) -> bool:
    """Whether the rules of a sheet's Excel 2010 extension the model does not read may name the edited sheet, as
    the file holds them."""
    from pyopenvba.apps.excel._charts_file import xml_tree
    from pyopenvba.apps.excel._conditional_extension import RULES_URI, XM

    package = owner.book.package
    if package is None or not owner.part_name or not package.has(owner.part_name):
        return True
    for ext in xml_tree(package.read(owner.part_name).decode('utf-8')).iter(NS + 'ext'):
        if ext.get('uri') != RULES_URI:
            continue
        for found in ext.iter(XM + 'f'):
            try:
                tokens = tokenize(found.text or '')
            except FormulaError:
                return True
            if any(token.kind == 'ref' and split_sheet(token.text)[0].casefold() == edited.name.casefold()
                   for token in tokens):
                return True
    return False


def _likeness(owner: Worksheet, rule: Rule) -> tuple[object, ...]:
    """What two rules Excel would join share: their kind and settings, their formulas as each cell reads them, and
    their format."""
    anchor = formula_anchor(rule.areas)
    formulas: list[str] = []
    for child in rule.root.findall(NS + 'formula'):
        try:
            formulas.append(shift_text('=' + decode(child.text or ''), 1000 - anchor.top, 1000 - anchor.left))
        except FormulaError:
            formulas.append(child.text or '')
    identifier = rule.root.get('dxfId')
    dxf = rule.dxf if rule.dxf is not None or identifier is None else owner.book.stylesheet.dxf(int(identifier))
    return (tuple(sorted((key, value) for key, value in rule.root.attrib.items() if key not in ('priority', 'dxfId'))),
            tuple(formulas), b'' if dxf is None else xml_bytes(dxf))


def planned(sheet: Worksheet, edit: Edit, *, below: bool = False) -> Callable[[], None]:
    """What an insert or delete does to every conditional format in the workbook, worked out and refused where not
    measured before anything changes; the call it answers makes those changes once the cells have moved."""
    book = sheet.book
    if any(one.conditional_foreign and (one is sheet or _names(one, sheet)) for one in book.sheets_):
        raise _unsupported("where an Excel 2010 extension holds conditional formats the model does not read")
    groups: list[_Group] = []
    for owner in book.sheets_:
        groups.extend(_group(owner, group, sheet, edit, below) for group in rule_groups(owner.conditional_formats))
    old_areas = {id(group): projected_areas(group.rules[0]) for group in groups}
    changes = [group for group in groups if group.areas != old_areas[id(group)] or group.texts or group.remade]
    remade = [group for group in changes if group.remade]
    highest = max((_priority(rule) for group in groups for rule in group.rules), default=0)
    extra = 0
    if remade:
        if len(remade) > 1:
            raise _unsupported('where several conditional formatting ranges would be made anew')
        made = remade[0]
        ahead = groups[:groups.index(made)]
        for group in changes:
            if group is made:
                continue
            if group.owner is not sheet or group.triggered or not group.areas:
                raise _unsupported('changing other conditional formats where one is made anew')
            if group.cut and group in ahead:
                # Each rule cut ahead of the one made anew uses up two priorities (rx_cut_before, lx_cell_cols_bar).
                if group.owner is not made.owner or group.result.removed or len(group.rules) > 1 or \
                        len(group.areas) > 1 or not isinstance(edit, LineEdit):
                    raise _unsupported('cutting other conditional formats where one is made anew')
                extra += 2
        # Excel joins rules alike that the edit cuts as it makes one anew (rx_two_cut_before), and leaves those it
        # does not (n_groups_meet).
        alike = [(group in changes, _likeness(group.owner, rule)) for group in groups
                 if group is not made and group.owner is made.owner
                 for rule in group.rules if rule.root.get('type') not in RANGE_RULES]
        if any(edited and sum(other == likeness for _, other in alike) > 1 for edited, likeness in alike):
            raise _unsupported('where Excel would join rules alike as it makes another anew')

    def commit() -> None:
        touched: set[int] = set()
        for group in changes:
            owner = group.owner
            for text, put in group.texts:
                put(text)
            for rule in group.rules:
                if group.areas != old_areas[id(group)]:
                    _moved(rule, list(group.areas))
                if not group.areas:
                    rule.deleted = True
                    owner.conditional_formats.remove(rule)
            touched.add(id(owner))
        for group in remade:
            from pyopenvba.apps.excel._conditional_border_context import register

            number = highest + extra
            for rule in sorted(group.rules, key=_priority):
                number += 1
                rule.root.set('priority', str(number))
                _made_anew(rule)
                register(group.owner, rule)
                group.owner.conditional_formats.remove(rule)
                group.owner.conditional_formats.append(rule)
        for owner in book.sheets_:
            if id(owner) in touched:
                changed(owner)

    return commit


def _moved(rule: Rule, areas: list[Area]) -> None:
    """A rule's new range, which it now shows in the order the edit left, its x14:cfRule keeping its place."""
    rule.areas = areas
    rule.imported = False
    if rule.extension_place is not None:
        place = rule.extension_place
        rule.extension_place = (place[0], sqref(areas), place[2])


def _made_anew(rule: Rule) -> None:
    """A rule Excel makes anew: as a rule made in this session, written as Excel writes one, its x14:cfRule keeping
    its id (tests/fixtures/conditional_edits/conditional_edits_resaved.xlsx, lx_bar_remade)."""
    rule.container = None
    rule.imported = rule.loaded = False
    rule.plain_load = None
    rule.child_order = None
    rule.extension_place = None
    if rule.state is not None:
        rule.state.dirty = True
