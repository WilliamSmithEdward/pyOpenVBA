"""What a chart reads from cells follows those cells as they move, as live Excel has it
(tests/fixtures/chart_references.json).

The movers work out what every reference becomes before anything moves, and carry it out afterwards. A point's look
and label stay with its cell, and moved to another place they lose their unique ids. A series' values becoming #REF!
take its points' looks and labels with them (tests/fixtures/chart_lost_cells.json). A chart read from a file keeps its
part, with only the references in it rewritten (tests/fixtures/loaded_chart_references.json); its points are numbered
inside that part, so cells moving between them are refused, and so is taking its cells away."""
from __future__ import annotations

import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import column_letter, parse_reference, quote_sheet
from pyopenvba._xml import escape_text, unescape
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import split_sheet, tokenize

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts import ChartState, SeriesState
    from pyopenvba.apps.excel._model import Workbook, Worksheet

#: A reference and the sheet whose formulas it reads as if it were one of, to what it becomes.
Change = Callable[[str, str], str]
Step = Callable[[], None]
#: A point's label that has lost its unique id, and still has to be written.
NO_ID = -2
#: A reference where a chart part spells one.
_PART_REFERENCE = re.compile(r'<((?:\w+:)?f|c15:sqref)>(.*?)</\1>', re.DOTALL)
#: What a chart part numbers by its place among the series' cells.
_NUMBERED = re.compile(r'<(?:\w+:)?(?:dPt|dLbl)>')


def renamed(formula: str, old: str, new: str) -> str:
    """``formula`` with its references to the sheet ``old`` naming it ``new``."""
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(formula):
        if token.kind != 'ref':
            continue
        sheet, reference = split_sheet(token.text)
        if sheet and sheet.casefold() == old.casefold():
            pieces.append((token.at, token.at + len(token.text), f'{quote_sheet(new)}!{reference}'))
    for first, last, replacement in reversed(pieces):
        formula = formula[:first] + replacement + formula[last:]
    return formula


def _slots(state: ChartState) -> list[tuple[object, str]]:
    """Where a chart keeps references: its series', and its titles' links."""
    slots: list[tuple[object, str]] = [(state.title_look, 'formula')]
    for axis in state.axes.values():
        if axis.title is not None:
            slots.append((axis.title.look, 'formula'))
    for entry in state.full_series or state.series:
        slots += [(entry, 'name_reference'), (entry, 'category_reference'), (entry, 'value_reference')]
    return slots


def _single(reference: str) -> tuple[str, int, int, int, int] | None:
    """A reference to one area, as its sheet's prefix and its corners; None for another."""
    try:
        areas = parse_reference(reference)
    except ValueError:
        return None
    if len(areas) != 1 or '!' not in reference:
        return None
    area = areas[0]
    return reference.rsplit('!', 1)[0] + '!', area.top, area.left, area.bottom, area.right


def _points_step(entry: SeriesState, turned: str, sheet: str, change: Change) -> Step | None:
    """The series' points going with their cells, when its values have moved."""
    before, after = _single(entry.value_reference), _single(turned)
    if before is None or after is None or not entry.points:
        return None
    prefix, top, left, bottom, right = before
    _, new_top, new_left, new_bottom, _ = after
    down = bottom > top
    places: dict[int, int] = {}
    for index in entry.points:
        row, column = (top + index, left) if down else (top, left + index)
        if row > bottom or column > right:
            continue
        cell = _single(change(f'={prefix}${column_letter(column)}${row}', sheet)[1:])
        if cell is not None:
            places[index] = cell[1] - new_top if down or new_bottom > new_top else cell[2] - new_left

    def step() -> None:
        kept = {}
        for index, point in entry.points.items():
            if index in places:
                if places[index] != index:
                    point.unique = -1
                    if point.label_unique >= 0:
                        point.label_unique = NO_ID
                kept[places[index]] = point
        entry.points = kept
    return step


def _values_gone(entry: SeriesState) -> Step:
    """What goes with a series' values: its points' looks and labels, its labels and its own look."""
    def step() -> None:
        entry.points = {}
        entry.data_labels = False
        entry.format_look = None
    return step


def _categories_gone(sheet: Worksheet, entry: SeriesState) -> Step:
    from pyopenvba.apps.excel._charts import lost, referenced_values

    numbers = not lost(entry.category_reference) and all(
        isinstance(value, (int, float)) and not isinstance(value, bool)
        for value in referenced_values(sheet, entry.category_reference))

    def step() -> None:
        entry.lost_numbers = numbers
    return step


def _extent(reference: str) -> tuple[int, int] | None:
    """How many rows and columns a reference to one area covers."""
    area = _single(reference)
    return None if area is None else (area[3] - area[1], area[4] - area[2])


def _part_step(state: ChartState, turned_text: Callable[[str], str]) -> Step | None:
    """The references in the part of a chart read from a file going with their cells."""
    numbered = _NUMBERED.search(state.original) is not None

    def follow(match: re.Match[str]) -> str:
        text = unescape(match.group(2))
        turned = turned_text(text)
        if turned == text:
            return match.group(0)
        if numbered and (_extent(text) is None or _extent(text) != _extent(turned)):
            raise VBAUnsupportedError('Moving cells between the points of a chart from a file is not implemented')
        return f'<{match.group(1)}>{escape_text(turned)}</{match.group(1)}>'

    rewritten = _PART_REFERENCE.sub(follow, state.original)
    if rewritten == state.original:
        return None

    def step() -> None:
        state.original = rewritten
        state.dirty = True
    return step


def plan(book: Workbook, change: Change) -> list[Step]:
    """All that moves, worked out before anything does; what cannot be done is refused here."""
    steps: list[Step] = []
    for sheet in book.sheets_:
        for state in sheet.chart_states.values():
            steps += _chart_plan(sheet, state, change)
    for chart_sheet in book.chart_sheets:
        steps += _chart_plan(chart_sheet.sheet, chart_sheet.state, change)
    return steps


def _chart_plan(sheet: Worksheet, state: ChartState, change: Change) -> list[Step]:
    steps: list[Step] = []

    def changed() -> None:
        state.dirty = True
        if not state.original:
            sheet.drawing_changed()

    def turned_text(text: str) -> str:
        turned = change(f'={text}', sheet.name)[1:]
        if turned != text and '#REF!' in turned and state.original:
            raise VBAUnsupportedError('Taking away the cells a chart from a file reads is not implemented')
        return turned

    def setting(holder: object, field: str, value: str) -> Step:
        def step() -> None:
            setattr(holder, field, value)
            changed()
        return step

    for holder, field in _slots(state):
        text: str = getattr(holder, field)
        if not text:
            continue
        turned = turned_text(text)
        if turned == text:
            continue
        if field == 'category_reference' and '#REF!' in turned:
            steps.append(_categories_gone(sheet, holder))  # pyright: ignore[reportArgumentType]
        if field == 'value_reference' and '#REF!' in turned:
            steps.append(_values_gone(holder))  # pyright: ignore[reportArgumentType]
        elif field == 'value_reference':
            moving = _points_step(holder, turned, sheet.name, change)  # pyright: ignore[reportArgumentType]
            if moving is not None:
                steps.append(moving)
        steps.append(setting(holder, field, turned))
    for entry in state.full_series or state.series:
        for bar in entry.error_bars:
            for field in ('plus', 'minus'):
                reference, values = getattr(bar, field)
                if reference:
                    turned = turned_text(reference)
                    if '#REF!' in turned:
                        raise VBAUnsupportedError('Taking away the cells an error bar reads is not implemented')
                    if turned != reference:
                        steps.append(setting(bar, field, (turned, values)))  # pyright: ignore[reportArgumentType]
    if state.original:
        following = _part_step(state, turned_text)
        if following is not None:
            steps.append(following)
    return steps


def changed_in(sheet: Worksheet, state: ChartState, change: Change) -> None:
    """Every reference of one chart changed now, as a sheet copied with it does."""
    carry_out(_chart_plan(sheet, state, change))


def carry_out(steps: list[Step]) -> None:
    for step in steps:
        step()
