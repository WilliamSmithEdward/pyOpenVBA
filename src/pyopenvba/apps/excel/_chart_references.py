"""What a chart reads from cells follows those cells as they move, as live Excel has it
(tests/fixtures/chart_references.json).

The movers work out what every reference becomes before anything moves, and carry it out afterwards. A point's look
and label stay with its cell, and moved to another place they lose their unique ids. A reference that would become
#REF! is refused, because Excel then writes extensions the model does not hold, and so is a moved reference of a chart
read from a file, whose part is kept as it came."""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import column_letter, parse_reference, quote_sheet
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


def plan(book: Workbook, change: Change) -> list[Step]:
    """All that moves, worked out before anything does; what cannot be done is refused here."""
    steps: list[Step] = []
    for sheet in book.sheets_:
        for state in sheet.chart_states.values():
            steps += _chart_plan(sheet, state, change)
    return steps


def _chart_plan(sheet: Worksheet, state: ChartState, change: Change) -> list[Step]:
    steps: list[Step] = []

    def changed() -> None:
        state.dirty = True
        sheet.drawing_changed()

    def turned_text(text: str) -> str:
        turned = change(f'={text}', sheet.name)[1:]
        if turned != text:
            if state.original:
                raise VBAUnsupportedError('Moving the cells a chart from a file reads is not implemented')
            if '#REF!' in turned:
                raise VBAUnsupportedError('Taking away the cells a chart reads is not implemented')
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
        if field == 'value_reference':
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
                    if turned != reference:
                        steps.append(setting(bar, field, (turned, values)))  # pyright: ignore[reportArgumentType]
    return steps


def carry_out(steps: list[Step]) -> None:
    for step in steps:
        step()
