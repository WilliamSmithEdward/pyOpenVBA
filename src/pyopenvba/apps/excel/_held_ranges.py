"""A Range follows its cells as a reference in a formula does, however it is held, and goes once they have gone: then
anything but its type raises 424 (scripts/measure_range_tracking.py, tests/fixtures/range_tracking.json,
tests/fixtures/range_holding.json)."""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area, split_sheet

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def _spelled(area: Area) -> str:
    return '=' + area.address(absolute=False)


def _follow(held: Range, rewrite: Callable[[str], str], target: Worksheet | None = None) -> None:
    from pyopenvba.apps.excel._model import GoneRange

    areas: list[Area] = []
    moved_to = held.sheet
    for area in held.areas:
        text = rewrite(_spelled(area))
        if '#REF!' in text:
            continue
        named, _ = split_sheet(text.removeprefix('='))
        if named and target is not None:
            moved_to = target
        areas.append(parse_area(text.removeprefix('='), sheet=moved_to.name))
    if not areas:
        held.__class__ = GoneRange
        return
    if moved_to is not held.sheet:
        held.sheet.held_ranges.discard(held)
        moved_to.held_ranges.add(held)
        held.sheet = moved_to
    held.areas = areas


def moved(sheet: Worksheet, rewrite: Callable[[str], str], target: Worksheet | None = None) -> None:
    """The ranges on a sheet, once its cells have moved as ``rewrite`` moves a reference. The selection and the active cell
    stay where they are, as places on the sheet, though a Range that was selected follows its cells: they go on as copies
    (tests/fixtures/range_holding.json)."""
    from pyopenvba.apps.excel._model import Range

    def kept(found: Range | None) -> Range | None:
        return None if found is None else Range(sheet, list(found.areas), whole=found.whole, lines=found.lines)

    sheet.cells_moved += 1
    selected, active = kept(sheet.selection_range), kept(sheet.active_cell_range)
    for held in list(sheet.held_ranges):
        if held is not selected and held is not active:
            _follow(held, rewrite, target)
    sheet.selection_range, sheet.active_cell_range = selected, active
