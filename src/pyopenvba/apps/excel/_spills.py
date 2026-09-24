"""Dynamic arrays: a formula Formula2 writes that answers with an array spills into the cells below and beside it.

Measured in live Excel (scripts/measure_dynamic_arrays.py,
tests/fixtures/dynamic_arrays.json). The formula's first cell holds it
and shows the answer's first item; each other cell of the block shows
its item with no formula of its own, Formula empty and HasFormula False,
and follows the formula as its answer changes, growing and shrinking
with it; an answer of one item spills nowhere. A cell in the way -- one
holding anything, one of a merged block, another formula's spill, or
the sheet's edge -- makes the formula #SPILL!, until it is out of the
way. Clearing a cell the formula spilled into, or writing an empty
string there, changes nothing; writing anything else there puts a cell
in the way. HasSpill is True for a range meeting a spill, SpillParent
is a spilled cell's formula's cell, SpillingToRange the formula's cell's
block, and A1# reads what spilled from A1.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel._engine_book import model_value
from pyopenvba.formula._calc.values import Array
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, NOTHING

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

#: What a formula whose answer cannot spill shows.
SPILL = ExcelError("#SPILL!")


#: Why a formula's answer cannot spill, as a file records it in the rich value behind #SPILL!: a cell in the way,
#: the sheet's edge, or a merged block (tests/fixtures/dynamic_arrays.json).
OBSTRUCTED = 1
BEYOND_THE_EDGE = 3
MERGED = 6


@dataclass
class Spill:
    """What a dynamic-array formula's answer wants: the block, its first cell the formula's, and whether a cell in
    the way keeps it from spilling there, and why."""

    area: Area
    blocked: bool = False
    why: int = OBSTRUCTED


def spilled(sheet: Worksheet, row: int, column: int) -> tuple[tuple[int, int], Area] | None:
    """The formula's first cell and the block of the spill a cell is part of, where an answer spilled over it."""
    for anchor, spill in sheet.spills.items():
        if not spill.blocked and spill.area.contains(row, column):
            return anchor, spill.area
    return None


def wanted(sheet: Worksheet, row: int, column: int) -> list[tuple[int, int]]:
    """The first cells of the formulas whose answers want a cell, spilled there or kept out by what is in the way."""
    return [anchor for anchor, spill in sheet.spills.items()
            if spill.area.contains(row, column) and anchor != (row, column)]


def place(sheet: Worksheet, anchor: tuple[int, int], answer: Array) -> tuple[object, set[tuple[int, int]]]:
    """Put a dynamic-array formula's answer in place: its first item in the formula's cell and the rest spilled
    beside it, or #SPILL! where something is in the way. The value for the formula's cell, and the cells whose value
    changed."""
    row, column = anchor
    changed = remove(sheet, anchor)
    if answer.height == 1 and answer.width == 1:
        return model_value(answer.rows[0][0]), changed
    area = Area(row, column, row + answer.height - 1, column + answer.width - 1)
    why = BEYOND_THE_EDGE if area.bottom > MAX_ROWS or area.right > MAX_COLUMNS else _in_the_way(sheet, anchor, area)
    if why:
        sheet.spills[anchor] = Spill(area, blocked=True, why=why)
        return SPILL, changed
    sheet.spills[anchor] = Spill(area)
    for each_row in range(area.top, area.bottom + 1):
        for each_column in range(area.left, area.right + 1):
            if (each_row, each_column) == anchor:
                continue
            cell = sheet.cell(each_row, each_column, create=True)
            assert cell is not None
            value = model_value(answer.rows[each_row - row][each_column - column])
            if cell.value != value or cell.spilled_from != anchor:
                changed.add((each_row, each_column))
            cell.value, cell.stale, cell.spilled_from = value, False, anchor
    return model_value(answer.rows[0][0]), changed


def remove(sheet: Worksheet, anchor: tuple[int, int]) -> set[tuple[int, int]]:
    """Take a formula's answer back out of the cells it spilled into, and forget what it wanted; the cells emptied."""
    spill = sheet.spills.pop(anchor, None)
    emptied: set[tuple[int, int]] = set()
    if spill is None or spill.blocked:
        return emptied
    for row in range(spill.area.top, spill.area.bottom + 1):
        for column in range(spill.area.left, spill.area.right + 1):
            cell = sheet.cells_.get((row, column))
            if cell is not None and cell.spilled_from == anchor:
                cell.value, cell.spilled_from = EMPTY, None
                sheet.settle(row, column)
                emptied.add((row, column))
    return emptied


def reconcile(sheet: Worksheet) -> None:
    """After cells moved, as a cut or an inserted row moves them: forget each spill whose first cell holds no
    dynamic-array formula now, with what it spilled, and empty each cell showing an answer no formula spills there;
    a dynamic-array formula that moved spills again from where it is."""
    for anchor in list(sheet.spills):
        cell = sheet.cells_.get(anchor)
        if cell is None or not cell.dynamic or not cell.formula:
            remove(sheet, anchor)
    for position, cell in list(sheet.cells_.items()):
        if cell.spilled_from is not None and cell.spilled_from not in sheet.spills:
            cell.value, cell.spilled_from = EMPTY, None
            sheet.settle(*position)


def _in_the_way(sheet: Worksheet, anchor: tuple[int, int], area: Area) -> int:
    """What stands in a block besides the formula's own spill, as a #SPILL! records it: a merged block, or a cell
    holding something, an array formula or another formula's spill; 0 for nothing."""
    if any(_meets(one, area) for one in sheet.merged_areas):
        return MERGED
    for (row, column), cell in sheet.cells_.items():
        if (row, column) == anchor or not area.contains(row, column):
            continue
        if cell.spilled_from not in (None, anchor) or cell.spilled_from is None and (cell.value is not EMPTY
                                                                                      or cell.formula):
            return OBSTRUCTED
    others = [*sheet.array_formulas.values(),
              *(spill.area for other, spill in sheet.spills.items() if other != anchor and not spill.blocked)]
    return OBSTRUCTED if any(_meets(one, area) for one in others) else 0


def _meets(one: Area, other: Area) -> bool:
    return one.top <= other.bottom and other.top <= one.bottom and one.left <= other.right and other.left <= one.right


def keeps(sheet: Worksheet, row: int, column: int, value: object) -> bool:
    """Whether writing ``value`` to a cell leaves it as it is: an empty string or nothing written where an answer
    spilled."""
    cell = sheet.cells_.get((row, column))
    return cell is not None and cell.spilled_from is not None and (value is EMPTY or value == "")


def has_spill(target: Range) -> object:
    """Whether a range meets a spill: its formula's cell, while it spills, or a cell it spilled into."""
    sheet = target.sheet
    for spill in sheet.spills.values():
        if spill.blocked:
            continue
        if any(area.top <= spill.area.bottom and spill.area.top <= area.bottom and area.left <= spill.area.right
               and spill.area.left <= area.right for area in target.areas):
            return True
    return False


def spill_parent(target: Range) -> object:
    """The formula's cell whose answer spilled over a range's first cell, or Nothing."""
    from pyopenvba.apps.excel._model import Range

    first = target.first
    found = spilled(target.sheet, first.top, first.left)
    if found is None:
        return NOTHING
    (row, column), _ = found
    return Range(target.sheet, [Area(row, column, row, column, target.sheet.name)])


def spilling_to(target: Range) -> object:
    """The whole block a formula in a range's first cell spilled into, or Nothing for any other cell."""
    from pyopenvba.apps.excel._model import Range

    first = target.first
    spill = target.sheet.spills.get((first.top, first.left))
    if spill is None or spill.blocked:
        return NOTHING
    area = spill.area
    return Range(target.sheet, [Area(area.top, area.left, area.bottom, area.right, target.sheet.name)])


__all__ = ["BEYOND_THE_EDGE", "MERGED", "OBSTRUCTED", "SPILL", "Spill", "has_spill", "keeps", "place", "reconcile",
           "remove", "spill_parent", "spilled", "spilling_to", "wanted"]
