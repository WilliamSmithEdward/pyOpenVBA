"""Blocks of data: which cells hold something, and the region around a cell that CurrentRegion answers.

Measured in live Excel (scripts/measure_current_region.py): 34 layouts,
eight of them seeded random grids read from every other cell.

- A cell holds something when it has a value or a formula: a formula that
  returns "", a space and a lone apostrophe count, a cell with only a
  format does not. End walks over format-only cells as it walks over
  empty ones.
- CurrentRegion starts from the top-left cell of the first area, however
  large the range, and grows a row or a column at a time while the ring
  of cells around it, corners included, holds something; so cells that
  touch only at a corner are one region, and an empty starting cell is
  kept in it. A merged area counts as holding something when its first
  cell does, and the region takes in any merged area it reaches.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel import _merges
from pyopenvba.interpreter._values import EMPTY

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def holds_content(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether a cell has a value or a formula, which a format alone does not give it."""
    cell = sheet.cells_.get((row, column))
    return cell is not None and (cell.value is not EMPTY or bool(cell.formula))


def _occupied(sheet: Worksheet, row: int, column: int) -> bool:
    if holds_content(sheet, row, column):
        return True
    merged = _merges.at(sheet, row, column)
    return merged is not None and holds_content(sheet, merged.top, merged.left)


def _any(sheet: Worksheet, top: int, left: int, bottom: int, right: int) -> bool:
    return any(_occupied(sheet, row, column) for row in range(top, bottom + 1) for column in range(left, right + 1))


def _with_merges(sheet: Worksheet, box: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    """A box grown until it cuts no merged area in two."""
    top, left, bottom, right = box
    grown = True
    while grown:
        grown = False
        for area in sheet.merged_areas:
            if area.top <= bottom and area.bottom >= top and area.left <= right and area.right >= left:
                wider = (min(top, area.top), min(left, area.left), max(bottom, area.bottom), max(right, area.right))
                if wider != (top, left, bottom, right):
                    top, left, bottom, right = wider
                    grown = True
    return top, left, bottom, right


def current_region(sheet: Worksheet, row: int, column: int) -> Area:
    """The region CurrentRegion answers for the cell at ``row`` and ``column``."""
    box = (row, column, row, column)
    while True:
        top, left, bottom, right = before = _with_merges(sheet, box)
        # The ring around the box, one cell out on every side and its corners with it; each side
        # runs on while it finds something, so a long block is crossed in one pass.
        while top > 1 and _any(sheet, top - 1, max(left - 1, 1), top - 1, min(right + 1, MAX_COLUMNS)):
            top -= 1
        while bottom < MAX_ROWS and _any(sheet, bottom + 1, max(left - 1, 1), bottom + 1, min(right + 1, MAX_COLUMNS)):
            bottom += 1
        while left > 1 and _any(sheet, max(top - 1, 1), left - 1, min(bottom + 1, MAX_ROWS), left - 1):
            left -= 1
        while right < MAX_COLUMNS and _any(sheet, max(top - 1, 1), right + 1, min(bottom + 1, MAX_ROWS), right + 1):
            right += 1
        box = (top, left, bottom, right)
        if box == before:
            return Area(top, left, bottom, right, sheet.name)
