"""Shapes following the cells they are placed over when column widths and row heights change, as Excel moves them
(tests/fixtures/shape_cell_moves.json).

A shape that moves and sizes with its cells keeps the cells under its corners and how far into each its corners
are; one that only moves keeps its top left cell and that offset, and its size; a free floating one stays put. An
offset the cell no longer has room for shrinks to the cell's width or height.
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.shapes._values import Shape

Sizes = Callable[[int], float]


@dataclass(frozen=True)
class Anchor:
    """Where a shape's corners are: each one's column and row, and how far into them, in points."""

    column: int
    across: float
    row: int
    down: float
    right_column: int
    right_across: float
    bottom_row: int
    bottom_down: float


def _located(sizes: Sizes, at: float, last: int) -> tuple[int, float]:
    """The line a distance falls in, and how far into it: a point on a boundary is in the line after it."""
    index, start = 1, 0.0
    while index < last and start + sizes(index) <= at + 1e-9:
        start += sizes(index)
        index += 1
    return index, at - start


def _start(sizes: Sizes, index: int) -> float:
    return sum(sizes(line) for line in range(1, index))


def anchors(sheet: Worksheet) -> list[tuple[Shape, Anchor]]:
    """The anchors of the sheet's shapes that follow their cells."""
    widths, heights = sheet.column_width_points, sheet.row_height_points
    held: list[tuple[Shape, Anchor]] = []
    for shape in sheet.shapes_:
        if shape.edit_as == 'absolute':
            continue
        column, across = _located(widths, shape.left, MAX_COLUMNS)
        row, down = _located(heights, shape.top, MAX_ROWS)
        right_column, right_across = _located(widths, shape.left + shape.width, MAX_COLUMNS)
        bottom_row, bottom_down = _located(heights, shape.top + shape.height, MAX_ROWS)
        held.append((shape, Anchor(column, across, row, down, right_column, right_across, bottom_row, bottom_down)))
    return held


def placed(sheet: Worksheet, held: list[tuple[Shape, Anchor]]) -> None:
    """Each shape put back over the cells it was anchored to."""
    from pyopenvba.apps.excel._shapes import ShapeObject

    widths, heights = sheet.column_width_points, sheet.row_height_points

    def edge(sizes: Sizes, index: int, offset: float) -> float:
        return _start(sizes, index) + min(offset, sizes(index))

    for shape, anchor in held:
        left = edge(widths, anchor.column, anchor.across)
        top = edge(heights, anchor.row, anchor.down)
        width, height = shape.width, shape.height
        if shape.edit_as == '':
            width = edge(widths, anchor.right_column, anchor.right_across) - left
            height = edge(heights, anchor.bottom_row, anchor.bottom_down) - top
        wanted = {'left': left, 'top': top, 'width': width, 'height': height}
        sheet.dims.kept_anchors[id(shape)] = (anchor, (left, top, width, height))
        changes = {name: value for name, value in wanted.items() if abs(getattr(shape, name) - value) > 1e-9}
        if not changes:
            continue
        if shape.source:
            # A loaded shape's anchor in the file is held by cells already: only where it reads changes.
            for name, value in changes.items():
                setattr(shape, name, value)
        else:
            ShapeObject(sheet, shape).placed_by_cells(**changes)


def anchored_cell(sheet: Worksheet, shape: Shape, corner: str) -> tuple[int, int] | None:
    """The column and row of the cell under a shape's corner, "top left" or "bottom right", as the anchor it was last
    placed by holds it, which may be hidden; None once the shape has moved since, or if it never was."""
    kept = sheet.dims.kept_anchors.get(id(shape))
    if kept is None:
        return None
    anchor, box = kept
    if any(abs(held - now) > 1e-9 for held, now in zip(box, (shape.left, shape.top, shape.width, shape.height),
                                                          strict=True)):
        return None
    if corner == 'top left':
        return anchor.column, anchor.row
    # Only a shape that sizes with its cells keeps its bottom right cell; one that only moves is where it is.
    return (anchor.right_column, anchor.bottom_row) if shape.edit_as == '' else None


__all__ = ['Anchor', 'anchored_cell', 'anchors', 'placed']
