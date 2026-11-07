"""Shapes following the cells they are placed over when column widths and row heights change, as Excel moves them
(tests/fixtures/shape_cell_moves.json).

A shape that moves and sizes with its cells keeps the cells under its corners and how far into each its corners
are; one that only moves keeps its top left cell and that offset, and its size; a free floating one stays put. An
offset the cell no longer has room for shrinks to the cell's width or height.
"""
from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.shapes._values import Shape

Sizes = Callable[[int], float]
#: Excel took a place of 100000 points as it was.
FAR_ENOUGH = 100000.0


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


class RowSizes:
    """The heights of a sheet's rows, gone through by runs: a sheet has a million rows and few of them differ, and a shape
    near the bottom must not cost a million lookups. The answers are those of a walk row by row, as a row is a whole number
    of pixels of 0.75 points, which adds up exactly."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        sheet.dims.settle_shapes()
        self.own = sheet.dims.rows_of_their_own()
        self.plain = sheet.row_height_points(MAX_ROWS + 1)

    def __call__(self, row: int) -> float:
        return self.sheet.row_height_points(row)

    def runs(self, last: int) -> Iterator[tuple[int, int, float]]:
        """First row, how many, and the height of each, up to but not including ``last``."""
        index = 1
        for row in self.own:
            if row >= last:
                break
            if row > index:
                yield index, row - index, self.plain
            yield row, 1, self(row)
            index = row + 1
        if index < last:
            yield index, last - index, self.plain

    def located(self, at: float, last: int) -> tuple[int, float]:
        start = 0.0
        for first, count, height in self.runs(last):
            if height <= 0:
                continue
            passed = min(count, int((at + 1e-9 - start) // height)) if start + height <= at + 1e-9 else 0
            if passed < count:
                return first + passed, at - (start + passed * height)
            start += count * height
        return last, at - start

    def start(self, index: int) -> float:
        return sum(count * height for _, count, height in self.runs(index))


def keep_on_sheet(sheet: Worksheet, shape: Shape) -> None:
    """A place past the sheet's end stops where the shape still fits on it, as Excel has it (tests/fixtures/excel_huge_numbers.json). Only a place too far for a
    sheet of ordinary rows and columns goes through the sheet to find out; a sheet of narrow ones keeps what it was given, as ever."""
    if shape.top > FAR_ENOUGH:
        shape.top = max(min(shape.top, RowSizes(sheet).start(MAX_ROWS + 1) - shape.height), 0.0)
    if shape.left > FAR_ENOUGH:
        shape.left = max(min(shape.left, sum(sheet.column_width_points(column) for column in range(1, MAX_COLUMNS + 1)) - shape.width), 0.0)


def _located(sizes: Sizes, at: float, last: int) -> tuple[int, float]:
    """The line a distance falls in, and how far into it: a point on a boundary is in the line after it."""
    if isinstance(sizes, RowSizes):
        return sizes.located(at, last)
    index, start = 1, 0.0
    while index < last and start + sizes(index) <= at + 1e-9:
        start += sizes(index)
        index += 1
    return index, at - start


def _start(sizes: Sizes, index: int) -> float:
    if isinstance(sizes, RowSizes):
        return sizes.start(index)
    return sum(sizes(line) for line in range(1, index))


def anchors(sheet: Worksheet) -> list[tuple[Shape, Anchor]]:
    """The anchors of the sheet's shapes that follow their cells."""
    sheet.dims.settle_shapes()
    widths, heights = sheet.column_width_points, RowSizes(sheet)
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


def placed(sheet: Worksheet, held: list[tuple[Shape, Anchor]], *, rewrite: bool = False) -> None:
    """Each shape put back over the cells it was anchored to. ``rewrite`` says the cells themselves moved, so that a
    loaded shape's anchor in the file no longer holds and is written again."""
    sheet.dims.placing = True
    try:
        _place(sheet, held, rewrite)
    finally:
        sheet.dims.placing = False


def _place(sheet: Worksheet, held: list[tuple[Shape, Anchor]], rewrite: bool) -> None:
    from pyopenvba.apps.excel._shapes import ShapeObject

    widths, heights = sheet.column_width_points, RowSizes(sheet)

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
        if rewrite and shape.source:
            ShapeObject(sheet, shape).placed_by_cells(**wanted)
            continue
        if not changes:
            continue
        if shape.source:
            # A loaded shape's anchor in the file is held by cells already: only where it reads changes.
            for name, value in changes.items():
                setattr(shape, name, value)
        else:
            ShapeObject(sheet, shape).placed_by_cells(**changes)


def remembered(sheet: Worksheet) -> None:
    """Each shape that follows its cells remembered with where it is anchored now."""
    if sheet.dims.placing:
        return
    kept: dict[int, tuple[Anchor, tuple[float, float, float, float]]] = {}
    for shape, anchor in anchors(sheet):
        kept[id(shape)] = (anchor, (shape.left, shape.top, shape.width, shape.height))
    sheet.dims.kept_anchors = kept


def replaced(sheet: Worksheet) -> None:
    """After row heights are worked out again, each shape still where its remembered anchor put it, put back by
    that anchor: rows grown by their fonts move the shapes over them (tests/fixtures/shape_structure_moves.json)."""
    held: list[tuple[Shape, Anchor]] = []
    for shape in sheet.shapes_:
        kept = sheet.dims.kept_anchors.get(id(shape))
        if kept is None or shape.edit_as == 'absolute':
            continue
        anchor, box = kept
        if all(abs(was - now) <= 1e-9 for was, now in zip(box, (shape.left, shape.top, shape.width, shape.height),
                                                          strict=True)):
            held.append((shape, anchor))
    if held:
        placed(sheet, held)


def lines_edited(sheet: Worksheet, held: list[tuple[Shape, Anchor]], *, rows: bool, start: int, count: int,
                 delete: bool) -> None:
    """The shapes put back after whole rows or columns were inserted or deleted (tests/fixtures/
    shape_structure_moves.json). A corner in a deleted line goes to the start of the line after them; a shape that
    sizes with its cells and so has nothing left between its corners is deleted, as Excel deletes it."""
    from pyopenvba.apps.excel._editing import interval
    from pyopenvba.apps.excel._shapes import ShapeObject

    limit = MAX_ROWS if rows else MAX_COLUMNS

    def mapped(line: int, offset: float) -> tuple[int, float]:
        result = interval(line, line, start, count, delete, limit)
        return (start, 0.0) if result is None else (result[0], offset)

    moved: list[tuple[Shape, Anchor]] = []
    for shape, anchor in held:
        if rows:
            row, down = mapped(anchor.row, anchor.down)
            bottom_row, bottom_down = mapped(anchor.bottom_row, anchor.bottom_down)
            anchor = Anchor(anchor.column, anchor.across, row, down, anchor.right_column, anchor.right_across,
                            bottom_row, bottom_down)
        else:
            column, across = mapped(anchor.column, anchor.across)
            right_column, right_across = mapped(anchor.right_column, anchor.right_across)
            anchor = Anchor(column, across, anchor.row, anchor.down, right_column, right_across, anchor.bottom_row,
                            anchor.bottom_down)
        collapsed = ((anchor.row, anchor.down) == (anchor.bottom_row, anchor.bottom_down) if rows
                     else (anchor.column, anchor.across) == (anchor.right_column, anchor.right_across))
        if shape.edit_as == '' and collapsed:
            ShapeObject(sheet, shape).Delete()
            continue
        moved.append((shape, anchor))
    placed(sheet, moved, rewrite=True)


def anchored_cell(sheet: Worksheet, shape: Shape, corner: str) -> tuple[int, int] | None:
    """The column and row of the cell under a shape's corner, "top left" or "bottom right", as the anchor it was last
    placed by holds it, which may be hidden; None once the shape has moved since, or if it never was."""
    kept = sheet.dims.kept_anchors.get(id(shape))
    if kept is None or shape.edit_as == 'absolute':
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
