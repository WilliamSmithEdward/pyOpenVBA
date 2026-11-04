"""Range.Insert with something on the clipboard: Excel's Insert Copied Cells and Insert Cut Cells.

What Copy put on the clipboard goes in as new cells and stays on the clipboard. What Cut put there moves: the cells
open up where it lands, its cells go there, and when it came along the same rows or columns the gap it leaves closes.
Everything here was read from tests/fixtures/range_audit.json.

A copy of whole lines goes into whole lines of its own kind, as many as the target holds whole copies of, else one copy.
A copy of cells into whole rows fills them when its width divides the row's, and fails above where it came from; a
width that does not divide takes one copy at the left, as does a height that more than one row of target does not
hold. Into whole columns it fails before where it came from, and otherwise would fill columns of a million cells,
which is not implemented. Copied cells into cells open up the block the target tiles to, an axis at a time, down unless
taller than wide.

A cut lands at the target's corner, whole lines included. It fails where it would land on itself or tear itself apart,
with its cells still on the clipboard.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel._clipboard import Clip, conditional_copy_admitted, move, sized
from pyopenvba.apps.excel._protection import whole
from pyopenvba.apps.excel._visible import unmeasured
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

_FAILED = "Insert method of Range class failed"


def insert(target: Range, clip: Clip, shift: object, below: bool) -> Range:
    """The clipboard inserted at ``target``, answering the cells it landed on."""
    if clip.cut:
        return _cut_in(target, clip, shift)
    return _copied_in(target, clip, shift, below)


def _lines(sheet: Worksheet, start: int, count: int, rows: bool) -> Range:
    from pyopenvba.apps.excel._model import Range

    end = start + count - 1
    if rows:
        return Range(sheet, [Area(start, 1, end, MAX_COLUMNS, sheet.name)], whole="rows")
    return Range(sheet, [Area(1, start, MAX_ROWS, end, sheet.name)], whole="columns")


def _span(area: Area, rows: bool) -> tuple[int, int]:
    return (area.top, area.bottom) if rows else (area.left, area.right)


def _meets(one: Area, other: Area) -> bool:
    return not (one.bottom < other.top or other.bottom < one.top or one.right < other.left or other.right < one.left)


def _down(shift: object, rows: int, columns: int) -> bool:
    """Which way cells open up: as Shift says, or with none down unless the block is taller than it is wide."""
    return rows <= columns if shift is MISSING else int(to_integer(shift, "Long")) == -4121  # xlDown


def _copied_in(target: Range, clip: Clip, shift: object, below: bool) -> Range:
    from pyopenvba.apps.excel._editing import edit, shift_cells
    from pyopenvba.apps.excel._model import Range

    sheet, area, source = target.sheet, target.first, clip.source
    first = clip.block.parts[0][0]
    here = source.sheet is sheet
    if first.whole_rows or first.whole_columns:
        rows = first.whole_rows
        if not (area.whole_rows if rows else area.whole_columns):
            raise error(1004, _FAILED)
        (start, end), (at, last) = _span(first, rows), _span(area, rows)
        if here and start < at <= end:
            raise error(1004, _FAILED)
        size, wanted = end - start + 1, last - at + 1
        count = wanted if wanted % size == 0 else size
        _admitted(sheet, source, _lines(sheet, at, count, rows).first)
        edit(_lines(sheet, at, count, rows), delete=False, below=below)
        source.copy_to(_lines(sheet, at, count, rows))
        return _lines(sheet, at, count, rows)
    height, width = clip.block.rows, clip.block.columns
    if area.whole_columns:
        if here and area.left < first.left:
            raise error(1004, _FAILED)
        raise VBAUnsupportedError("inserting copied cells as whole columns is not implemented")
    if area.whole_rows:
        fills = MAX_COLUMNS % width == 0
        holds = area.rows % height == 0
        tiled = fills and (holds or area.rows == 1)
        if tiled and here and area.top < first.top:
            raise error(1004, _FAILED)
        count = area.rows if tiled and holds else height
        landed = _lines(sheet, area.top, count, True).first if tiled else Area(
            area.top, 1, area.top + height - 1, width, sheet.name)
        _admitted(sheet, source, landed)
        edit(_lines(sheet, area.top, count, True), delete=False, below=below)
        source.copy_to(Range(sheet, [landed]))
        return Range(sheet, [landed])
    rows = area.rows if area.rows % height == 0 else height
    columns = area.columns if area.columns % width == 0 else width
    landed = Area(area.top, area.left, area.top + rows - 1, area.left + columns - 1, sheet.name)
    _admitted(sheet, source, landed)
    target.shifting_cells(landed, vertical=_down(shift, rows, columns), delete=False)
    shift_cells(Range(sheet, [landed]), landed, delete=False, vertical=_down(shift, rows, columns), below=below)
    source.copy_to(Range(sheet, [landed]))
    return Range(sheet, [landed])


def _admitted(sheet: Worksheet, source: Range, landed: Area) -> None:
    """What would refuse the paste, refused before any cells open up for it."""
    from pyopenvba.apps.excel._model import Range

    unmeasured(Range(sheet, [landed]), "Inserting copied cells")
    whole(sheet, [landed], "Inserting copied cells")
    conditional_copy_admitted(source.sheet, list(source.areas), sheet, [landed])


def _cut_in(target: Range, clip: Clip, shift: object) -> Range:
    from pyopenvba.apps.excel._editing import edit, shift_cells
    from pyopenvba.apps.excel._merges import intersects
    from pyopenvba.apps.excel._model import Range

    sheet, source = target.sheet, clip.source
    area, first = target.first, clip.area
    here = source.sheet is sheet
    wide, tall = first.whole_rows, first.whole_columns
    if wide or tall:
        if (wide and area.left != 1) or (tall and area.top != 1):
            raise error(1004, _FAILED)
        size = first.rows if wide else first.columns
        at = area.top if wide else area.left
        start, end = _span(first, wide)
        landing = _lines(sheet, at, size, wide).first
        if here and start <= at <= end:
            raise error(1004, _FAILED)
        if sized(source.sheet, first) or sized(sheet, landing):
            raise VBAUnsupportedError("cutting whole rows or columns that carry sizes or formats of their own is not "
                                      "implemented")
        _refused(sheet, first, landing, source)
        edit(_lines(sheet, at, size, wide), delete=False)
        gap = source.first
        move(source, gap, Range(sheet, [landing]))
        if here:
            edit(_lines(sheet, gap.top if wide else gap.left, size, wide), delete=True)
        return source.place_now() if here else _lines(sheet, at, size, wide)
    landing = Area(area.top, area.left, area.top + first.rows - 1, area.left + first.columns - 1, sheet.name)
    same_rows = here and (first.top, first.bottom) == (landing.top, landing.bottom)
    same_columns = here and (first.left, first.right) == (landing.left, landing.right)
    down = same_columns if shift is MISSING and (same_rows or same_columns) else _down(shift, first.rows, first.columns)
    closes = same_columns if down else same_rows
    # On itself it may only start before itself, along the way the cells open up.
    before = landing.top < first.top if down else landing.left < first.left
    if here and _meets(landing, first) and not (closes and before):
        raise error(1004, _FAILED)
    if here and not closes:
        # Cells opening up would tear it apart: part of it in their way, the rest not.
        if down and first.top > landing.bottom and first.left <= landing.right and landing.left <= first.right:
            raise error(1004, _FAILED)
        if not down and first.left > landing.right and first.top <= landing.bottom and landing.top <= first.bottom:
            raise error(1004, _FAILED)
    _refused(sheet, first, landing, source)
    if any(intersects(first, one) for one in source.sheet.merged_areas):
        raise VBAUnsupportedError("cutting merged cells is not implemented")
    target.shifting_cells(landing, vertical=down, delete=False)
    shift_cells(Range(sheet, [landing]), landing, delete=False, vertical=down)
    gap = source.first
    move(source, gap, Range(sheet, [landing]))
    if closes:
        shift_cells(Range(sheet, [gap]), gap, delete=True, vertical=down)
    return source.place_now() if here else Range(sheet, [landing])


def _refused(sheet: Worksheet, first: Area, landing: Area, source: Range) -> None:
    """What would refuse the move, refused before any cells open up for it."""
    from pyopenvba.apps.excel._merges import intersects
    from pyopenvba.apps.excel._model import Range

    unmeasured(Range(sheet, [landing]), "Inserting cut cells")
    whole(source.sheet, [first], "Cut")
    whole(sheet, [landing], "Cut")
    conditional_copy_admitted(source.sheet, [first], sheet, [landing])
    if any(intersects(landing, one) for one in sheet.merged_areas):
        raise VBAUnsupportedError("cutting merged cells is not implemented")
    if any(_meets(first, table.area) for table in source.sheet.tables) or any(
            _meets(landing, table.area) for table in sheet.tables):
        raise VBAUnsupportedError("inserting cut cells at or from a table is not implemented")
