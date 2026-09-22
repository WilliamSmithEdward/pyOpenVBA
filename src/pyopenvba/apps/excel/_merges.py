"""Merged-cell geometry and measured value retention."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def intersects(a: Area, b: Area) -> bool:
    return a.top <= b.bottom and b.top <= a.bottom and a.left <= b.right and b.left <= a.right


def contains(a: Area, b: Area) -> bool:
    return a.contains(b.top, b.left) and a.contains(b.bottom, b.right)


def at(sheet: Worksheet, row: int, column: int) -> Area | None:
    return next((area for area in sheet.merged_areas if area.contains(row, column)), None)


def writable(sheet: Worksheet, row: int, column: int) -> bool:
    area = at(sheet, row, column)
    return area is None or (row, column) == (area.top, area.left)


def flag(target: Range) -> object:
    if len(target.areas) != 1:
        raise VBAUnsupportedError("MergeCells on multiple areas is not implemented")
    overlaps = [area for area in target.sheet.merged_areas if intersects(area, target.first)]
    if not overlaps:
        return False
    return True if len(overlaps) == 1 and contains(overlaps[0], target.first) else NULL


def validate_clear(target: Range) -> None:
    for merged in target.sheet.merged_areas:
        if any(intersects(merged, area) for area in target.areas):
            if not any(contains(area, merged) for area in target.areas):
                raise error(1004, "Cannot change part of a merged cell")


def unmerge(target: Range) -> None:
    sheet = target.sheet
    kept = [merged for merged in sheet.merged_areas
            if not any(intersects(merged, area) for area in target.areas)]
    if kept != sheet.merged_areas:
        sheet.merged_areas = kept
        sheet.merges_dirty = True
        sheet.touched()


def merge(target: Range, across: bool) -> None:
    if len(target.areas) != 1:
        raise VBAUnsupportedError("Merge on multiple areas is not implemented")
    sheet, area = target.sheet, target.first
    if across and any(one.rows > 1 and intersects(one, area) for one in sheet.merged_areas):
        raise VBAUnsupportedError("Across merges over existing multi-row merges are not implemented")
    areas = (Area(row, area.left, row, area.right, sheet.name)
             for row in range(area.top, area.bottom + 1)) if across else iter([area])
    for wanted in areas:
        while True:
            overlapping = [one for one in sheet.merged_areas if intersects(one, wanted)]
            expanded = Area(min([wanted.top] + [one.top for one in overlapping]),
                            min([wanted.left] + [one.left for one in overlapping]),
                            max([wanted.bottom] + [one.bottom for one in overlapping]),
                            max([wanted.right] + [one.right for one in overlapping]), sheet.name)
            if expanded == wanted:
                break
            wanted = expanded
        if wanted.rows * wanted.columns == 1:
            continue
        entries = sorted((pos, cell) for pos, cell in sheet.cells_.items() if wanted.contains(*pos))
        first = next((cell for _, cell in entries if cell.formula or cell.value is not EMPTY), None)
        retained = (first.value, first.formula, first.stale) if first is not None else (EMPTY, "", False)
        for (row, column), cell in entries:
            cell.value, cell.formula, cell.stale = EMPTY, "", False
            sheet.cell_changed(row, column)
        if first is not None:
            anchor = sheet.cell(wanted.top, wanted.left, create=True)
            assert anchor is not None
            anchor.value, anchor.formula, anchor.stale = retained
            sheet.cell_changed(wanted.top, wanted.left)
        sheet.merged_areas = [one for one in sheet.merged_areas if not intersects(one, wanted)]
        sheet.merged_areas.append(wanted)
        sheet.merges_dirty = True
        sheet.touched()
