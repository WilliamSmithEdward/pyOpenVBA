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
        # A merged cell's font does not make its rows taller; unmerged, it can.
        sheet.dims.invalidate_growth()
        sheet.touched()


def _mapped(area: Area, *, rows: bool, start: int, count: int, delete: bool) -> Area | None:
    """A merged area once lines are inserted or deleted across it, as a reference to it would follow; None when
    no more than one of its cells is left, which is no merge at all (tests/fixtures/merge_structure_edits.json)."""
    from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS
    from pyopenvba.apps.excel._editing import interval

    low, high = (area.top, area.bottom) if rows else (area.left, area.right)
    result = interval(low, high, start, count, delete, MAX_ROWS if rows else MAX_COLUMNS)
    if result is None:
        return None
    moved = (Area(result[0], area.left, result[1], area.right, area.sheet) if rows
             else Area(area.top, result[0], area.bottom, result[1], area.sheet))
    return moved if moved.rows * moved.columns > 1 else None


def lines_planned(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> list[Area]:
    """The sheet's merged areas after whole rows or columns go in or out: each moves, grows or shrinks with them
    (tests/fixtures/merge_structure_edits.json)."""
    return [moved for moved in (_mapped(area, rows=rows, start=start, count=count, delete=delete)
                                for area in sheet.merged_areas) if moved is not None]


def shift_planned(sheet: Worksheet, band: tuple[int, int], *, vertical: bool, start: int, count: int,
                  delete: bool) -> list[Area]:
    """The sheet's merged areas after cells in a band shift: one wholly inside the band moves as with whole lines,
    and one the moving cells cut through is unmerged, its value staying with its cell; Excel asks first unless
    DisplayAlerts is off (tests/fixtures/merge_structure_edits.json)."""
    planned: list[Area] = []
    for area in sheet.merged_areas:
        low, high = (area.left, area.right) if vertical else (area.top, area.bottom)
        reach = area.bottom if vertical else area.right
        if high < band[0] or low > band[1] or reach < start:
            planned.append(area)
        elif band[0] <= low and high <= band[1]:
            moved = _mapped(area, rows=vertical, start=start, count=count, delete=delete)
            if moved is not None:
                planned.append(moved)
        elif sheet.book.application.display_alerts:
            raise VBAUnsupportedError("A cell shift that unmerges merged cells with DisplayAlerts on, where Excel "
                                      "asks first, is not implemented")
    return planned


def replaced(sheet: Worksheet, planned: list[Area]) -> None:
    """The sheet's merged areas set to what an edit planned."""
    if planned != sheet.merged_areas:
        sheet.merged_areas = planned
        sheet.merges_dirty = True
        sheet.dims.invalidate_growth()
        sheet.touched()


def merge(target: Range, across: bool) -> None:
    if len(target.areas) != 1:
        raise VBAUnsupportedError("Merge on multiple areas is not implemented")
    sheet, area = target.sheet, target.first
    if across and any(one.rows > 1 and intersects(one, area) for one in sheet.merged_areas):
        raise VBAUnsupportedError("Across merges over existing multi-row merges are not implemented")
    if sheet.data_tables:
        from pyopenvba.apps.excel._data_tables import refuse

        refuse(sheet, [area], "Merging", whole=True)
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
        if sheet.notes:
            from pyopenvba.apps.excel._notes import merged

            merged(sheet, wanted)
        sheet.merged_areas = [one for one in sheet.merged_areas if not intersects(one, wanted)]
        sheet.merged_areas.append(wanted)
        sheet.merges_dirty = True
        sheet.dims.invalidate_growth()
        sheet.touched()
