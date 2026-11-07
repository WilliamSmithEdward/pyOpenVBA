"""Merged-cell geometry and measured value retention."""
from __future__ import annotations

from collections.abc import Sequence
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


#: The most cells a merge gives its first cell's format to, all of one column.
MOST_FORMATTED = 1048576


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
    asked = False
    for area in sheet.merged_areas:
        low, high = (area.left, area.right) if vertical else (area.top, area.bottom)
        reach = area.bottom if vertical else area.right
        if high < band[0] or low > band[1] or reach < start:
            planned.append(area)
        elif band[0] <= low and high <= band[1]:
            moved = _mapped(area, rows=vertical, start=start, count=count, delete=delete)
            if moved is not None:
                planned.append(moved)
        elif sheet.book.application.display_alerts and not asked:
            asked = True
            _ask_to_unmerge(sheet, delete=delete)
    return planned


#: Excel's prompt before a shift unmerges merged cells (tests/fixtures/merge_shifts.json), two spaces and all.
UNMERGE_PROMPT = "This operation will cause some merged cells to unmerge.  Do you wish to continue?"


def _ask_to_unmerge(sheet: Worksheet, *, delete: bool) -> None:
    """Excel asks once, OK and Cancel with OK the default, before anything moves. OK goes on as with DisplayAlerts off;
    Cancel is 1004 and leaves the sheet as it was. The answer comes from answer_dialogs, as MsgBox's does: 1 for OK, 2 for
    Cancel, and with none queued the default."""
    from pyopenvba.exceptions import VBADialogError
    from pyopenvba.interpreter._values import to_integer

    interpreter = sheet.book.application.interpreter
    answer = 1 if interpreter is None else to_integer(interpreter.show_message(UNMERGE_PROMPT, 1, "Microsoft Excel"))
    if answer == 2:
        raise error(1004, ("Delete" if delete else "Insert") + " method of Range class failed")
    if answer != 1:
        raise VBADialogError(f"Excel's prompt to unmerge cells takes 1 for OK or 2 for Cancel, not {answer!r}")


def replaced(sheet: Worksheet, planned: list[Area]) -> None:
    """The sheet's merged areas set to what an edit planned."""
    if planned != sheet.merged_areas:
        sheet.merged_areas = planned
        sheet.merges_dirty = True
        sheet.dims.invalidate_growth()
        sheet.touched()


def _take_first_format(sheet: Worksheet, wanted: Area) -> None:
    """Every cell of a merge takes its first cell's format, and keeps it when unmerged. Of the borders only the ends of
    a single row or column stay, each cell's own (tests/fixtures/range_areas.json)."""
    from pyopenvba.apps.excel import _styles as S

    if wanted.rows * wanted.columns > MOST_FORMATTED:
        raise VBAUnsupportedError("merging more than a million cells is not implemented")
    first = sheet.style_at(wanted.top, wanted.left)
    changes: list[tuple[int, int, S.Style]] = []
    for row in range(wanted.top, wanted.bottom + 1):
        for column in range(wanted.left, wanted.right + 1):
            own = sheet.style_at(row, column).border
            one_row, one_column = wanted.rows == 1, wanted.columns == 1
            border = S.Border(left=own.left if one_row and column == wanted.left else S.NO_SIDE,
                              right=own.right if one_row and column == wanted.right else S.NO_SIDE,
                              top=own.top if one_column and row == wanted.top else S.NO_SIDE,
                              bottom=own.bottom if one_column and row == wanted.bottom else S.NO_SIDE)
            changes.append((row, column, S.applying(first, "border", border=border) if border != first.border else first))
    for row, column, style in changes:
        if style != sheet.style_at(row, column):
            cell = sheet.cell(row, column, create=True)
            assert cell is not None
            cell.style = style
            sheet.cell_changed(row, column)


def overlap(areas: Sequence[Area]) -> bool:
    """Whether any two of ``areas`` share a cell."""
    return any(intersects(one, other) for at_one, one in enumerate(areas) for other in areas[at_one + 1:])


def merge(target: Range, across: bool) -> None:
    if len(target.areas) != 1:
        # Each area alone, and areas that overlap are 1004 (tests/fixtures/range_areas.json).
        if overlap(target.areas):
            raise error(1004, "Merge method of Range class failed")
        from pyopenvba.apps.excel._model import Range

        for one in target.areas:
            merge(Range(target.sheet, [one]), across)
        return
    sheet, area = target.sheet, target.first
    if across and any(one.rows > 1 and intersects(one, area) for one in sheet.merged_areas):
        # A merge of several rows comes apart first, and each row then keeps its own first value.
        unmerge(target)
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
        _take_first_format(sheet, wanted)
        if sheet.notes:
            from pyopenvba.apps.excel._notes import merged

            merged(sheet, wanted)
        sheet.merged_areas = [one for one in sheet.merged_areas if not intersects(one, wanted)]
        sheet.merged_areas.append(wanted)
        sheet.merges_dirty = True
        sheet.dims.invalidate_growth()
        sheet.touched()
