"""Pivot tables under inserts and deletes of rows, columns and cells, as Excel moves them (scripts/measure_pivot_shifts.py,
tests/fixtures/pivot_shifts_*.json).

A table, its page fields included, moves whole when every cell of it lies in the part of the sheet that shifts: rows or
columns going in at its first line or before it, or out before it, and cells going in or out above it or to its left
across a band that spans all its columns, or rows. One the shift would cut, a band reaching part of it, rows or columns
going in or out through it, or cells going in or out in the blank row between its page fields and itself, is Excel's
1004 and nothing changes. A delete that takes all of it takes the table (Range._taking_whole_pivot_tables). A cell
insert or delete inside the table is an edit of the report itself in Excel -- deleting an item's label hides the item,
inserting there adds a calculated item -- which the model does not make.

A cache whose source list lies on the edited sheet follows it as a reference would; the table keeps its values until
it is refreshed. Held PivotTable objects keep reading the table, which moves in place.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._editing import CellShift
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

#: Excel's words for an insert or delete that would cut a pivot table (tests/fixtures/pivot_shifts_a.json).
CUT = ("We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to "
       "change the report. If you are trying to insert or delete cells, move the PivotTable and try again.")


def _boxes(table: PivotTable) -> tuple[Area, Area | None]:
    """The table's own cells, TableRange1, and its page fields' cells, with the blank row under them left out."""
    location = table.root.find(NS + 'location')
    assert location is not None
    body = parse_area(location.get('ref', 'A1'))
    pages = int(location.get('rowPageCount', '0'))
    columns = int(location.get('colPageCount', '0'))
    if not table.root.findall(f'{NS}pageFields/{NS}pageField') or pages < 1 or columns < 1:
        return body, None
    return body, Area(body.top - pages - 1, body.left, body.top - 2, max(body.right, body.left + 3 * columns - 2))


def _meets(one: Area, other: Area) -> bool:
    return one.top <= other.bottom and other.top <= one.bottom and one.left <= other.right and other.left <= one.right


def _whole(body: Area, pages: Area | None) -> Area:
    if pages is None:
        return body
    return Area(pages.top, min(body.left, pages.left), body.bottom, max(body.right, pages.right))


def _moved(table: PivotTable, rows: bool, by: int) -> Callable[[], None]:
    def move() -> None:
        location = table.root.find(NS + 'location')
        assert location is not None
        area = parse_area(location.get('ref', 'A1'))
        area = (Area(area.top + by, area.left, area.bottom + by, area.right) if rows
                else Area(area.top, area.left + by, area.bottom, area.right + by))
        location.set('ref', area.address(absolute=False))
        table.changed = True
        table.sheet.book.saved = False
    return move


def lines_planned(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> Callable[[], None]:
    """Whole rows (or columns) ``start`` on going in or out: each table before them stays, each after them moves, and
    one they run through is 1004. What carries the moves out is returned; nothing changes until it is called."""
    from pyopenvba.apps.excel._editing import interval

    steps: list[Callable[[], None]] = []
    end = start + count - 1
    for table in list(getattr(sheet, 'pivot_tables', [])):
        box = _whole(*_boxes(table))
        low, high = (box.top, box.bottom) if rows else (box.left, box.right)
        if high < start:
            continue
        if delete:
            if start <= low and high <= end:
                # Range.Delete took it already; anything else deleting all of it is not measured.
                raise VBAUnsupportedError('Deleting every line of a pivot table this way is not implemented')
            if low <= end:
                raise error(1004, CUT)
            steps.append(_moved(table, rows, -count))
        else:
            if low < start:
                raise error(1004, CUT)
            steps.append(_moved(table, rows, count))
    sources = _sources_planned(sheet, lambda box: interval(box[0], box[1], start, count, delete,
                                                          MAX_ROWS if rows else MAX_COLUMNS), rows)
    return _carried(steps + sources)


def cells_planned(sheet: Worksheet, area: Area, shift: CellShift) -> Callable[[], None]:
    """``area`` going in or out with ``shift``: as lines_planned, each table across the band after it moving whole, one
    the band reaches part of being 1004, and one ``area`` touches being an edit of the report, which is refused."""
    from pyopenvba.apps.excel._editing import shifted_box

    steps: list[Callable[[], None]] = []
    vertical, band = shift.vertical, shift.band
    for table in list(getattr(sheet, 'pivot_tables', [])):
        body, pages = _boxes(table)
        box = _whole(body, pages)
        if _meets(area, box):
            if shift.delete and area.top <= box.top and area.left <= box.left and box.bottom <= area.bottom and \
                    box.right <= area.right:
                raise VBAUnsupportedError('Deleting all of a pivot table this way is not implemented')
            if _meets(area, body) or (pages is not None and _meets(area, pages)):
                raise VBAUnsupportedError('Inserting or deleting cells inside a pivot table, which edits the report '
                                          'in Excel, is not implemented')
            raise error(1004, CUT)
        across = (box.left, box.right) if vertical else (box.top, box.bottom)
        along = box.top if vertical else box.left
        if across[1] < band[0] or across[0] > band[1] or along < shift.start:
            continue
        if not (band[0] <= across[0] and across[1] <= band[1]):
            raise error(1004, CUT)
        steps.append(_moved(table, vertical, -shift.count if shift.delete else shift.count))
    sources: list[Callable[[], None]] = []
    for cache, origin, ref in _local_sources(sheet):
        moved = shifted_box((ref.top, ref.left, ref.bottom, ref.right), shift)
        if moved is None:
            raise VBAUnsupportedError("A pivot cache's source list deleted is not implemented")
        sources.append(_source_moved(cache, origin, Area(*moved)))
    return _carried(steps + sources)


def _local_sources(sheet: Worksheet) -> list[tuple[PivotCache, XmlElement, Area]]:
    caches: list[PivotCache] = []
    for cache in sheet.book.pivot_caches:
        if cache not in caches:
            caches.append(cache)
    for owner in sheet.book.sheets_:
        for table in getattr(owner, 'pivot_tables', []):
            if table.cache not in caches:
                caches.append(table.cache)
    found: list[tuple[PivotCache, XmlElement, Area]] = []
    for cache in caches:
        origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
        if origin is None or origin.get('name') or origin.get('sheet', '').casefold() != sheet.name.casefold():
            continue
        try:
            ref = parse_area(origin.get('ref', ''))
        except ValueError:
            raise VBAUnsupportedError("Moving a pivot cache's source that is not a block of cells is not implemented") \
                from None
        found.append((cache, origin, ref))
    return found


def _sources_planned(sheet: Worksheet, move: Callable[[tuple[int, int]], tuple[int, int] | None],
                     rows: bool) -> list[Callable[[], None]]:
    steps: list[Callable[[], None]] = []
    for cache, origin, ref in _local_sources(sheet):
        lines = move((ref.top, ref.bottom) if rows else (ref.left, ref.right))
        if lines is None:
            raise VBAUnsupportedError("A pivot cache's source list deleted is not implemented")
        area = Area(lines[0], ref.left, lines[1], ref.right) if rows else Area(ref.top, lines[0], ref.bottom, lines[1])
        steps.append(_source_moved(cache, origin, area))
    return steps


def _source_moved(cache: PivotCache, origin: XmlElement, area: Area) -> Callable[[], None]:
    from pyopenvba.apps.excel._model import Range

    def move() -> None:
        old = origin.get('ref')
        new = area.address(absolute=False)
        if old == new:
            return
        origin.set('ref', new)
        # SourceData and a refresh read the list where it is now.
        if isinstance(cache.pending_source, Range) or cache.pending_source is None:
            cache.pending_source = None
            cache.source_text = None
        cache.changed = True
    return move


def _carried(steps: list[Callable[[], None]]) -> Callable[[], None]:
    def carry() -> None:
        for step in steps:
            step()
    return carry
