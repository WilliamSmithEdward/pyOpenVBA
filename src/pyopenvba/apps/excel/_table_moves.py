"""Range.Cut through a table: a block holding all of a table takes it along, to another sheet too. Otherwise the table
stays; its header row names its columns afresh, and cells landing beside it make it grow (scripts/measure_table_cut.py,
tests/fixtures/table_cut.json)."""
from __future__ import annotations

from typing import TYPE_CHECKING, Any

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._tables import edited, table_at

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._tables import Table


def _inside(area: Area, block: Area) -> bool:
    return (block.top <= area.top and area.bottom <= block.bottom
            and block.left <= area.left and area.right <= block.right)


def _shifted(area: Area, down: int, across: int) -> Area:
    return Area(area.top + down, area.left + across, area.bottom + down, area.right + across)


def carried(sheet: Worksheet, block: Area) -> list[Table]:
    """The tables a block being cut holds all of."""
    return [table for table in sheet.tables if _inside(table.area, block)]


def _carry(table: Table, target: Worksheet, down: int, across: int) -> None:
    table.area = _shifted(table.area, down, across)
    if table.filter is not None:
        table.filter.area = _shifted(table.filter.area, down, across)
    if table.filter_ref is not None:
        first, second = table.filter_ref
        table.filter_ref = _shifted(first, down, across), _shifted(second, down, across)
    if target is not table.sheet:
        table.sheet.tables.remove(table)
        target.tables.append(table)
        table.sheet = target
        table.view = None
    table.changed = True


def _grown(table: Table, landing: Area) -> None:
    """Cells landing on or right beside a table's block, past its edge, bring the table out to them."""
    from pyopenvba.apps.excel._table_edits import resize

    whole = table.area
    touches = (landing.left <= whole.right + 1 and whole.left <= landing.right
               and landing.top <= whole.bottom + 1 and whole.top <= landing.bottom)
    if touches and (landing.right > whole.right or landing.bottom > whole.bottom):
        resize(table, Area(whole.top, whole.left, max(whole.bottom, landing.bottom), max(whole.right, landing.right)))


def leaving(sheet: Worksheet, moving_cells: dict[tuple[int, int], Any], target: Worksheet, down: int, across: int,
            moving: list[Table]) -> None:
    """Cells on their way out of a table keep its look. Cells moved around inside it stay plain."""
    from pyopenvba.apps.excel._table_unlist import cut_out

    for table in sheet.tables:
        if any(table is one for one in moving):
            continue
        leavers = {place: cell for place, cell in moving_cells.items() if table.area.contains(*place)
                   and not (target is sheet and table.area.contains(place[0] + down, place[1] + across))}
        if leavers:
            cut_out(table, leavers)


def settled(sheet: Worksheet, block: Area, target: Worksheet, landing: Area, moving: list[Table], down: int,
            across: int) -> None:
    """The tables once a block's cells have gone from ``block`` on ``sheet`` to ``landing`` on ``target``."""
    for table in moving:
        _carry(table, target, down, across)
    for table in list(target.tables):
        if not any(table is one for one in moving):
            _grown(table, landing)
    for owner, area in ((sheet, block), (target, landing)):
        for table in list(owner.tables):
            if any(table is one for one in moving) or not table.headers:
                continue
            top = table.area.top
            for column in range(max(area.left, table.area.left), min(area.right, table.area.right) + 1):
                if area.top <= top <= area.bottom and table_at(owner, top, column) is table:
                    edited(owner, top, column)
