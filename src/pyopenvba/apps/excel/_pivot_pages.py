"""PivotTable.ShowPages: a copy of a pivot table on a sheet of its own for each item of a page field, as Excel
makes them.

Measured with scripts/measure_pivot_show_pages.py (tests/fixtures/pivot_show_pages/). Excel makes a sheet for
each item the field shows, from its last item to its first, each before the sheet active then, so the sheets end
in the field's order with the first item's active, in the table's workbook whichever is active. An item a macro
hid has no sheet, nor has the item the field chooses; a field allowing several items gives each item showing a
sheet. A copy keeps the table's definition whole but for its name and its page field, which chooses the copy's
item alone, and stands at the top left of its sheet. A sheet takes its item's name cut to 26 characters, with
" (2)", " (3)" and on after it where another sheet has that name in any case, and keeps the name Sheet and a
number Excel gave it as it added it where no sheet can take the item's: past 31 characters, holding any of
\\ / ? * [ ] :, or an apostrophe first or last. History, which a sheet cannot be named by a macro, is no bar.
With no field named, the first page field the table lists makes the sheets. The events come once every sheet
is made, sheet by sheet in the order they were made: NewSheet, then the sheet it was put before deactivating, it
activating, and its table's cells changing, calculated and updated, as an application's handlers saw them, the
last sheet active all the while. ShowPages returns True.
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_filters import (
    MULTIPLE,
    PAGE,
    allow_multiple,
    choose,
    chosen,
    flag,
    listed,
    orientation,
    page_of,
    place_of,
)
from pyopenvba.apps.excel._pivots import NS, PivotField, PivotTable
from pyopenvba.apps.excel._xstring import encode_attribute as encoded
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: The characters no sheet's name may hold.
REFUSED = '\\/?*[]:'
#: How much of an item's name a sheet takes.
KEPT = 26


def page_name(item: str, taken: set[str]) -> str | None:
    """The name of the sheet ShowPages makes for an item, other sheets having the names ``taken`` (in lower case),
    or None where no sheet can take the item's name."""
    if not item or len(item) > 31 or any(char in item for char in REFUSED) or item[0] == "'" or item[-1] == "'":
        return None
    base = item[:KEPT]
    if base.endswith("'"):
        # A cut ending in an apostrophe is no sheet's name either; Excel was not measured on one.
        return None
    name, number = base, 2
    while name.lower() in taken:
        name = f'{base} ({number})'
        number += 1
    return name


def _page_field(table: PivotTable, wanted: object) -> PivotField:
    """The page field ShowPages makes its sheets of: the one named or given, or the first the table lists."""
    if wanted is MISSING:
        pages = table.root.findall(f'{NS}pageFields/{NS}pageField')
        if not pages:
            raise error(1004)
        index = int(pages[0].get('fld', '0'))
        return PivotField(table, index, table.root.findall(f'{NS}pivotFields/{NS}pivotField')[index])
    if isinstance(wanted, PivotField):
        if wanted.table.entry is not table:
            raise VBAUnsupportedError("PivotTable.ShowPages with another table's field is not implemented")
        field = wanted
    elif isinstance(wanted, str):
        try:
            found = table.PivotFields(wanted)
        except VBARuntimeError:
            raise VBAUnsupportedError('PivotTable.ShowPages naming no field of the table, which stops Excel itself, '
                                      'is not implemented') from None
        assert isinstance(found, PivotField)
        field = found
    else:
        raise VBAUnsupportedError('PivotTable.ShowPages with a PageField that is neither a field nor its name is '
                                  'not implemented')
    if field.data or field.index < 0 or orientation(field.root) != PAGE:
        raise error(1004)
    return field


def _copied(table: PivotTable, index: int, place: int) -> XmlElement:
    """The table's definition with its page field choosing the item at ``place`` in the field's list alone."""
    root = deepcopy(table.root)
    page = page_of(root, index)
    assert page is not None
    choose(page, place)
    allow_multiple(root.findall(f'{NS}pivotFields/{NS}pivotField')[index], False)
    return root


def show_pages(table: PivotTable, wanted: object) -> bool:
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._pivot_create import next_table_name
    from pyopenvba.apps.excel._pivot_layout import apply_render, plan_render
    from pyopenvba.apps.excel._protection import ADDING, refuse_structure

    table = table.entry
    field = _page_field(table, wanted)
    book = table.sheet.book
    refuse_structure(book, ADDING)
    entries = listed(field.root)
    if any(flag(entry, 'm') for entry in entries):
        raise VBAUnsupportedError('PivotTable.ShowPages on a field listing items its source lost is not implemented')
    location = table.root.find(NS + 'location')
    if location is None:
        raise error(1004)
    names = {item.x: str(item.Name()) for item in field.listed_items()}
    choice = None if flag(field.root, MULTIPLE) else chosen(page_of(table.root, field.index))
    pages = [(place, names[place_of(entry)]) for place, entry in enumerate(entries)
             if not flag(entry, 'h') and place != choice and place_of(entry) in names]
    if not pages:
        return True
    # A layout the model cannot make stops ShowPages before it adds a sheet; each copy's differs only by its item.
    plan_render(table, _copied(table, field.index, pages[-1][0]))
    # The page fields fill the rows above the table's body, and a blank row (tests/fixtures/pivot_show_pages/
    # anchored, two_pages).
    body = f"A{int(location.get('rowPageCount', '1')) + 2}"
    if book.application.active_book is not book:
        book.application.activate_book(book)
    made: list[tuple[Worksheet | None, PivotTable, Range]] = []
    for place, item in reversed(pages):
        default = book.allocate_sheet_names(1)[0]
        name = page_name(item, {sheet.name.lower() for sheet in book.sheets_}) or default
        left = book.active_sheet
        sheet = book.add_sheet(name, at=book.sheets_.index(left) if left is not None else None)
        root = _copied(table, field.index, place)
        root.set('name', encoded(next_table_name(sheet)))
        found = root.find(NS + 'location')
        assert found is not None
        found.set('ref', body)
        copy = PivotTable(sheet, table.cache, root)
        copy.sort_names = dict(table.sort_names)
        plan = plan_render(copy, root)
        apply_render(plan)
        sheet.pivot_tables.insert(0, copy)
        book.activate_sheet(sheet)
        made.append((left, copy, Range(sheet, [plan.extent])))
    book.saved = False
    # Excel raises each sheet's events once it has made every sheet (tests/fixtures/pivot_show_pages/ events).
    for left, copy, cells in made:
        _events.new_sheet(book, copy.sheet, left)
        _events.changed(cells)
        if book.calculator.automatic:
            book.calculator.calculate_all(only=copy.sheet, background=True)
        _events.pivot_refresh_calculated(copy.sheet)
        _events.pivot_updated(copy)
    return True
