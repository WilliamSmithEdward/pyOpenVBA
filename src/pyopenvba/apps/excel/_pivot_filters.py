"""Hiding a pivot table's items and choosing its page fields' items, as Excel does.

A field's items are the ``items`` list of its pivotField, in the order the
table shows them, each naming its value in the cache by ``x``: ``h`` hides
one, and ``m`` marks one whose value a refresh no longer found in the
source, which the cache keeps as an unused (``u``) shared item. A page
field's chosen item is its pageField's ``item``, a place in that list, and
``multipleItemSelectionAllowed`` lets it show several items, hidden ones
left out. A field on no axis lists no items until a macro touches one.

Measured with scripts/measure_pivot_filters.py
(tests/fixtures/pivot_filters/): a record counts only where no field on an
axis hides its item and each page field's chosen item is its own; a field
on no axis filters nothing. The last item showing cannot be hidden. Hiding
a page field's chosen item moves the choice to the next item showing,
around to the first. CurrentPage turns off the choice of several items
and keeps the hidden ones; it cannot choose a hidden item. Visible on a
page field choosing one item says whether the item is the chosen one. A
refresh keeps each list's order, marks an item gone from the source and
adds a new one at the end, in the cache's order, hidden where the field
hides any item. Excel saves a page field that chooses an item and allows
several and then cannot open its own file; the model writes such a field
as choosing the one item.
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error, string_order, to_bool, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotCache, PivotField, PivotItem, PivotTable

ROW, COLUMN, PAGE = 1, 2, 3
AXES = {'axisRow': ROW, 'axisCol': COLUMN, 'axisPage': PAGE}
ALL = '(All)'
SEVERAL = '(Multiple Items)'
MULTIPLE = 'multipleItemSelectionAllowed'
#: The order the schema gives a pivotField's attributes, which Excel writes them in.
FIELD_ORDER = ('name axis dataField subtotalCaption showDropDowns hiddenLevel uniqueMemberProperty compact '
               'allDrilled numFmtId outline subtotalTop dragToRow dragToCol multipleItemSelectionAllowed dragToPage '
               'dragToData dragOff showAll insertBlankRow serverField insertPageBreak autoShow topAutoShow '
               'hideNewItems measureFilter includeNewItemsInFilter itemPageCount sortType dataSourceSort '
               'nonAutoSortDefault rankBy defaultSubtotal').split()
#: A listed item: its value's place in the cache, whether it is hidden and whether its value is gone from the source.
Listed = tuple[int, bool, bool]


def flag(element: XmlElement, name: str) -> bool:
    return element.get(name) in ('1', 'true')


def listed(field: XmlElement) -> list[XmlElement]:
    """A pivotField's items in the order the table shows them, its subtotal entries left out."""
    return [item for item in field.findall(f'{NS}items/{NS}item') if item.get('t', 'data') == 'data']


def place_of(item: XmlElement) -> int:
    return int(item.get('x', '-1'))


def item_node(x: int, *, hide: bool = False, gone: bool = False) -> XmlElement:
    """An item as Excel writes one, its flags ahead of ``x`` as the schema orders them."""
    attributes = {'h': '1'} if hide else {}
    if gone:
        attributes['m'] = '1'
    attributes['x'] = str(x)
    return node('item', attributes)


def restate(item: XmlElement, *, hide: bool) -> None:
    rest = {key: value for key, value in item.attrib.items() if key not in ('h', 'm', 'x')}
    fresh = item_node(place_of(item), hide=hide, gone=flag(item, 'm'))
    item.attrib.clear()
    item.attrib.update(fresh.attrib)
    item.attrib.update(rest)


def field_attributes(attributes: dict[str, str]) -> dict[str, str]:
    """A pivotField's attributes in the schema's order, a dataField that is not one left out as Excel leaves it."""
    kept = {key: value for key, value in attributes.items()
            if not (key == 'dataField' and value in ('0', 'false'))}
    ordered = {name: kept[name] for name in FIELD_ORDER if name in kept}
    ordered.update((key, value) for key, value in kept.items() if key not in ordered)
    return ordered


def orientation(field: XmlElement) -> int:
    return AXES.get(field.get('axis', ''), 0)


def page_of(root: XmlElement, index: int) -> XmlElement | None:
    return next((page for page in root.findall(f'{NS}pageFields/{NS}pageField')
                 if int(page.get('fld', '-1')) == index), None)


def chosen(page: XmlElement | None) -> int | None:
    """The place in its field's list of the item a page field shows, or None where it shows every item it can."""
    if page is None or page.get('item') is None:
        return None
    return int(page.get('item', '0'))


def choose(page: XmlElement, position: int | None) -> None:
    rest = {key: value for key, value in page.attrib.items() if key not in ('fld', 'item', 'hier')}
    attributes = {'fld': page.get('fld', '0')}
    if position is not None:
        attributes['item'] = str(position)
    attributes['hier'] = page.get('hier', '-1')
    page.attrib.clear()
    page.attrib.update(attributes)
    page.attrib.update(rest)


def allow_multiple(field: XmlElement, allowed: bool) -> None:
    attributes = {key: value for key, value in field.attrib.items() if key != MULTIPLE}
    if allowed:
        attributes[MULTIPLE] = '1'
    field.attrib.clear()
    field.attrib.update(field_attributes(attributes))


def next_shown(items: list[XmlElement], position: int) -> int | None:
    """The next item after ``position`` that shows and is still in the source, around to the first."""
    for step in range(1, len(items) + 1):
        place = (position + step) % len(items)
        if not flag(items[place], 'h') and not flag(items[place], 'm'):
            return place
    return None


def unused(entry: XmlElement) -> bool:
    return flag(entry, 'u')


def shared(cache_field: XmlElement) -> list[XmlElement]:
    found = cache_field.find(NS + 'sharedItems')
    return list(found) if found is not None else []


def sorted_places(entries: list[XmlElement]) -> list[int]:
    """The places of a cache field's values in the order a field first lists them: sorted, unused ones left out."""
    from pyopenvba.apps.excel._pivot_layout import order

    return [place for place in order(entries) if not unused(entries[place])]


def carried(root: XmlElement, before: list[XmlElement], after: list[XmlElement],
            on_axes: set[int]) -> list[list[Listed] | None]:
    """Each field's list of items, carried from the cache fields ``before`` to those ``after``, or None for a field
    that lists none.

    A listed item keeps its place, its value found again by what it is, and
    is marked gone where the cache keeps the value unused; a value the list
    lacks follows at the end, in the cache's order, hidden where the list
    hides any item. A field newly on an axis lists its values sorted.
    """
    from pyopenvba.apps.excel._pivot_layout import item_key

    lists: list[list[Listed] | None] = []
    for index, field in enumerate(root.findall(f'{NS}pivotFields/{NS}pivotField')):
        old = shared(before[index]) if index < len(before) else []
        new = shared(after[index]) if index < len(after) else []
        if field.find(NS + 'items') is None:
            lists.append([(place, False, False) for place in sorted_places(new)] if index in on_axes else None)
            continue
        # A value spelled the same is the same value; any other is found by what it is, its text read only then.
        spelled = {(entry.tag, entry.get('v')): spot for spot, entry in reversed(list(enumerate(new)))}
        places: dict[tuple[str, object], int] = {}
        found: list[Listed] = []
        taken: set[int] = set()
        for item in listed(field):
            was = place_of(item)
            place: int | None = None
            if 0 <= was < len(old):
                place = spelled.get((old[was].tag, old[was].get('v')))
                if place is None:
                    if not places:
                        places = {item_key(entry): spot for spot, entry in reversed(list(enumerate(new)))}
                    place = places.get(item_key(old[was]))
            if place is not None and place not in taken:
                taken.add(place)
                found.append((place, flag(item, 'h'), unused(new[place])))
        hide = any(hidden for _, hidden, _ in found)
        found += [(place, hide, False) for place, entry in enumerate(new) if place not in taken and not unused(entry)]
        lists.append(found)
    return lists


def retain_missing(cache: PivotCache, fields: XmlElement, reports: list[PivotTable]) -> None:
    """Keep each value a report lists that a refresh no longer finds in the source, as an unused shared item
    after the others, in the cache's old order (tests/fixtures/pivot_filters/ refreshed_two_gone)."""
    from pyopenvba.apps.excel._pivot_layout import item_key

    before = cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    wanted: list[set[int]] = [set() for _ in before]
    for table in reports:
        for index, pivot in enumerate(table.root.findall(f'{NS}pivotFields/{NS}pivotField')[:len(wanted)]):
            wanted[index].update(place_of(item) for item in listed(pivot))
    for index, field in enumerate(list(fields)[:len(before)]):
        values = field.find(NS + 'sharedItems')
        if not wanted[index] or values is None:
            continue
        known = {item_key(entry) for entry in values}
        for place, entry in enumerate(shared(before[index])):
            if place in wanted[index] and item_key(entry) not in known:
                kept = deepcopy(entry)
                kept.set('u', '1')
                values.append(kept)
                known.add(item_key(entry))
        values.set('count', str(len(values)))


def choices(root: XmlElement, lists: list[list[Listed] | None]) -> dict[int, int | None]:
    """Each page field's chosen item, as a place in its carried list: the one chosen before, or the next showing
    where that one is hidden or gone."""
    found: dict[int, int | None] = {}
    for page in root.findall(f'{NS}pageFields/{NS}pageField'):
        index = int(page.get('fld', '0'))
        position = chosen(page)
        listing = lists[index] if 0 <= index < len(lists) else None
        if position is None or listing is None or not 0 <= position < len(listing):
            found[index] = None
        elif listing[position][1] or listing[position][2]:
            found[index] = next_shown([item_node(place, hide=hide, gone=gone) for place, hide, gone in listing],
                                      position)
        else:
            found[index] = position
    return found


def passing(ids: list[list[int]], lists: list[list[Listed] | None], axes: set[int],
            pages: dict[int, int | None]) -> list[int]:
    """The records a table counts: none whose item a field on an axis hides, and each page field's own."""
    hidden: dict[int, set[int]] = {}
    for index in axes:
        listing = lists[index]
        places: set[int] = {place for place, hide, _ in listing if hide} if listing is not None else set()
        if places:
            hidden[index] = places
    pinned: dict[int, int] = {}
    for index, position in pages.items():
        listing = lists[index]
        if position is not None and listing is not None:
            pinned[index] = listing[position][0]
    return [row for row, record in enumerate(ids)
            if all(record[index] not in places for index, places in hidden.items())
            and all(record[index] == place for index, place in pinned.items())]


def page_caption(listing: list[Listed], position: int | None, multiple: bool) -> int | str:
    """What a page field's cell shows: the chosen item, by its value's place in the cache, or (All) or (Multiple
    Items); a field allowing several shows the one item it leaves showing."""
    if position is not None:
        return listing[position][0]
    if not any(hide for _, hide, _ in listing):
        return ALL
    showing = [place for place, hide, gone in listing if not hide and not gone]
    if multiple and len(showing) == 1:
        return showing[0]
    return SEVERAL


def savable(root: XmlElement) -> XmlElement:
    """A report as the model saves it: a page field choosing an item no longer allows several, so that Excel can
    open the file (tests/fixtures/pivot_filters/ page_then_multiple)."""
    fields = root.findall(f'{NS}pivotFields/{NS}pivotField')
    clash = [index for page in root.findall(f'{NS}pageFields/{NS}pageField')
             if chosen(page) is not None and 0 <= (index := int(page.get('fld', '0'))) < len(fields)
             and flag(fields[index], MULTIPLE)]
    if not clash:
        return root
    copied = deepcopy(root)
    copies = copied.findall(f'{NS}pivotFields/{NS}pivotField')
    for index in clash:
        allow_multiple(copies[index], False)
    return copied


# The object model: PivotItem.Visible, PivotField.CurrentPage and the rest.

def _view(field: PivotField) -> tuple[PivotTable, XmlElement]:
    """The table a field is of and its pivotField, refusing the data fields and the Values field."""
    if field.data or field.index < 0:
        raise VBAUnsupportedError('Filtering a pivot table by its values is not implemented')
    return field.table.entry, field.root


def _render(table: PivotTable, root: XmlElement) -> None:
    from pyopenvba.apps.excel._pivot_layout import render

    render(table, root, list(table.root.findall(f'{NS}dataFields/{NS}dataField')))


def _changed(table: PivotTable) -> None:
    table.changed = True
    table.sheet.book.saved = False


def _ensure_items(table: PivotTable, field: XmlElement, index: int) -> bool:
    """List a field's items, sorted, where it lists none yet, as Excel does as a macro first reads them; whether
    the field lists them now. A field whose values its cache keeps in each record lists none."""
    if field.find(NS + 'items') is not None:
        return True
    cache_fields = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    entries = shared(cache_fields[index]) if index < len(cache_fields) else []
    if not entries:
        return False
    items = node('items')
    for place in sorted_places(entries):
        items.append(item_node(place))
    items.append(node('item', {'t': 'default'}))
    items.set('count', str(len(items)))
    field.insert(0, items)
    # The save writes the list; reading the items is no change a macro made.
    table.changed = True
    return True


def list_on_reading(field: PivotField) -> None:
    """List the items of a field on no axis as a macro first reads its PivotItems, as Excel does
    (tests/fixtures/pivot_filters/ count_only)."""
    if not field.data and field.index >= 0 and field.root.find(NS + 'items') is None:
        _ensure_items(field.table.entry, field.root, field.index)


def first_for_names(field: PivotField) -> bool:
    """Whether a name finds the first of a field's items: Excel's, for a field on no axis hiding no item
    (tests/fixtures/pivot_filters/ lookup_twice, lookup_after_hide)."""
    return (not field.data and field.index >= 0 and orientation(field.root) == 0
            and not any(flag(item, 'h') for item in listed(field.root)))


def visible(item: PivotItem) -> bool:
    if item.x < 0:
        return True
    field = item.field.root
    if orientation(field) == PAGE and not flag(field, MULTIPLE):
        position = chosen(page_of(item.field.table.entry.root, item.field.index))
        items = listed(field)
        return position is not None and 0 <= position < len(items) and place_of(items[position]) == item.x
    return not any(flag(entry, 'h') for entry in listed(field) if place_of(entry) == item.x)


def set_visible(item: PivotItem, value: object) -> None:
    shown = to_bool(value)
    if item.x < 0:
        raise VBAUnsupportedError("Hiding the Values field's items or a page field's (All) is not implemented")
    table, live = _view(item.field)
    index = item.field.index
    laid_out = orientation(live) != 0
    root = deepcopy(table.root) if laid_out else table.root
    field = root.findall(f'{NS}pivotFields/{NS}pivotField')[index]
    if not _ensure_items(table, field, index):
        raise VBAUnsupportedError('Hiding an item of a field its cache keeps no list of values for is not implemented')
    items = listed(field)
    target = next((entry for entry in items if place_of(entry) == item.x), None)
    if target is None:
        raise error(1004)
    if shown == (not flag(target, 'h')):
        return
    if not shown:
        if not any(entry is not target and not flag(entry, 'h') for entry in items):
            raise error(1004)
        page = page_of(root, index) if orientation(field) == PAGE else None
        position = chosen(page)
        restate(target, hide=True)
        if page is not None and position is not None and 0 <= position < len(items) and items[position] is target:
            choose(page, next_shown(items, position))
    else:
        restate(target, hide=False)
    if laid_out:
        _render(table, root)
    else:
        _changed(table)


def all_item(field: PivotField) -> PivotItem:
    from pyopenvba.apps.excel._pivots import PivotItem

    return PivotItem(field, None, 1)


def _chosen_item(field: PivotField) -> PivotItem | None:
    position = chosen(page_of(field.table.entry.root, field.index))
    items = listed(field.root)
    if position is None or not 0 <= position < len(items):
        return None
    return next((item for item in field.listed_items() if item.x == place_of(items[position])), None)


def current_page(field: PivotField) -> PivotItem:
    if field.data or field.index < 0 or orientation(field.root) != PAGE:
        raise error(1004)
    found = _chosen_item(field)
    return found if found is not None else all_item(field)


def set_current_page(field: PivotField, value: object) -> None:
    from pyopenvba.apps.excel._pivots import PivotItem

    if field.data or field.index < 0 or orientation(field.root) != PAGE:
        raise error(1004)
    table = field.table.entry
    name = str(value.Name()) if isinstance(value, PivotItem) else to_text(value)
    root = deepcopy(table.root)
    page = page_of(root, field.index)
    pivot = root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index]
    assert page is not None
    if name.casefold() == ALL.casefold():
        choose(page, None)
    else:
        found = next((item for item in field.listed_items()
                      if string_order(str(item.Name()).casefold(), name.casefold()) == 0), None)
        items = listed(pivot)
        position = next((place for place, entry in enumerate(items)
                         if found is not None and place_of(entry) == found.x), None)
        if position is None or flag(items[position], 'h'):
            raise error(1004)
        choose(page, position)
    allow_multiple(pivot, False)
    _render(table, root)


def multiple_allowed(field: PivotField) -> bool:
    return not field.data and field.index >= 0 and flag(field.root, MULTIPLE)


def set_multiple_allowed(field: PivotField, value: object) -> None:
    allowed = to_bool(value)
    table, live = _view(field)
    if flag(live, MULTIPLE) == allowed:
        return
    if orientation(live) == PAGE:
        root = deepcopy(table.root)
        allow_multiple(root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index], allowed)
        _render(table, root)
    else:
        allow_multiple(live, allowed)
        _changed(table)


def _clear(root: XmlElement, index: int | None) -> bool:
    """Show every item of one field, or of every field, and drop their page fields' choices; whether any item was
    hidden or chosen."""
    cleared = False
    for place, field in enumerate(root.findall(f'{NS}pivotFields/{NS}pivotField')):
        if index is None or place == index:
            for item in listed(field):
                if flag(item, 'h'):
                    restate(item, hide=False)
                    cleared = True
    for page in root.findall(f'{NS}pageFields/{NS}pageField'):
        if (index is None or int(page.get('fld', '-1')) == index) and chosen(page) is not None:
            choose(page, None)
            cleared = True
    return cleared


def clear_table(table: PivotTable) -> None:
    table = table.entry
    root = deepcopy(table.root)
    if _clear(root, None):
        _render(table, root)


def clear_field(field: PivotField) -> None:
    table, live = _view(field)
    if orientation(live) == 0:
        if _clear(table.root, field.index):
            _changed(table)
        return
    root = deepcopy(table.root)
    if _clear(root, field.index):
        _render(table, root)


def items_showing(field: PivotField, showing: bool) -> list[object]:
    """VisibleItems or HiddenItems: a page field's chosen item, or (All), against the rest; else the items each
    shows or hides."""
    items = field.listed_items()
    if field.data or field.index < 0:
        return list(items) if showing else []
    if orientation(field.root) == PAGE:
        mine = _chosen_item(field)
        if showing:
            return [mine if mine is not None else all_item(field)]
        return [item for item in items if mine is None or item.x != mine.x]
    return [item for item in items if visible(item) == showing]
