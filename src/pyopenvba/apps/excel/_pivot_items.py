"""A pivot field's calculated items, PivotField.CalculatedItems, as Excel adds, works out, keeps and saves them.

A calculated item is an item of a field whose values a formula works out from the field's other items. Its cache
field lists it after the source's values, marked f="1", each table lists it after the field's other items, marked
the same, and the cache keeps its formula in calculatedItems, each item there named as Field[Item], with a
pivotArea naming the calculated item. Measured with scripts/measure_pivot_calculated_items.py (tests/fixtures/
pivot_calculated_items/):

- CalculatedItems.Add takes a name no item of the field has, in any case, the empty name too, and a formula over
  the field's items, named bare, in apostrophes, as Field[Item] or by place as Field[1], with numbers, text and
  worksheet functions. An item no field has, a field or another field's item is error 1004, and so is a field on
  no axis or on the page area, a data field, the Values field, a field of groups, and any data field of the table
  showing an average, a standard deviation or a variance.
- The formula reads back with each item named as its field names it, bare or in apostrophes, each function in
  capitals and each number as a formula spells it, and a place kept as a place; the space before an item moves
  after it, as in a calculated field's formula, and once the file is opened again the formula reads without its
  spaces.
- In a cell the item is its formula over the values the items it names have there, their hidden items' too, an
  item no record reaches there 0. A total over its field takes every item's value together, each calculated
  item's as it is at each combination of the other fields' items, by the data field's function: summed for a
  sum or a count, the greatest, the least, multiplied. That is why Excel takes no average, standard deviation or
  variance beside one. The item shows under every item outside it, and every item of the fields inside it shows
  under it.
- Sorted by name its field keeps its calculated items last; sorted by a value they sort with the rest. A field
  with one cannot go to the page area, take custom subtotals or group by ranges: error 1004. Deleting an item
  another names leaves #NAME? in the other's formula.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._pivot_create import encoded, node
from pyopenvba.apps.excel._pivot_formulas import Piece, Refused, engine_node, fields, quoted, read, spelled, \
    worked_out
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import EMPTY, MISSING, error, string_order, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._pivot_calculated import Measures
    from pyopenvba.apps.excel._pivot_filters import Listed
    from pyopenvba.apps.excel._pivots import PivotCache, PivotField, PivotItem, PivotTable
    from pyopenvba.formula._calc.nodes import Node

#: Functions a data field cannot show beside a calculated item (tests/fixtures/pivot_calculated_items/ add_average,
#: function_stdev, function_var).
REFUSED_FUNCTIONS = frozenset({'average', 'stdDev', 'stdDevp', 'var', 'varp'})


def is_formula(entry: XmlElement) -> bool:
    """Whether a cache field's item is a calculated item's."""
    return entry.get('f') in ('1', 'true')


def _fields(cache: PivotCache) -> list[XmlElement]:
    return cache.root.findall(f'{NS}cacheFields/{NS}cacheField')


def _shared(cache_field: XmlElement) -> list[XmlElement]:
    found = cache_field.find(NS + 'sharedItems')
    return list(found) if found is not None else []


def _reference(element: XmlElement) -> tuple[int, int] | None:
    """The field and the place of the item a calculatedItem's pivotArea names."""
    reference = element.find(f'{NS}pivotArea/{NS}references/{NS}reference')
    spot = reference.find(NS + 'x') if reference is not None else None
    if reference is None or spot is None:
        return None
    return int(reference.get('field', '-1')), int(spot.get('v', '-1'))


def entries(cache: PivotCache) -> list[tuple[int, int, XmlElement]]:
    """Each calculated item the cache keeps: its field, its place in the field's items, and its element."""
    container = cache.root.find(NS + 'calculatedItems')
    found: list[tuple[int, int, XmlElement]] = []
    elements: list[XmlElement] = list(container) if container is not None else []
    for element in elements:
        named = _reference(element)
        if named is not None:
            found.append((named[0], named[1], element))
    return found


def has_items(cache: PivotCache, field: int | None = None) -> bool:
    return any(field is None or each == field for each, _, _ in entries(cache))


def _element(cache: PivotCache, field: int, place: int) -> XmlElement | None:
    return next((element for each, spot, element in entries(cache) if (each, spot) == (field, place)), None)


# Reading and spelling a formula over a field's items.

def _names(cache: PivotCache, field: int, cache_fields: list[XmlElement] | None = None) -> list[str]:
    """Each of a field's items as a formula names it: as its field shows it, in ``cache_fields`` where given, else
    in the cache's own."""
    from pyopenvba.apps.excel._pivot_layout import field_format, item_caption

    cache_field = (cache_fields if cache_fields is not None else _fields(cache))[field]
    code = field_format(cache.book.stylesheet.number_formats, None, cache_field)
    return [decode_text(entry.get('v', '')) if entry.tag == NS + 's' else item_caption(entry, code)
            for entry in _shared(cache_field)]


def _parse(cache: PivotCache, field: int, text: str, cache_fields: list[XmlElement] | None = None) -> list[Piece]:
    """A calculated item's formula over its field's items, Refused where Excel refuses it."""
    names = _names(cache, field, cache_fields)
    owner = decode_text((cache_fields if cache_fields is not None else _fields(cache))[field].get('name', ''))

    def place(name: str) -> int | None:
        return next((spot for spot, each in enumerate(names) if string_order(each.casefold(), name.casefold()) == 0),
                    None)

    def named(field_name: str, inside: str) -> tuple[str, int] | None:
        if string_order(field_name.casefold(), owner.casefold()) != 0:
            return None
        if inside.startswith("'") and inside.endswith("'") and len(inside) > 1:
            found = place(inside[1:-1].replace("''", "'"))
            return None if found is None else ('field', found)
        if inside.isdecimal():
            return 'position', int(inside)
        found = place(inside)
        return None if found is None else ('field', found)

    return read(text, place, named)


def _spelling(cache: PivotCache, field: int, pieces: list[Piece], *, stored: bool, given: bool = True) -> str:
    names = _names(cache, field)
    owner = quoted(decode_text(_fields(cache)[field].get('name', '')))
    if stored:
        return spelled(pieces, lambda spot: f'{owner}[{quoted(names[spot])}]', bare=True,
                       place_of=lambda place: f'{owner}[{place}]')
    return '=' + spelled(pieces, lambda spot: names[spot], given=given, place_of=lambda place: f'{owner}[{place}]')


def formula_of(item: PivotItem) -> object:
    """PivotItem.Formula: a calculated item's formula as Excel reads it back; error 1004 for any other item."""
    cache = item.field.table.entry.cache
    if item.field.data or item.field.index < 0 or item.x < 0:
        raise error(1004)
    element = _element(cache, item.field.index, item.x)
    if element is None:
        raise error(1004)
    given = cache.given_items.get((item.field.index, item.x))
    if given is not None:
        return given
    try:
        pieces = _parse(cache, item.field.index, decode_text(element.get('formula', '')))
    except Refused:
        raise VBAUnsupportedError('A calculated item whose formula the model cannot read is not implemented') from None
    # Read from a file the formula has no spaces (tests/fixtures/pivot_calculated_items/ spelled).
    return _spelling(cache, item.field.index, [replace(piece, space='') for piece in pieces], stored=False,
                     given=False)


def is_calculated(item: PivotItem) -> bool:
    entries_ = _shared(_fields(item.field.table.entry.cache)[item.field.index]) if item.field.index >= 0 else []
    return 0 <= item.x < len(entries_) and is_formula(entries_[item.x]) and not item.field.data


# Adding, changing and deleting calculated items.

def _admitted(field: PivotField) -> None:
    """Error 1004 where Excel refuses a calculated item to a field (tests/fixtures/pivot_calculated_items/
    errors*, add_average)."""
    from pyopenvba.apps.excel._pivot_groups import field_group, is_database

    table = field.table.entry
    if field.data or field.index < 0 or field.root.get('axis') not in ('axisRow', 'axisCol'):
        raise error(1004)
    cache_field = _fields(table.cache)[field.index]
    if not is_database(cache_field):
        raise error(1004)
    if field_group(cache_field) is not None:
        raise VBAUnsupportedError('Calculated items in a grouped pivot field are not implemented')
    if any(report.get('subtotal') in REFUSED_FUNCTIONS for report in table.root.findall(f'{NS}dataFields/{NS}dataField')):
        raise error(1004)


def _tables(cache: PivotCache) -> list[PivotTable]:
    return [table for sheet in cache.book.sheets_ for table in sheet.pivot_tables if table.cache is cache]


def _render_all(cache: PivotCache) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    for table in _tables(cache):
        table.changed = True
        rerender(table, deepcopy(table.root))
    cache.changed = True
    cache.book.saved = False


def _list(table: PivotTable, field: int, place: int) -> None:
    """Add a calculated item to a table's list of its field's items, after the others, where it lists them."""
    from pyopenvba.apps.excel._pivot_filters import item_node

    pivot = table.root.findall(f'{NS}pivotFields/{NS}pivotField')[field]
    items = pivot.find(NS + 'items')
    if items is None:
        return
    data = [item for item in items if item.get('t', 'data') == 'data']
    rest = [item for item in items if item.get('t', 'data') != 'data']
    items[:] = [*data, item_node(place, formula=True), *rest]
    items.set('count', str(len(items)))


def _listing(field: PivotField, written: str) -> None:
    """List the items of each other field the formula names as Field[Item], as Excel lists them while it reads the
    formula, whether it then takes it or not (tests/fixtures/pivot_calculated_items/ other_field_item,
    errors_hidden_field)."""
    from pyopenvba.apps.excel._pivot_filters import list_on_reading
    from pyopenvba.apps.excel._pivot_formulas import item_owners
    from pyopenvba.apps.excel._pivots import PivotField

    named = {name.casefold() for name in item_owners(written)}
    table = field.table.entry
    for index, cache_field in enumerate(_fields(table.cache)):
        if index != field.index and decode_text(cache_field.get('name', '')).casefold() in named:
            other = table.PivotFields(index + 1)
            if isinstance(other, PivotField):
                list_on_reading(other)


def add(field: PivotField, name: object, formula: object) -> object:
    """CalculatedItems.Add: a new item of the field, after its others, worked out by the formula; the PivotItem."""
    text, written = to_text(name), to_text(formula)
    if not field.data and field.index >= 0:
        _listing(field, written)
    _admitted(field)
    table = field.table.entry
    cache = table.cache
    if any(string_order(each.casefold(), text.casefold()) == 0 for each in _names(cache, field.index)):
        raise error(1004)
    try:
        pieces = _parse(cache, field.index, written)
        engine_node(pieces, table.sheet)
    except Refused:
        raise error(1004) from None
    cache_field = _fields(cache)[field.index]
    shared = cache_field.find(NS + 'sharedItems')
    if shared is None:
        shared = node('sharedItems')
        cache_field.insert(0, shared)
    place = len(shared)
    shared.append(node('s', {'v': encoded(text), 'f': '1'}))
    shared.set('count', str(len(shared)))
    container = cache.root.find(NS + 'calculatedItems')
    if container is None:
        container = node('calculatedItems')
        anchor = cache.root.find(NS + 'cacheFields')
        assert anchor is not None
        cache.root.insert(list(cache.root).index(anchor) + 1, container)
    element = node('calculatedItem', {'formula': encoded(_spelling(cache, field.index, pieces, stored=True))})
    area = node('pivotArea', {'cacheIndex': '1', 'outline': '0', 'fieldPosition': '0'})
    references = node('references', {'count': '1'})
    reference = node('reference', {'field': str(field.index), 'count': '1'})
    reference.append(node('x', {'v': str(place)}))
    references.append(reference)
    area.append(references)
    element.append(area)
    container.append(element)
    container.set('count', str(len(container)))
    cache.given_items[(field.index, place)] = _spelling(cache, field.index, pieces, stored=False)
    for each in _tables(cache):
        _list(each, field.index, place)
    _render_all(cache)
    return next(item for item in field.listed_items() if item.x == place)


def add_missing(field: PivotField, name: object) -> None:
    """PivotItems.Add: an item no record holds, after the field's others, the cache keeping its value as unused and
    the table listing it as missing; the table's rows stay as they were (tests/fixtures/pivot_filter_order.json)."""
    from pyopenvba.apps.excel._pivot_filters import item_node

    text = to_text(name)
    table = field.table.entry
    cache = table.cache
    if field.data or field.index < 0:
        raise VBAUnsupportedError('PivotItems.Add on this field is not implemented')
    if any(string_order(each.casefold(), text.casefold()) == 0 for each in _names(cache, field.index)):
        raise VBAUnsupportedError('PivotItems.Add of an item the field has is not implemented')
    shared = _fields(cache)[field.index].find(NS + 'sharedItems')
    if shared is None or any(entry.tag != NS + 's' for entry in shared):
        raise VBAUnsupportedError('PivotItems.Add on a field of other than text is not implemented')
    place = len(shared)
    shared.append(node('s', {'v': encoded(text), 'u': '1'}))
    shared.set('count', str(len(shared)))
    for each in _tables(cache):
        items = each.root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index].find(NS + 'items')
        if items is None:
            continue
        data = [item for item in items if item.get('t', 'data') == 'data']
        rest = [item for item in items if item.get('t', 'data') != 'data']
        items[:] = [*data, item_node(place, gone=True), *rest]
        items.set('count', str(len(items)))
    _render_all(cache)


def set_formula(item: PivotItem, value: object) -> None:
    """PivotItem.Formula set: a calculated item worked out by another formula; error 1004 for any other item."""
    cache = item.field.table.entry.cache
    element = _element(cache, item.field.index, item.x) if item.x >= 0 and not item.field.data else None
    if element is None:
        raise error(1004)
    try:
        pieces = _parse(cache, item.field.index, to_text(value))
        engine_node(pieces, item.field.table.entry.sheet)
    except Refused:
        raise error(1004) from None
    element.set('formula', encoded(_spelling(cache, item.field.index, pieces, stored=True)))
    cache.given_items[(item.field.index, item.x)] = _spelling(cache, item.field.index, pieces, stored=False)
    _render_all(cache)


def delete(item: PivotItem) -> None:
    """PivotItem.Delete: a calculated item taken away, #NAME? where another's formula names it, the items after it
    in the cache moved up one; error 1004 for any other item (tests/fixtures/pivot_calculated_items/ delete,
    delete_member, delete_referenced)."""
    from pyopenvba.apps.excel._pivot_formulas import without

    cache = item.field.table.entry.cache
    field, place = item.field.index, item.x
    element = _element(cache, field, place) if place >= 0 and not item.field.data else None
    if element is None:
        raise error(1004)
    container = cache.root.find(NS + 'calculatedItems')
    assert container is not None
    # The other items' formulas, read while the gone item is still there to name.
    others: list[tuple[int, XmlElement, list[Piece]]] = []
    for each, spot, other in entries(cache):
        if each == field and other is not element:
            try:
                others.append((spot, other, without(_parse(cache, field, decode_text(other.get('formula', ''))),
                                                    place)))
            except Refused:
                continue
    names = _names(cache, field)
    del names[place]
    owner = quoted(decode_text(_fields(cache)[field].get('name', '')))
    container.remove(element)
    given = {key: value for key, value in cache.given_items.items() if key[0] != field}
    for spot, other, pieces in others:
        # The items after the gone one move up a place.
        moved = [replace(piece, field=piece.field - 1) if piece.kind == 'field' and piece.field > place else piece
                 for piece in pieces]
        after = spot - 1 if spot > place else spot
        reference = other.find(f'{NS}pivotArea/{NS}references/{NS}reference/{NS}x')
        if reference is not None:
            reference.set('v', str(after))
        other.set('formula', encoded(spelled(moved, lambda at: f'{owner}[{quoted(names[at])}]', bare=True,
                                             place_of=lambda at: f'{owner}[{at}]')))
        if (field, spot) in cache.given_items:
            given[(field, after)] = '=' + spelled(moved, lambda at: names[at], place_of=lambda at: f'{owner}[{at}]')
    cache.given_items = given
    container.set('count', str(len(container)))
    if not len(container):
        cache.root.remove(container)
    shared = _fields(cache)[field].find(NS + 'sharedItems')
    assert shared is not None
    shared.remove(shared[place])
    shared.set('count', str(len(shared)))
    _renumber_records(cache, field, place)
    for table in _tables(cache):
        pivot = table.root.findall(f'{NS}pivotFields/{NS}pivotField')[field]
        items = pivot.find(NS + 'items')
        if items is None:
            continue
        for entry in list(items):
            spot = int(entry.get('x', '-1')) if entry.get('t', 'data') == 'data' else -1
            if spot == place:
                items.remove(entry)
            elif spot > place:
                entry.set('x', str(spot - 1))
        items.set('count', str(len(items)))
    _render_all(cache)


def _renumber_records(cache: PivotCache, field: int, place: int) -> None:
    """The records' references to a field's items after one, moved up one as the item before them goes."""
    from pyopenvba.apps.excel._pivot_aggregate import records

    records(cache)
    if cache.records is None:
        return
    for record in cache.records.findall(NS + 'r'):
        if field < len(record) and record[field].tag == NS + 'x':
            spot = int(record[field].get('v', '0'))
            if spot > place:
                record[field].set('v', str(spot - 1))


class CalculatedItems(VBACollection, ExcelObject):
    """A field's calculated items, in the order its table lists them."""

    vba_type_name = 'CalculatedItems'

    def __init__(self, field: PivotField) -> None:
        self.field = field

    def vba_items(self) -> list[object]:
        if self.field.data or self.field.index < 0:
            return []
        return [item for item in self.field.listed_items() if is_calculated(item)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        from pyopenvba.apps.excel._pivots import PivotItem

        if isinstance(index, str):
            for item in items:
                if isinstance(item, PivotItem) and string_order(str(item.Name()).casefold(), index.casefold()) == 0:
                    return item
            raise error(1004)
        position = int(to_integer(index, 'Long'))
        if 1 <= position <= len(items):
            return items[position - 1]
        # An index past the items is error 9 (tests/fixtures/pivot_calculated_items/ collection).
        raise error(9)

    @member
    def Parent(self) -> object:
        return self.field

    @method
    def Add(self, Name: object, Formula: object, UseStandardFormula: object = MISSING) -> object:
        return add(self.field, Name, Formula)


# Working calculated items out in a table.

def item_field(cache_fields: list[XmlElement], cache: PivotCache, axes: list[int]) -> int | None:
    """The field on the table's rows or columns whose calculated items the table shows, if any."""
    found = sorted({field for field, _, _ in entries(cache) if field in axes})
    if len(found) > 1:
        raise VBAUnsupportedError('Calculated items in more than one field of a pivot table are not implemented')
    return found[0] if found else None


def _combined(parts: list[object], kind: str) -> object:
    """Values taken together as a total over a field takes its items' (tests/fixtures/pivot_calculated_items/
    function_*)."""
    present = [part for part in parts if part is not EMPTY]
    for part in present:
        if isinstance(part, ExcelError):
            return part
    numbers = [float(part) for part in present if isinstance(part, (int, float)) and not isinstance(part, bool)]
    if not numbers:
        return EMPTY
    if kind == 'max':
        return max(numbers)
    if kind == 'min':
        return min(numbers)
    result = 1.0 if kind == 'product' else 0.0
    for number in numbers:
        result = result * number if kind == 'product' else result + number
    return result


class Workout:
    """A table's calculated items worked out: the rows they add to its records, one for each item at each
    combination of the other fields' items on its rows and columns, and the value of any set of rows, the real
    rows' by their data field's function and each added row's by its item's formula there, together as a total
    over the field takes its items."""

    def __init__(self, cache: PivotCache, cache_fields: list[XmlElement], field: int, listing: list[Listed],
                 records: list[list[XmlElement]], ids: list[list[int]], counted: list[int], other: list[int],
                 measures: Measures, sheet: Worksheet) -> None:
        self.field, self.records, self.ids, self.other = field, records, ids, other
        self.measures, self.sheet = measures, sheet
        self.real = len(ids)
        self.made: dict[int, tuple[int, tuple[int, ...]]] = {}
        order = [place for place, _, _ in listing]
        self.formulas: dict[int, tuple[list[Piece], Node]] = {}
        for each, place, element in entries(cache):
            if each != field:
                continue
            try:
                pieces = _parse(cache, field, decode_text(element.get('formula', '')), cache_fields)
            except Refused:
                raise VBAUnsupportedError('A calculated item whose formula the model cannot read is not '
                                          'implemented') from None
            # A place names the item at that place in the field's list (tests/fixtures/pivot_calculated_items/
            # position_reference).
            pieces = [Piece('field', piece.text, piece.space, order[piece.field - 1])
                      if piece.kind == 'position' and 0 < piece.field <= len(order) else piece for piece in pieces]
            if any(piece.kind == 'position' for piece in pieces):
                raise VBAUnsupportedError('A calculated item naming a place past its field\'s items is not '
                                          'implemented')
            self.formulas[place] = pieces, engine_node(pieces, sheet)
        self.shown = [place for place, hide, _ in listing if not hide and place in self.formulas]
        # The records reaching each item at each combination of the other fields' items, hidden items' too.
        self.reaching: dict[tuple[int, tuple[int, ...]], list[int]] = {}
        self.contexts: dict[tuple[int, ...], int] = {}
        for row in counted:
            context = tuple(ids[row][index] for index in other)
            self.contexts.setdefault(context, row)
            self.reaching.setdefault((ids[row][field], context), []).append(row)
        self.worked: dict[tuple[int, int, tuple[int, ...]], object] = {}

    def extend(self, selected: list[int]) -> list[int]:
        """Add a row for each calculated item showing at each combination of the other fields' items, and say
        which rows the table counts."""
        added: list[int] = []
        for place in self.shown:
            for context, row in self.contexts.items():
                fresh = list(self.ids[row])
                fresh[self.field] = place
                self.ids.append(fresh)
                self.records.append(self.records[row])
                self.made[len(self.ids) - 1] = place, context
                added.append(len(self.ids) - 1)
        return [*selected, *added]

    def _worked(self, report: XmlElement, place: int, context: tuple[int, ...], seen: tuple[int, ...] = ()) -> object:
        key = id(report), place, context
        if key not in self.worked:
            if place in seen:
                raise VBAUnsupportedError('Calculated items that name each other in a loop are not implemented')
            pieces, made = self.formulas[place]
            values: dict[int, object] = {}
            for named in fields(pieces):
                if named in self.formulas:
                    values[named] = self._worked(report, named, context, (*seen, place))
                    continue
                rows = self.reaching.get((named, context), [])
                total = self.measures.value(report, [self.records[row] for row in rows]) if rows else EMPTY
                values[named] = 0.0 if total is EMPTY else total
            self.worked[key] = worked_out(made, values, self.sheet)
        return self.worked[key]

    def value(self, report: XmlElement, rows: list[int], other: str = '') -> object:
        """The value a set of the table's rows shows for a data field."""
        real = [self.records[row] for row in rows if row < self.real]
        added = [self.made[row] for row in rows if row >= self.real]
        base = self.measures.value(report, real, other)
        if not added:
            return base
        if other:
            raise VBAUnsupportedError('A custom subtotal beside a calculated item is not implemented')
        parts = [base, *(self._worked(report, place, context) for place, context in added)]
        return _combined(parts, report.get('subtotal', 'sum'))


def carry_items(before: XmlElement, after: XmlElement, new_fields: list[XmlElement]) -> bool:
    """Carry a cache's calculated items over a refresh: each one after its field's values from the source, in
    ``new_fields``, at its new place there (tests/fixtures/pivot_calculated_items/ refresh); whether it had any."""
    from pyopenvba.apps.excel._pivot_groups import is_database

    container = before.find(NS + 'calculatedItems')
    for stale in after.findall(NS + 'calculatedItems'):
        after.remove(stale)
    if container is None:
        return False
    old_fields = before.findall(f'{NS}cacheFields/{NS}cacheField')
    moved = deepcopy(container)
    for element in moved:
        named = _reference(element)
        if named is None or not 0 <= named[0] < min(len(old_fields), len(new_fields)) \
                or not is_database(new_fields[named[0]]):
            continue
        field, place = named
        old_shared = _shared(old_fields[field])
        shared = new_fields[field].find(NS + 'sharedItems')
        if shared is None or not 0 <= place < len(old_shared):
            continue
        shared.append(deepcopy(old_shared[place]))
        shared.set('count', str(len(shared)))
        spot = element.find(f'{NS}pivotArea/{NS}references/{NS}reference/{NS}x')
        assert spot is not None
        spot.set('v', str(len(shared) - 1))
    anchor = after.find(NS + 'cacheFields')
    assert anchor is not None
    after.insert(list(after).index(anchor) + 1, moved)
    return True
