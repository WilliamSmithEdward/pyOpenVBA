"""A pivot table's calculated fields: PivotTable.CalculatedFields, and a calculated field's Formula, as Excel keeps
and shows them.

A calculated field is a cache field with a formula and no values of its own: the cache keeps it after the fields
of the source, ``formula`` its formula as :mod:`pyopenvba.apps.excel._pivot_formulas` spells it and
``databaseField="0"``, and each table of the cache a field for it that no axis takes. In the data area it shows
its formula worked out over the sums of the fields it names, whatever function the data field shows, and nothing
on a subtotal line by another function. Measured with scripts/measure_pivot_calculated_fields.py
(tests/fixtures/pivot_calculated_fields/).
"""
from __future__ import annotations

from collections.abc import Callable, Sequence
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._pivot_create import encoded, node
from pyopenvba.apps.excel._pivot_filters import field_attributes
from pyopenvba.apps.excel._pivot_formulas import Piece, Refused, engine_node, fields, read, spelled, without, \
    worked_out
from pyopenvba.apps.excel._pivots import NS, PivotField, decode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, VBAInt, error, string_order, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable
    from pyopenvba.formula._calc.nodes import Node

#: A calculated field's own field in a table: one no axis takes, with no subtotal.
TABLE_FIELD = {'dragToRow': '0', 'dragToCol': '0', 'dragToPage': '0', 'showAll': '0', 'defaultSubtotal': '0'}
#: XlPivotFieldDataType: a calculated field reads as text, as Excel's does.
TEXT, NUMBER, DATE = -4158, -4145, 2


def is_calculated(cache_field: XmlElement) -> bool:
    return cache_field.get('formula') is not None and cache_field.get('databaseField') == '0'


def _cache_fields(cache: PivotCache) -> list[XmlElement]:
    return cache.root.findall(f'{NS}cacheFields/{NS}cacheField')


def source_count(cache_fields: Sequence[XmlElement]) -> int:
    """How many of a cache's fields come from its source: all but the calculated ones, which follow them."""
    count = sum(1 for field in cache_fields if not is_calculated(field))
    if any(is_calculated(field) for field in cache_fields[:count]):
        raise VBAUnsupportedError('A pivot cache with a calculated field ahead of its source fields is not implemented')
    return count


def _tables(cache: PivotCache) -> list[PivotTable]:
    return [table for sheet in cache.book.sheets_ for table in sheet.pivot_tables if table.cache is cache]


def captions(table: PivotTable) -> list[str]:
    """Each field's name as the table shows it: its caption, else its cache field's."""
    cache_fields = _cache_fields(table.cache)
    out: list[str] = []
    for index, field in enumerate(table.root.findall(f'{NS}pivotFields/{NS}pivotField')):
        own = field.get('name')
        out.append(decode_text(own) if own is not None else
                   decode_text(cache_fields[index].get('name', '')) if index < len(cache_fields) else '')
    return out


def _finder(names: list[str]) -> Callable[[str], int | None]:
    """The place of the field a formula names, as ``names`` name the fields, in any case."""
    def resolve(name: str) -> int | None:
        wanted = name.casefold()
        return next((index for index, one in enumerate(names) if string_order(one.casefold(), wanted) == 0), None)
    return resolve


def _cache_names(cache: PivotCache) -> list[str]:
    return [decode_text(field.get('name', '')) for field in _cache_fields(cache)]


def stored(cache: PivotCache, index: int) -> list[Piece]:
    """A calculated field's formula as its cache keeps it, each field named as the cache names it."""
    fields_ = _cache_fields(cache)
    try:
        return read(decode_text(fields_[index].get('formula', '')), _finder(_cache_names(cache)))
    except Refused:
        raise VBAUnsupportedError('A calculated field whose formula the model cannot read is not implemented') from None


def _reaches(cache: PivotCache, start: set[int], target: int) -> bool:
    """Whether the fields ``start`` names lead, through calculated fields' formulas, to ``target``."""
    fields_ = _cache_fields(cache)
    seen: set[int] = set()
    pending = list(start)
    while pending:
        index = pending.pop()
        if index == target:
            return True
        if index in seen or not 0 <= index < len(fields_) or not is_calculated(fields_[index]):
            continue
        seen.add(index)
        pending.extend(fields(stored(cache, index)))
    return False


def _shown(table: PivotTable, index: int) -> str:
    """A calculated field's formula as the table shows it, each field named by its caption: as Excel spelled it
    back where a macro gave it in this session, and read again, as Excel reads a file's as it opens it, where the
    file gave it (tests/fixtures/pivot_calculated_fields/ spaced_loose, once opened again)."""
    names = captions(table)
    given = decode_text(_cache_fields(table.cache)[index].get('name', '')) not in table.cache.given_formulas
    return '=' + spelled(stored(table.cache, index), lambda place: names[place] if place < len(names) else '',
                         given=given)


def _changed(cache: PivotCache) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    cache.changed = True
    for table in _tables(cache):
        if table.root.findall(f'{NS}dataFields/{NS}dataField'):
            rerender(table, deepcopy(table.root))
        table.changed = True
    cache.book.saved = False


def _refused(table: PivotTable) -> object:
    """A formula Excel refuses to add: with alerts off it adds nothing and says nothing, and with them on it shows a
    message the model cannot (tests/fixtures/pivot_calculated_fields/ errors)."""
    if table.sheet.book.application.display_alerts:
        raise VBAUnsupportedError('A calculated field whose formula Excel refuses shows a message, which is not '
                                  'implemented')
    return NOTHING


def add(table: PivotTable, name: object, formula: object) -> object:
    """CalculatedFields.Add: a new calculated field, at the end of its cache's fields and of each table's; a name a
    field has is error 1004."""
    table = table.entry
    caption, text = to_text(name), to_text(formula)
    if any(string_order(one.casefold(), caption.casefold()) == 0 for one in [*_cache_names(table.cache),
                                                                            *captions(table)]):
        raise error(1004)
    try:
        pieces = read(text, _finder(captions(table)))
        engine_node(pieces, table.sheet)
    except Refused:
        return _refused(table)
    cache = table.cache
    container = cache.root.find(NS + 'cacheFields')
    assert container is not None
    cache_names = _cache_names(cache)
    container.append(node('cacheField', {'name': encoded(caption), 'numFmtId': '0',
                                         'formula': encoded(spelled(pieces, lambda place: cache_names[place])),
                                         'databaseField': '0'}))
    container.set('count', str(len(container)))
    cache.given_formulas.add(caption)
    for one in _tables(cache):
        own = one.root.find(NS + 'pivotFields')
        assert own is not None
        own.append(node('pivotField', dict(TABLE_FIELD)))
        own.set('count', str(len(own)))
        one.changed = True
    cache.changed = True
    table.sheet.book.saved = False
    index = len(container) - 1
    return PivotField(table, index, table.root.findall(f'{NS}pivotFields/{NS}pivotField')[index])


def _calculated_index(field: PivotField) -> int | None:
    fields_ = _cache_fields(field.table.entry.cache)
    if field.data or not 0 <= field.index < len(fields_) or not is_calculated(fields_[field.index]):
        return None
    return field.index


def calculated_view(field: PivotField) -> bool:
    """Whether a field is a calculated one, not its data field."""
    return _calculated_index(field) is not None


def rename(field: PivotField, value: object, *, caption: bool) -> None:
    """Caption: the name a field shows by and formulas name it by, its cache's name kept. Name does the same on a
    field on an axis; on a calculated field it is error 1004, and on a field in the data area alone it lists the
    field's items, takes the name and then is error 1004, as in Excel (tests/fixtures/pivot_calculated_fields/
    renamed_caption, renamed, caption_lookup, source_caption, source_name)."""
    from pyopenvba.apps.excel._pivot_filters import ensure_items, rerender

    if field.data or field.index < 0:
        raise VBAUnsupportedError("Renaming a pivot table's data field or its Values field is not implemented")
    table = field.table.entry
    on_axis = field.root.get('axis') is not None
    quirk = not caption and not on_axis and field.root.get('dataField') in ('1', 'true')
    if not caption and calculated_view(field):
        raise error(1004)
    if not caption and not on_axis and not quirk:
        raise VBAUnsupportedError('Naming a pivot field on no axis is not implemented')
    text = to_text(value)
    if any(string_order(one.casefold(), text.casefold()) == 0
           for place, one in enumerate(captions(table)) if place != field.index):
        raise error(1004)
    if quirk:
        ensure_items(table, field.root, field.index)
    field.root.set('name', encoded(text))
    ordered = field_attributes(dict(field.root.attrib))
    field.root.attrib.clear()
    field.root.attrib.update(ordered)
    table.changed = True
    table.sheet.book.saved = False
    if on_axis:
        # Outline and tabular forms show a field's name over its items.
        rerender(table, deepcopy(table.root))
    if quirk:
        raise error(1004)


def formula(field: PivotField) -> object:
    """Formula: a calculated field's, as the table names its fields; error 1004 on any other field."""
    index = _calculated_index(field)
    if index is None:
        raise error(1004)
    return _shown(field.table.entry, index)


def set_formula(field: PivotField, value: object) -> None:
    """Formula on a calculated field: a new formula, error 1004 where Excel refuses it, one leading back to the field
    itself among them. On any other field it does nothing, as in Excel (tests/fixtures/pivot_calculated_fields/
    errors, circular)."""
    index = _calculated_index(field)
    if index is None:
        return
    table = field.table.entry
    try:
        pieces = read(to_text(value), _finder(captions(table)))
        engine_node(pieces, table.sheet)
    except Refused:
        raise error(1004) from None
    if _reaches(table.cache, fields(pieces), index):
        raise error(1004)
    cache_names = _cache_names(table.cache)
    _cache_fields(table.cache)[index].set('formula', encoded(spelled(pieces, lambda place: cache_names[place])))
    table.cache.given_formulas.add(cache_names[index])
    _changed(table.cache)


def is_calculated_field(field: PivotField) -> object:
    """IsCalculated: error 1004 on a data field (tests/fixtures/pivot_calculated_fields/ margin)."""
    if field.data:
        raise error(1004)
    return _calculated_index(field) is not None


def data_type(field: PivotField) -> object:
    """DataType: text for a calculated field and a field of text, a number for one of numbers, a date for one of
    dates (tests/fixtures/pivot_calculated_fields/ items_of_calculated, caption_lookup)."""
    fields_ = _cache_fields(field.table.entry.cache)
    if field.index < 0 or field.index >= len(fields_) or _calculated_index(field) is not None:
        return VBAInt(TEXT, 'Long')
    shared = fields_[field.index].find(NS + 'sharedItems')
    flags = shared.attrib if shared is not None else {}
    if flags.get('containsDate') in ('1', 'true'):
        return VBAInt(DATE, 'Long')
    if flags.get('containsNumber') in ('1', 'true') and flags.get('containsString', '1') in ('0', 'false'):
        return VBAInt(NUMBER, 'Long')
    return VBAInt(TEXT, 'Long')


def delete(field: PivotField) -> None:
    """Delete on a calculated field: gone from its cache and every table of it, its data fields with it, and #NAME?
    where another formula named it (tests/fixtures/pivot_calculated_fields/ deleted, delete_referenced)."""
    index = _calculated_index(field)
    if index is None:
        raise VBAUnsupportedError('Deleting a pivot field other than a calculated one is not implemented')
    cache = field.table.entry.cache
    fields_ = _cache_fields(cache)
    cache_names = _cache_names(cache)
    for place, other in enumerate(fields_):
        if place != index and is_calculated(other) and index in fields(pieces := stored(cache, place)):
            other.set('formula', encoded(spelled(without(pieces, index), lambda one: cache_names[one], given=False)))
    container = cache.root.find(NS + 'cacheFields')
    assert container is not None
    container.remove(fields_[index])
    container.set('count', str(len(container)))
    for table in _tables(cache):
        own = table.root.find(NS + 'pivotFields')
        assert own is not None
        own.remove(own[index])
        own.set('count', str(len(own)))
        data = table.root.find(NS + 'dataFields')
        if data is not None:
            for entry in list(data):
                place = int(entry.get('fld', '0'))
                if place == index:
                    data.remove(entry)
                elif place > index:
                    entry.set('fld', str(place - 1))
            if len(data):
                data.set('count', str(len(data)))
            else:
                table.root.remove(data)
    _changed(cache)


class CalculatedFields(VBACollection, ExcelObject):
    """A table's calculated fields, in its cache's order."""

    vba_type_name = 'CalculatedFields'

    def __init__(self, table: PivotTable) -> None:
        self.table = table.entry

    def vba_items(self) -> list[object]:
        table = self.table
        views = table.root.findall(f'{NS}pivotFields/{NS}pivotField')
        return [PivotField(table, index, views[index]) for index, field in enumerate(_cache_fields(table.cache))
                if is_calculated(field) and index < len(views)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                if isinstance(item, PivotField) and string_order(str(item.Name()).casefold(), index.casefold()) == 0:
                    return item
            raise error(1004)
        position = int(to_integer(index, 'Long'))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(1004)

    @member
    def Parent(self) -> object:
        return self.table

    @method
    def Add(self, Name: object, Formula: object, UseStandardFormula: object = MISSING) -> object:
        return add(self.table, Name, Formula)


class Measures:
    """Each data field's value over a set of a cache's records: its function's over the values of its field, or a
    calculated field's formula worked out over the sums of the fields it names."""

    def __init__(self, cache_fields: Sequence[XmlElement], sheet: Worksheet) -> None:
        self.cache_fields = list(cache_fields)
        self.sheet = sheet
        self._formulas: dict[int, tuple[list[Piece], Node]] = {}

    def calculated(self, index: int) -> bool:
        return 0 <= index < len(self.cache_fields) and is_calculated(self.cache_fields[index])

    def value(self, report: XmlElement, chosen: Sequence[Sequence[XmlElement]], other: str = '') -> object:
        """A data field's value over ``chosen`` records; ``other`` the function a subtotal line shows instead of the
        data field's own, under which a calculated field shows nothing (tests/fixtures/pivot_calculated_fields/
        count_subtotals)."""
        from pyopenvba.apps.excel._pivot_aggregate import aggregate

        source = int(report.get('fld', '0'))
        if not self.calculated(source):
            return aggregate([record[source] for record in chosen], other or report.get('subtotal', 'sum'))
        if other:
            return EMPTY
        return self._worked_out(source, chosen)

    def _worked_out(self, index: int, chosen: Sequence[Sequence[XmlElement]]) -> object:
        from pyopenvba.apps.excel._pivot_aggregate import aggregate

        pieces, node_ = self._read(index)
        values: dict[int, object] = {}
        for place in fields(pieces):
            if self.calculated(place):
                values[place] = self._worked_out(place, chosen)
            else:
                total = aggregate([record[place] for record in chosen], 'sum')
                values[place] = 0.0 if total is EMPTY else total
        return worked_out(node_, values, self.sheet)

    def _read(self, index: int) -> tuple[list[Piece], Node]:
        if index not in self._formulas:
            names = [decode_text(field.get('name', '')) for field in self.cache_fields]
            try:
                pieces = read(decode_text(self.cache_fields[index].get('formula', '')), _finder(names))
                self._formulas[index] = pieces, engine_node(pieces, self.sheet)
            except Refused:
                raise VBAUnsupportedError('A calculated field whose formula the model cannot read is not '
                                          'implemented') from None
        return self._formulas[index]
