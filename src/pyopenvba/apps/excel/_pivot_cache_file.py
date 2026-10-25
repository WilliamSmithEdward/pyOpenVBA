"""A pivot cache as Excel saves it: its fields' shared items and the records beside them.

The model keeps every field's values as shared items, each record naming them by place. Excel's saved caches,
every one of 2,477 fields in the tests' fixtures, list a field's items only where a table of the cache lists
items for the field, a field it puts or has put on its rows, columns or page area, and give the other fields'
values in the records themselves. A field's sharedItems carries only the attributes that differ from the schema's
defaults, in the schema's order, and the count where it lists items; numbers are spelled in the fewest digits
that read back the same, a whole one without a point, an exponent's E in upper case (tests/fixtures/
pivot_grouping/, pivot_layouts/ and the rest).
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._pivots import NS

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._pivots import PivotCache

#: A sharedItems' attributes in the schema's order, and those with defaults, which Excel leaves out.
ORDER = ('containsSemiMixedTypes containsNonDate containsDate containsString containsBlank containsMixedTypes '
         'containsNumber containsInteger minValue maxValue minDate maxDate count longText').split()
DEFAULTS = {'containsSemiMixedTypes': '1', 'containsNonDate': '1', 'containsDate': '0', 'containsString': '1',
            'containsBlank': '0', 'containsMixedTypes': '0', 'containsNumber': '0', 'containsInteger': '0',
            'longText': '0'}


def number_text(value: float) -> str:
    """A cached number as Excel spells it: the fewest digits that read back the same, a whole one without a point,
    the exponent's E in upper case."""
    text = repr(float(value))
    if text.endswith('.0'):
        text = text[:-2]
    return '0' if text == '-0' else text.replace('e', 'E')


def _respelled(entry: XmlElement) -> XmlElement:
    """A value as Excel writes it, a number in its spelling."""
    made = deepcopy(entry)
    if made.tag == NS + 'n' and made.get('v') is not None:
        made.set('v', number_text(float(made.get('v', '0'))))
    return made


def summary(values: list[XmlElement], *, listed: bool) -> dict[str, str]:
    """The attributes of a field's sharedItems for its distinct values, as Excel writes them."""
    from pyopenvba.apps.excel._pivot_groups import automatic_dates

    types = {entry.tag[len(NS):] for entry in values}
    # Logicals and errors count with text (see _pivot_layout.cache_layout).
    flags = {'s' if kind in ('e', 'b') else kind for kind in types}
    attributes = {'containsSemiMixedTypes': str(int(bool(flags & {'s', 'm'}))),
                  'containsNonDate': str(int(bool(types - {'d', 'm'}))),
                  'containsDate': str(int('d' in types)), 'containsString': str(int('s' in flags)),
                  'containsBlank': str(int('m' in types)), 'containsMixedTypes': str(int(len(flags - {'m'}) > 1)),
                  'containsNumber': str(int('n' in types and 'd' not in types))}
    numbers = [float(entry.get('v', '0')) for entry in values if entry.tag == NS + 'n']
    if numbers and 'd' not in types:
        attributes['containsInteger'] = str(int(all(value.is_integer() for value in numbers)))
        attributes['minValue'] = number_text(min(numbers))
        attributes['maxValue'] = number_text(max(numbers))
    found = automatic_dates(values)
    if found is not None:
        attributes['minDate'], attributes['maxDate'] = found[0].isoformat(), found[1].isoformat()
    if listed:
        attributes['count'] = str(len(values))
    if any(entry.tag == NS + 's' and len(entry.get('v', '')) > 255 for entry in values):
        attributes['longText'] = '1'
    return {name: attributes[name] for name in ORDER if name in attributes and DEFAULTS.get(name) != attributes[name]}


def listed_fields(cache: PivotCache, book: Workbook) -> set[int]:
    """The fields a table of the cache lists items for."""
    listed: set[int] = set()
    for sheet in book.sheets_:
        for table in sheet.pivot_tables:
            if table.cache is cache:
                for index, field in enumerate(table.root.findall(f'{NS}pivotFields/{NS}pivotField')):
                    if field.find(NS + 'items') is not None:
                        listed.add(index)
    return listed


def saved(cache: PivotCache, book: Workbook) -> tuple[XmlElement, XmlElement | None]:
    """The cache's definition and records as Excel saves them; a cache the model does not lay out, a rich one or one
    not over cells, as it stands."""
    root, records = deepcopy(cache.root), deepcopy(cache.records) if cache.records is not None else None
    source = root.find(NS + 'cacheSource')
    fields = root.find(NS + 'cacheFields')
    if cache.rich_records or fields is None or source is None or source.get('type', 'worksheet') != 'worksheet':
        return root, records
    rows = [entry for entry in records if entry.tag == NS + 'r'] if records is not None else []
    listed = listed_fields(cache, book)
    for index, field in enumerate(fields):
        shared = field.find(NS + 'sharedItems')
        if shared is None or field.get('databaseField') in ('0', 'false'):
            continue
        entries = list(shared)
        if not entries:
            if records is None or not rows or not all(index < len(row) for row in rows):
                continue
            values = list({(entry.tag, tuple(sorted(entry.attrib.items()))): entry
                           for entry in (row[index] for row in rows)}.values())
        else:
            values = entries
        keep = index in listed and bool(entries)
        if not keep and entries:
            if records is None:
                # Records the model never read stay as the file had them, and the items they name with them.
                continue
            for row in rows:
                if index < len(row) and row[index].tag == NS + 'x':
                    place = int(row[index].get('v', '0'))
                    row[index] = _respelled(entries[place]) if 0 <= place < len(entries) else row[index]
        rest = {key: value for key, value in shared.attrib.items() if key not in ORDER}
        shared.attrib.clear()
        shared.attrib.update(summary(values, listed=keep))
        shared.attrib.update(rest)
        shared[:] = [_respelled(entry) for entry in entries] if keep else []
    for row in rows:
        for position, entry in enumerate(row):
            if entry.tag == NS + 'n':
                row[position] = _respelled(entry)
    return root, records
