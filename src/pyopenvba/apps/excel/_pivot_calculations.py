"""A data field's values shown as one of Excel's calculations, Show Values As, as Excel works them out and keeps them.

PivotField.Calculation shows a data field's values as a difference from a base item, a percent of it or a percent
difference from it, the item being one of the base field's or the one before or after; as a running total over the
base field, plain or as a percent; as a percent of the grand total, of the row's or the column's total, or of the
parent row's, column's or base field's; as an index; or as a rank along the base field. The file keeps the
calculation as the data field's showDataAs, those Excel 2010 added as the pivotShowAs of its extension, the base field
by its place among the cache's fields and the base item by its place in that field's items.

Measured with scripts/measure_pivot_calculations.py (tests/fixtures/pivot_calculations/). A value no record reaches
is 0 in a percent of a total or of a parent and in an index, and counts as 0 in a difference and a running total;
measured against a base item it is #NULL!. A base value no record reaches leaves the cell empty, a base item the axis
does not show there makes it #N/A, and a base field on no axis makes every value #N/A. A line the base field is not
in shows nothing. Ranks are dense, equal values sharing one.
"""
from __future__ import annotations

from collections.abc import Callable, Sequence
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivot_layout import Line
    from pyopenvba.apps.excel._pivots import PivotField

X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
#: Excel 2010's extension of a data field, which keeps the calculations Excel 2010 added.
DATA_EXTENSION = '{E15A36E0-9728-4e99-A89B-3F7291B0FE68}'
#: XlPivotFieldCalculation -> the calculation's name in the file, and whether Excel 2010's extension keeps it.
CALCULATIONS: dict[int, tuple[str, bool]] = {
    -4143: ('normal', False), 2: ('difference', False), 3: ('percent', False), 4: ('percentDiff', False),
    5: ('runTotal', False), 6: ('percentOfRow', False), 7: ('percentOfCol', False), 8: ('percentOfTotal', False),
    9: ('index', False), 10: ('percentOfParentRow', True), 11: ('percentOfParentCol', True),
    12: ('percentOfParent', True), 13: ('percentOfRunningTotal', True), 14: ('rankAscending', True),
    15: ('rankDescending', True)}
CODES = {name: code for code, (name, _) in CALCULATIONS.items()}
SHOW_DATA_AS = {name for name, extended in CALCULATIONS.values() if not extended}
SHOW_AS = {name for name, extended in CALCULATIONS.values() if extended}
#: The calculations that take a base field, and those of them that take one of its items too.
BASE_FIELD = {'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfParent', 'percentOfRunningTotal',
              'rankAscending', 'rankDescending'}
BASE_ITEM = {'difference', 'percent', 'percentDiff'}
PERCENTS = {'percent', 'percentDiff', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'percentOfParentRow',
            'percentOfParentCol', 'percentOfParent', 'percentOfRunningTotal'}
#: The base item a value stands before or after, as the file keeps those two.
PREVIOUS, NEXT = 1048828, 1048829
#: 0.00%, the format a percent calculation shows its values in.
PERCENT_FORMAT = 10
#: A dataField's attributes in the schema's order.
DATA_ORDER = ('name', 'fld', 'subtotal', 'showDataAs', 'baseField', 'baseItem', 'numFmtId')

Keys = tuple[tuple[int, int], ...]


def extension_of(entry: XmlElement) -> XmlElement | None:
    """A dataField's Excel 2010 options, where it has them."""
    for extension in entry.findall(f'{NS}extLst/{NS}ext'):
        if extension.get('uri') == DATA_EXTENSION:
            return extension.find(f'{{{X14}}}dataField')
    return None


def calculation_of(entry: XmlElement) -> str:
    """A dataField's calculation by its name in the file."""
    options = extension_of(entry)
    shown = options.get('pivotShowAs') if options is not None else None
    return shown if shown else entry.get('showDataAs', 'normal')


def _data(field: PivotField) -> XmlElement:
    """A data field's dataField; error 1004 on any other field (tests/fixtures/pivot_calculations/ errors)."""
    if not field.data:
        raise error(1004)
    return field.root


def set_data_attribute(entry: XmlElement, name: str, value: str) -> None:
    """Set a dataField's attribute, its attributes in the schema's order."""
    entry.set(name, value)
    items = sorted(entry.attrib.items(),
                   key=lambda item: DATA_ORDER.index(item[0]) if item[0] in DATA_ORDER else len(DATA_ORDER))
    entry.attrib.clear()
    entry.attrib.update(items)


def _changed(field: PivotField) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    table = field.table.entry
    rerender(table, deepcopy(table.root))
    table.sheet.book.saved = False


def calculation(field: PivotField) -> object:
    return VBAInt(CODES.get(calculation_of(_data(field)), -4143), 'Long')


def set_calculation(field: PivotField, value: object) -> None:
    """Calculation: one of the fifteen, else error 1004. The base field and item stay as they were, and the values'
    format follows the calculation as Excel's does (tests/fixtures/pivot_calculations/ format_normal)."""
    entry = _data(field)
    try:
        chosen = to_integer(value)
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(1004) from None
    if chosen not in CALCULATIONS:
        raise error(1004)
    name, extended = CALCULATIONS[chosen]
    old = calculation_of(entry)
    entry.attrib.pop('showDataAs', None)
    extensions = entry.find(NS + 'extLst')
    if extensions is not None:
        for extension in list(extensions):
            if extension.get('uri') == DATA_EXTENSION:
                extensions.remove(extension)
        if not len(extensions):
            entry.remove(extensions)
    if extended:
        container = entry.find(NS + 'extLst')
        if container is None:
            container = node('extLst')
            entry.append(container)
        extension = node('ext', {'uri': DATA_EXTENSION, 'xmlns:x14': X14})
        extension.append(XmlElement(f'{{{X14}}}dataField', {'pivotShowAs': name}))
        container.append(extension)
    elif name != 'normal':
        set_data_attribute(entry, 'showDataAs', name)
    _number_format(entry, old, name)
    _list_base(field)
    unsort_stepped(field.table.entry.root)
    _changed(field)


def _kind(name: str) -> str:
    if name in PERCENTS:
        return 'percent'
    if name in ('index', 'rankAscending', 'rankDescending'):
        return 'general'
    return 'plain'


def _number_format(entry: XmlElement, old: str, new: str) -> None:
    """A percent calculation shows 0.00% unless the field showed one already, and an index and a rank General, as
    does any other calculation after a percent, an index or a rank; else the field keeps its format
    (tests/fixtures/pivot_calculations/ format_kept, format_index, format_running, percent_then_running,
    own_percent_then_difference, own_format_then_percent_row)."""
    before, after = _kind(old), _kind(new)
    if after == 'percent' and before != 'percent':
        set_data_attribute(entry, 'numFmtId', str(PERCENT_FORMAT))
    elif after == 'general' or (after == 'plain' and before != 'plain'):
        entry.attrib.pop('numFmtId', None)


def _cache_fields(field: PivotField) -> list[XmlElement]:
    return field.table.entry.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')


def base_field(field: PivotField) -> object:
    """BaseField: the base field's name where the calculation takes one, else error 1004."""
    entry = _data(field)
    if calculation_of(entry) not in BASE_FIELD:
        raise error(1004)
    fields = _cache_fields(field)
    index = int(entry.get('baseField', '0'))
    return decode_text(fields[index].get('name', '')) if 0 <= index < len(fields) else ''


def set_base_field(field: PivotField, value: object) -> None:
    """BaseField: a field of the table by name, where the calculation takes one; else error 1004 (tests/fixtures/
    pivot_calculations/ base_before, errors). The base item goes back to the field's first, whatever field it was
    (base_item_same_field, base_previous_new_field)."""
    entry = _data(field)
    if calculation_of(entry) not in BASE_FIELD or isinstance(value, VBAObject):
        raise error(1004)
    wanted = to_text(value).casefold()
    index = next((place for place, one in enumerate(_cache_fields(field))
                  if decode_text(one.get('name', '')).casefold() == wanted), None)
    if index is None:
        raise error(1004)
    set_data_attribute(entry, 'baseField', str(index))
    set_data_attribute(entry, 'baseItem', '0')
    _list_base(field)
    _changed(field)


def _items(field: PivotField, index: int) -> list[int]:
    """The base field's items as the cache places them, in the order the field lists them: its own list, or else
    sorted as it would first list them."""
    from pyopenvba.apps.excel._pivot_filters import listed, shared, sorted_places

    table = field.table.entry
    definitions = table.root.findall(f'{NS}pivotFields/{NS}pivotField')
    fields = _cache_fields(field)
    if not 0 <= index < min(len(definitions), len(fields)):
        return []
    if definitions[index].find(NS + 'items') is not None:
        return [int(item.get('x', '0')) for item in listed(definitions[index])]
    return sorted_places(shared(fields[index]))


def _caption(field: PivotField, index: int, place: int) -> str:
    from pyopenvba.apps.excel._pivot_filters import shared
    from pyopenvba.apps.excel._pivot_layout import item_caption

    entries = shared(_cache_fields(field)[index])
    return item_caption(entries[place], 'General') if 0 <= place < len(entries) else ''


def base_item(field: PivotField) -> object:
    """BaseItem: the base item's name, (previous) or (next), where the calculation takes one, else error 1004."""
    entry = _data(field)
    if calculation_of(entry) not in BASE_ITEM:
        raise error(1004)
    place = int(entry.get('baseItem', '0'))
    if place in (PREVIOUS, NEXT):
        return '(previous)' if place == PREVIOUS else '(next)'
    index = int(entry.get('baseField', '0'))
    items = _items(field, index)
    return _caption(field, index, items[place]) if 0 <= place < len(items) else ''


def set_base_item(field: PivotField, value: object) -> None:
    """BaseItem: an item of the base field by name, (previous) or (next), where the calculation takes one; else
    error 1004."""
    entry = _data(field)
    if calculation_of(entry) not in BASE_ITEM:
        raise error(1004)
    wanted = to_text(value).casefold()
    if wanted in ('(previous)', '(next)'):
        place = PREVIOUS if wanted == '(previous)' else NEXT
    else:
        index = int(entry.get('baseField', '0'))
        _list_base(field)
        found = next((position for position, item in enumerate(_items(field, index))
                      if _caption(field, index, item).casefold() == wanted), None)
        if found is None:
            raise error(1004)
        place = found
    set_data_attribute(entry, 'baseItem', str(place))
    unsort_stepped(field.table.entry.root)
    _changed(field)


def stepping(entry: XmlElement) -> int | None:
    """The field a dataField's values are shown against the item before or after along, or None."""
    if calculation_of(entry) in BASE_ITEM and entry.get('baseItem') in (str(PREVIOUS), str(NEXT)):
        return int(entry.get('baseField', '0'))
    return None


def unsort_stepped(root: XmlElement) -> None:
    """A field sorted by a data field's values while some data field shows its values against the item before or
    after along it: Excel takes the sort away, leaving the field a rankBy of 0, and the field shows in its items'
    order, whichever came first (tests/fixtures/pivot_calculations/ sort_steps_previous, sort_other_previous,
    sort_units_previous, previous_then_sort)."""
    from pyopenvba.apps.excel._pivot_filters import field_attributes
    from pyopenvba.apps.excel._pivot_sorts import set_sort, sort_kind, sort_scope

    stepped = {stepping(entry) for entry in root.findall(f'{NS}dataFields/{NS}dataField')}
    for index, definition in enumerate(root.findall(f'{NS}pivotFields/{NS}pivotField')):
        if index not in stepped or sort_kind(definition) is None or sort_scope(definition) is None:
            continue
        set_sort(definition, None, None)
        attributes = dict(definition.attrib)
        attributes['rankBy'] = '0'
        definition.attrib.clear()
        definition.attrib.update(field_attributes(attributes))


def _list_base(field: PivotField) -> None:
    """A calculation that takes a base item lists its base field's items where the field lists none, as Excel does
    once the field is named (tests/fixtures/pivot_calculations/ difference_off_axis_no_item; running_off_axis
    lists none)."""
    from pyopenvba.apps.excel._pivot_filters import ensure_items

    entry = field.root
    if calculation_of(entry) not in BASE_ITEM:
        return
    table = field.table.entry
    definitions = table.root.findall(f'{NS}pivotFields/{NS}pivotField')
    index = int(entry.get('baseField', '0'))
    if 0 <= index < len(definitions):
        ensure_items(table, definitions[index], index)


def set_number_format(field: PivotField, value: object) -> None:
    """NumberFormat on a data field: the format its values show in."""
    if not field.data:
        raise VBAUnsupportedError('PivotField.NumberFormat can be set on a data field only')
    entry = field.root
    identifier = field.table.entry.sheet.book.stylesheet.format_id(to_text(value))
    if identifier:
        set_data_attribute(entry, 'numFmtId', str(identifier))
    else:
        entry.attrib.pop('numFmtId', None)
    _changed(field)


# -- the values a calculation shows

#: A cell a calculation reads: the items it stands at on the rows and on the columns, and whether it keeps the
#: function of the cell's own row line and own column line, where a subtotal shows another function than the data
#: field's. A total or a parent group the cell reads takes the data field's own function.
Raw = Callable[[Keys, Keys, bool, bool], object]


def calculated(reports: list[XmlElement], rows: list[int], columns: list[int], row_lines: Sequence[Line],
               column_lines: Sequence[Line], display: list[list[int]], records: list[tuple[list[XmlElement], list[int]]],
               function_of: Callable[[Line], str], *, measure_of: Callable[[Line, Line], int]
               ) -> dict[tuple[int, int], object]:
    """The value cells, by row line and column line, that the data fields with a calculation show, worked out from
    ``records``, each with the items it stands at in every field."""
    from pyopenvba.apps.excel._pivot_aggregate import aggregate

    members: dict[tuple[int, int], set[int]] = {}
    for position, (_, items) in enumerate(records):
        for field in (*rows, *columns):
            members.setdefault((field, items[field]), set()).add(position)
    memo: dict[tuple[int, Keys, str], object] = {}

    def total(measure: int, row_keys: Keys, column_keys: Keys, function: str) -> object:
        wanted = (*row_keys, *column_keys)
        key = measure, wanted, function
        if key not in memo:
            places: set[int] | range = range(len(records))
            for one in sorted((members.get(pair, set()) for pair in wanted), key=len):
                places = one if isinstance(places, range) else places & one
            source = int(reports[measure].get('fld', '0'))
            # In the records' order, as the table sums its own values.
            found = [records[place][0][source] for place in sorted(places)]
            memo[key] = aggregate(found, function) if found else EMPTY
        return memo[key]

    axes = Axes([line.keys for line in row_lines]), Axes([line.keys for line in column_lines])
    shown: dict[tuple[int, int], object] = {}
    for offset, line in enumerate(row_lines):
        if line.header:
            continue
        for column, column_line in enumerate(column_lines):
            measure = measure_of(line, column_line)
            if measure >= len(reports) or calculation_of(reports[measure]) == 'normal':
                continue
            own_function = reports[measure].get('subtotal', 'sum')
            row_function, column_function = function_of(line), function_of(column_line)

            def raw(row_keys: Keys, column_keys: Keys, row_own: bool, column_own: bool, measure: int = measure,
                    own_function: str = own_function, row_function: str = row_function,
                    column_function: str = column_function) -> object:
                function = (row_function if row_own else '') or (column_function if column_own else '') or own_function
                return total(measure, row_keys, column_keys, function)

            # A group's own line that shows a subtotal by another function above the lines inside it stands for the
            # group where a percent of the parent row reads it (tests/fixtures/pivot_calculations/
            # count_subtotals_parent_row, count_bottom_parent_row).
            top = bool(line.function) and line.kind == 'data'
            shown[offset, column] = shown_value(reports[measure], rows, columns, line.keys, column_line.keys, raw,
                                                axes, display, top=top)
    return shown


class Axes:
    """An axis's lines by the items they stand at, in the order they show."""

    def __init__(self, lines: list[Keys]) -> None:
        self.lines = set(lines)
        self.items = {pair for keys in lines for pair in keys}
        #: The lines that differ in one place alone, by that place and the items at the others, as they show.
        self.runs: dict[tuple[int, Keys], list[int]] = {}
        for keys in dict.fromkeys(lines):
            for place, (_, item) in enumerate(keys):
                self.runs.setdefault((place, keys[:place] + keys[place + 1:]), []).append(item)

    def run(self, keys: Keys, place: int) -> list[int]:
        """The items the line's field at ``place`` stands at on the lines that differ from it there alone."""
        return self.runs.get((place, keys[:place] + keys[place + 1:]), [])


def shown_value(entry: XmlElement, rows: list[int], columns: list[int], row_keys: Keys, column_keys: Keys,
                raw: Raw, axes: tuple[Axes, Axes], display: list[list[int]], *, top: bool = False) -> object:
    """What one value cell of a dataField shows of its calculation. ``raw`` gives the field's value at any items of
    the row and column fields, EMPTY where no record reaches them; ``axes`` the row and column lines, and
    ``display`` each field's items as its list orders them."""
    name = calculation_of(entry)
    value = raw(row_keys, column_keys, True, True)
    if name == 'percentOfTotal':
        return _ratio(value, raw((), (), False, False))
    if name == 'percentOfRow':
        return _ratio(value, raw(row_keys, (), True, False))
    if name == 'percentOfCol':
        return _ratio(value, raw((), column_keys, False, True))
    if name == 'index':
        # The value times the grand total, over the row's total, over the column's, in that order: Excel's doubles
        # (tests/fixtures/pivot_calculations/ index, count_subtotals_index).
        across, down = _number(raw(row_keys, (), True, False)), _number(raw((), column_keys, False, True))
        if across == 0 or down == 0:
            return ExcelError('#DIV/0!')
        return _number(value) * _number(raw((), (), False, False)) / across / down
    if name == 'percentOfParentRow':
        # The parent of a line at the top is the grand total, and the grand total's is itself (tests/fixtures/
        # pivot_calculations/ parent_row).
        own = raw(row_keys, column_keys, not top, True)
        parent = raw(row_keys[:-1], column_keys, False, True) if row_keys else own
        return EMPTY if parent is EMPTY else _ratio(own, parent)
    if name == 'percentOfParentCol':
        parent = raw(row_keys, column_keys[:-1], True, False) if column_keys else value
        return EMPTY if parent is EMPTY else _ratio(value, parent)
    base = int(entry.get('baseField', '0'))
    on_rows = base in rows
    if not on_rows and base not in columns:
        return ExcelError('#N/A')
    keys = row_keys if on_rows else column_keys
    axis = axes[0] if on_rows else axes[1]
    item = int(entry.get('baseItem', '0'))
    fixed = name in BASE_ITEM and item not in (PREVIOUS, NEXT)
    target = display[base][item] if fixed and 0 <= item < len(display[base]) else -1
    if fixed and (base, target) not in axis.items:
        # A base item the axis does not show makes every value #N/A (tests/fixtures/pivot_calculations/
        # hidden_base_item).
        return ExcelError('#N/A')
    place = next((position for position, (field, _) in enumerate(keys) if field == base), None)
    if place is None:
        return EMPTY

    def moved(other: int) -> Keys:
        return keys[:place] + ((base, other),) + keys[place + 1:]

    def at(other: int) -> object:
        return raw(moved(other), column_keys, True, True) if on_rows else raw(row_keys, moved(other), True, True)

    # The base field's items on the lines that differ from this one in that field alone, as they show: a running
    # total, a rank and the item before or after follow them (tests/fixtures/pivot_calculations/
    # two_rows_difference_previous, sort_inner_running).
    sequence = axis.run(keys, place)
    own = keys[place][1]
    position = sequence.index(own) if own in sequence else 0
    if name == 'percentOfParent':
        ancestor = keys[:place + 1]
        mine = raw(row_keys, column_keys, not top, True) if on_rows else value
        parent = raw(ancestor, column_keys, False, True) if on_rows else raw(row_keys, ancestor, True, False)
        return EMPTY if parent is EMPTY else _ratio(mine, parent)
    if name in ('runTotal', 'percentOfRunningTotal'):
        running = sum(_number(at(other)) for other in sequence[:position + 1])
        if name == 'runTotal':
            return running
        # Out of the value at the line's items but the base field's (tests/fixtures/pivot_calculations/
        # count_subtotals_percent_running).
        rest = keys[:place] + keys[place + 1:]
        whole = raw(rest, column_keys, False, True) if on_rows else raw(row_keys, rest, True, False)
        return _ratio(running, whole)
    if name in ('rankAscending', 'rankDescending'):
        if value is EMPTY:
            return EMPTY
        ranked = sorted({_number(one) for one in map(at, sequence) if one is not EMPTY},
                        reverse=name == 'rankDescending')
        return float(ranked.index(_number(value)) + 1)
    if item == PREVIOUS:
        target = sequence[position - 1] if position > 0 else own
    elif item == NEXT:
        target = sequence[position + 1] if position + 1 < len(sequence) else own
    if target == own:
        return 1.0 if name == 'percent' and value is not EMPTY else EMPTY
    if moved(target) not in axis.lines:
        return ExcelError('#N/A')
    other = at(target)
    if name == 'difference':
        return _number(value) - _number(other)
    if value is EMPTY:
        return ExcelError('#NULL!')
    if other is EMPTY:
        return EMPTY
    if name == 'percent':
        return _ratio(value, other)
    return _ratio(_number(value) - _number(other), other)


def _number(value: object) -> float:
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else 0.0


def _ratio(top: object, bottom: object) -> object:
    for one in (top, bottom):
        if isinstance(one, ExcelError):
            return one
    if _number(bottom) == 0:
        return ExcelError('#DIV/0!')
    return _number(top) / _number(bottom)
