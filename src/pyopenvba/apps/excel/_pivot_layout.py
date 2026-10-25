"""Cached-record PivotTable axes, compact rows and column subtotal layouts."""
from __future__ import annotations

import datetime as dt
from collections.abc import Callable, Sequence
from copy import deepcopy
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, MAX_COLUMNS, MAX_ROWS, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._pivot_calculated import Measures, calculated_view, source_count
from pyopenvba.apps.excel._pivot_calculations import calculated, calculation_of
from pyopenvba.apps.excel._pivot_create import node, part_root
from pyopenvba.apps.excel._pivot_criteria import label_hidden, value_filtered
from pyopenvba.apps.excel._pivot_criteria import rules as filter_rules
from pyopenvba.apps.excel._pivot_filters import (
    MULTIPLE,
    carried,
    choices,
    choose,
    field_attributes,
    flag,
    item_node,
    listed,
    page_caption,
    passing,
)
from pyopenvba.apps.excel._pivot_forms import (
    SUBTOTALS,
    TABLE_ORDER,
    FieldLayout,
    compact_captions,
    compact_indent,
    functions_of,
    label_columns,
    layout_of,
    on,
    set_flag,
)
from pyopenvba.apps.excel._pivot_groups import automatic_dates, items_of, layout_groups, refresh_ranges
from pyopenvba.apps.excel._pivot_items import REFUSED_FUNCTIONS, Workout, has_items, item_field
from pyopenvba.apps.excel._pivot_options import shown_text, values_row_shown as shows_values_row
from pyopenvba.apps.excel._pivot_sorts import value_sorts
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.apps.excel._styles import Alignment, Style, applying
from pyopenvba.apps.excel._typing import Typed
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.formula._calc import collate
from pyopenvba.formula._display import format_value
from pyopenvba.formula._values import ExcelError, number_text
from pyopenvba.interpreter._values import EMPTY, MISSING, VBADate, error, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable

VALUES = -2
#: The item a line stands at for a field it totals over every item of (tests/fixtures/pivot_layouts/ inner_sum_count).
ALL_ITEMS = 1048832
#: The item a line stands at for a field inside a collapsed item, which the line skips (tests/fixtures/
#: pivot_show_detail/ values_rows).
COLLAPSED = -1
_MONTHS:tuple[Sequence[str], ...] = ('Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.lower().split(),
           'January February March April May June July August September October November December'.lower().split(),
           'Sun Mon Tue Wed Thu Fri Sat'.lower().split(),
           'Sunday Monday Tuesday Wednesday Thursday Friday Saturday'.lower().split())


def scalar(entry: XmlElement) -> object:
    kind, raw = entry.tag[len(NS):], decode_text(entry.get('v', ''))
    if kind == 'n':
        return float(raw)
    if kind == 'd':
        return (dt.datetime.fromisoformat(raw) - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    if kind == 'b':
        return raw in ('1', 'true')
    if kind == 'm':
        return EMPTY
    return raw


def item_key(entry: XmlElement) -> tuple[str, object]:
    value = scalar(entry)
    return entry.tag, collate.sort_key(value) if isinstance(value, str) else value


def shown(entry: XmlElement) -> object:
    value = scalar(entry)
    if value is EMPTY:
        return '(blank)'
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    return value


#: The format a date item shows in where its field has none (tests/fixtures/pivot_date_filters/ mixed_format).
DATE_FORMAT = 'm/d/yyyy'


def field_format(codes: dict[int, str], pivot_field: XmlElement | None, cache_field: XmlElement | None) -> str:
    """The format a field shows its items in: its own, else its cache field's, else General."""
    for element in (pivot_field, cache_field):
        identifier = element.get('numFmtId', '0') if element is not None else '0'
        if identifier.isdecimal() and identifier != '0':
            return codes.get(int(identifier), 'General')
    return 'General'


def item_caption(entry: XmlElement, code: str) -> str:
    """An item's name: a number or a date shown in its field's format, a date in m/d/yyyy where the field has none,
    a number in General as VBA spells it (tests/fixtures/pivot_date_filters/)."""
    return label_text(item_label(entry, code))


def label_text(label: object) -> str:
    """The text a label cell shows."""
    if isinstance(label, Typed):
        value = label.value.serial if isinstance(label.value, VBADate) else label.value
        return format_value(value, label.number_format or 'General')
    if isinstance(label, float):
        return number_text(label)
    return str(label)


def item_label(entry: XmlElement, code: str) -> object:
    """What an item's label cell holds: a date as a Date in its field's format, or m/d/yyyy; a number in the field's
    format where it has one; anything else as the value it is (tests/fixtures/pivot_date_filters/ datecell)."""
    value = shown(entry)
    if isinstance(value, float) and entry.tag == NS + 'd':
        return Typed(VBADate(value), code if code != 'General' else DATE_FORMAT)
    if isinstance(value, float) and code != 'General':
        return Typed(value, code)
    return value


def order(items: list[XmlElement]) -> list[int]:
    text = [str(scalar(entry)).lower() for entry in items if entry.tag == NS + 's' and entry.get('f') is None]
    custom = next((sequence for sequence in _MONTHS if text and all(value in sequence for value in text)), None)

    def key(index: int) -> tuple[int, object]:
        entry = items[index]
        if entry.get('f') in ('1', 'true'):
            # Calculated items follow the field's values, in the order they were made (tests/fixtures/
            # pivot_calculated_items/ add_two).
            return 6, index
        kind, value = entry.tag[len(NS):], scalar(entry)
        if kind == 'n':
            return 0, value
        if kind == 's':
            return 1, custom.index(str(value).lower()) if custom is not None else collate.sort_key(str(value))
        if kind == 'b':
            return 2, int(bool(value))
        if kind == 'e':
            return 3, str(value)
        if kind == 'd':
            return 4, value
        return 5, 0

    return sorted(range(len(items)), key=key)


@dataclass
class Line:
    items: tuple[int, ...] = ()
    keys: tuple[tuple[int, int], ...] = ()
    data: int = 0
    #: 'data', 'grand', 'blank', or a subtotal's type: 'default' or one of the functions Subtotals names.
    kind: str = 'data'
    #: A line whose cells show no values: a group's own line where its subtotals show below it, or a blank line.
    header: bool = False
    #: The function a group's own line shows its subtotal by, where that is one Subtotals names.
    function: str = ''


def _lines(axis: list[int], ids: list[list[int]], positions: list[list[int]], data_count: int,
           *, rows: bool, grand: bool, selected: list[int] | None = None,
           sorts: dict[int, tuple[bool, Callable[[list[int]], tuple[int, float, str]]]] | None = None,
           layouts: dict[int, FieldLayout] | None = None) -> list[Line]:
    """An axis's lines in the order they show. A row field in compact or outline form has a line of its own for
    each item, ahead of the lines inside it, which shows the item's subtotal where the field has one alone at the
    top; every other subtotal, and every one of a field in tabular form or on the columns, shows below the lines
    inside the item, one line for each function (tests/fixtures/pivot_layouts/)."""
    if not axis:
        return [Line()]
    output: list[Line] = []
    ranks = [{item: rank for rank, item in enumerate(field)} for field in positions]
    layouts = layouts if layouts is not None else {}

    def visit(depth: int, selected: list[int], items: tuple[int, ...], keys: tuple[tuple[int, int], ...],
              data: int = 0) -> None:
        field = axis[depth]
        grouped: dict[int, list[int]] = {}
        if field != VALUES:
            for row in selected:
                grouped.setdefault(ids[row][field], []).append(row)
        choices = list(range(data_count)) if field == VALUES else sorted(grouped, key=ranks[field].__getitem__)
        if sorts and field in sorts:
            # Sorted by a data field's values: equal ones keep the list's order, and descending turns it round.
            descending, key = sorts[field]
            values = {choice: key(grouped[choice]) for choice in choices}
            choices.sort(key=values.__getitem__)
            if descending:
                choices.reverse()
        for choice in choices:
            inner = selected if field == VALUES else grouped[choice]
            assigned = items + ((choice if field == VALUES else ranks[field][choice]),)
            constraints = keys if field == VALUES else keys + ((field, choice),)
            measure = choice if field == VALUES else data
            if depth + 1 == len(axis):
                output.append(Line(assigned, constraints, measure))
                continue
            layout = layouts.get(field, FieldLayout(functions=()) if field == VALUES else FieldLayout())
            values_below = VALUES in axis[depth + 1:]
            if field != VALUES and choice in layout.collapsed:
                # A collapsed item shows its total alone: on a line of its own, or on one for each data field where
                # the Values field is inside it, the fields between standing at no item (tests/fixtures/
                # pivot_show_detail/ collapse_item, tabular_collapse, values_rows, columns_collapse).
                if values_below:
                    if rows and layout.mode != 'tabular':
                        output.append(Line(assigned, constraints, measure, header=True))
                    gap = axis[depth + 1:].index(VALUES)
                    for each in range(data_count):
                        output.append(Line(assigned + (COLLAPSED,) * gap + (each,), constraints, each))
                else:
                    output.append(Line(assigned, constraints, measure))
                continue
            functions = layout.functions if any(index != VALUES for index in axis[depth + 1:]) else ()
            lone = next(iter(functions)) if len(functions) == 1 else ''
            own_line = rows and layout.mode != 'tabular'
            on_top = own_line and layout.top and bool(lone) and not values_below
            if own_line:
                output.append(Line(assigned, constraints, measure, header=not on_top,
                                   function=lone if on_top and lone in SUBTOTALS else ''))
            visit(depth + 1, inner, assigned, constraints, measure)
            if not on_top:
                for function in functions:
                    for subtotal in range(data_count) if values_below else [measure]:
                        output.append(Line(assigned, constraints, subtotal, function))
            if rows and layout.blank:
                output.append(Line(assigned, constraints, measure, 'blank', header=True))

    visit(0, selected if selected is not None else list(range(len(ids))), (), ())
    inner = axis[-1]
    functions = layouts.get(inner, FieldLayout()).functions if inner != VALUES else ()
    if len(axis) > 1 and functions and functions != ('default',):
        # The innermost field's own subtotals, each item's over every item outside it, show ahead of the grand
        # total, the fields outside it standing at no item (tests/fixtures/pivot_layouts/ inner_sum_count,
        # column_inner_sum_count).
        found = {ids[row][inner] for row in (selected if selected is not None else range(len(ids)))}
        outside = (ALL_ITEMS,) * (len(axis) - 1)
        for choice in sorted(found, key=ranks[inner].__getitem__):
            for function in functions:
                output.append(Line(outside + (ranks[inner][choice],), ((inner, choice),), 0, function))
    if grand and any(field != VALUES for field in axis):
        for measure in range(data_count) if VALUES in axis else [0]:
            output.append(Line(data=measure, kind='grand'))
    return output


def _line_xml(name: str, lines: list[Line]) -> XmlElement:
    """An axis's lines as Excel writes them: each repeating by ``r`` the items it shares with the line before it,
    a kind of data, a first value and an item at 0 left out (tests/fixtures/pivot_sorts/)."""
    container = node(name, {'count': str(len(lines))})
    previous: tuple[int, ...] = ()
    for line in lines:
        repeated = 0
        while (repeated < min(len(previous), len(line.items) - 1)
               and previous[repeated] == line.items[repeated]):
            repeated += 1
        # A field a collapsed item skips counts as repeated (tests/fixtures/pivot_show_detail/ values_rows).
        while repeated < len(line.items) - 1 and line.items[repeated] == COLLAPSED:
            repeated += 1
        attributes = {'t': line.kind} if line.kind != 'data' else {}
        if repeated:
            attributes['r'] = str(repeated)
        if line.data:
            attributes['i'] = str(line.data)
        item = node('i', attributes)
        for value in line.items[repeated:]:
            item.append(node('x', {'v': str(value)} if value else None))
        if not line.items and line.kind == 'grand':
            item.append(node('x'))
        container.append(item)
        previous = line.items
    return container


def cache_layout(cache: PivotCache, records: list[list[XmlElement]]) -> tuple[XmlElement, XmlElement, list[list[int]], list[list[int]]]:
    from pyopenvba.apps.excel._pivots import refuse_1904_dates

    refuse_1904_dates(cache)
    fields = deepcopy(cache.root.find(NS + 'cacheFields'))
    assert fields is not None
    # A calculated field keeps no values, and lists none (tests/fixtures/pivot_calculated_fields/).
    sources = source_count(list(fields))
    dictionaries: list[dict[tuple[str, object], int]] = []
    for field in fields[:sources]:
        shared = field.find(NS + 'sharedItems')
        if shared is None:
            shared = node('sharedItems')
            field.append(shared)
        dictionaries.append({item_key(entry): index for index, entry in enumerate(shared)})
    ids: list[list[int]] = []
    encoded = (deepcopy(cache.records) if cache.records is not None
               else part_root('pivotCacheRecords', relationships=True))
    encoded.set('count', str(len(records)))
    encoded[:] = [entry for entry in encoded if entry.tag != NS + 'r']
    for record in records:
        row: list[int] = []
        output = node('r')
        for column, entry in enumerate(record):
            shared = fields[column].find(NS + 'sharedItems')
            assert shared is not None
            key = item_key(entry)
            index = dictionaries[column].get(key)
            if index is None:
                index = len(shared)
                dictionaries[column][key] = index
                shared.append(deepcopy(entry))
            row.append(index)
            output.append(node('x', {'v': str(index)}))
        ids.append(row)
        encoded.insert(len(ids) - 1, output)
    positions: list[list[int]] = []
    for field in fields[:sources]:
        shared = field.find(NS + 'sharedItems')
        assert shared is not None
        types = {entry.tag[len(NS):] for entry in shared}
        # Shared-item summaries classify logicals and errors with strings.
        # Declaring Boolean+text as mixed makes native Excel reject the cache, and so does a field of numbers with
        # a blank not declared semi-mixed (tests/fixtures/pivot_grouping/ mixed_gappy).
        flags = {'s' if kind in ('e', 'b') else kind for kind in types}
        for attribute in ('minValue', 'maxValue', 'minDate', 'maxDate', 'containsInteger'):
            shared.attrib.pop(attribute, None)
        shared.attrib.update(count=str(len(shared)), containsString=str(int('s' in flags)),
                             containsNumber=str(int('n' in types and 'd' not in types)),
                             containsDate=str(int('d' in types)), containsBlank=str(int('m' in types)),
                             containsMixedTypes=str(int(len(flags - {'m'}) > 1)),
                             containsSemiMixedTypes=str(int(bool(flags & {'s', 'm'}))),
                             containsNonDate=str(int(bool(types - {'d', 'm'}))))
        numeric = [float(entry.get('v', '0')) for entry in shared if entry.tag == NS + 'n']
        if numeric and 'd' not in types:
            shared.attrib.update(minValue=repr(min(numeric)), maxValue=repr(max(numeric)),
                                 containsInteger=str(int(all(value.is_integer() for value in numeric))))
        found = automatic_dates(list(shared))
        if found is not None:
            # Excel's maxDate is the midnight after the latest day, or the latest moment where it has a time
            # (tests/fixtures/pivot_date_filters/, tests/fixtures/pivot_grouping/ mixed_whole, mixed_timed).
            shared.attrib.update(minDate=found[0].isoformat(), maxDate=found[1].isoformat())
        positions.append(order(list(shared)))
    # A grouping by ranges or periods follows its values where it takes its start or end from them, and every
    # grouped field gives each record an item of its own (tests/fixtures/pivot_grouping/).
    refresh_ranges(list(fields))
    layout_groups(list(fields), sources, records, ids, positions)
    return fields, encoded, ids, positions


CacheLayout = tuple[XmlElement, XmlElement, list[list[int]], list[list[int]]]


@dataclass
class RenderPlan:
    table: PivotTable
    root: XmlElement
    fields: XmlElement
    records: XmlElement
    extent: Area
    previous_areas: list[Area]
    cells: dict[tuple[int, int], object]
    #: Each formatted cell's format, in the order Excel makes them.
    styles: list[tuple[tuple[int, int], Style]]


def plan_render(table: PivotTable, definition: XmlElement, data: list[XmlElement] | None = None,
                *, cache: PivotCache | None = None, normalized: CacheLayout | None = None,
                cell_limit: int | None = None) -> RenderPlan:
    from pyopenvba.apps.excel._pivot_aggregate import records

    table = table.entry
    cache = cache if cache is not None else table.cache
    records_ = records(cache)
    root = deepcopy(definition)
    reports = data if data is not None else list(table.root.findall(f'{NS}dataFields/{NS}dataField'))
    report_names = [decode_text(report.get('name', '')) for report in reports]
    report_totals = ['Total ' + name for name in report_names]
    rows = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}rowFields/{NS}field')]
    columns = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}colFields/{NS}field')]
    pages = [int(entry.get('fld', '0')) for entry in reversed(root.findall(f'{NS}pageFields/{NS}pageField'))]
    page_rows = page_columns = 0
    over_then_down = root.get('pageOverThenDown', '0') in ('1', 'true')
    if pages:
        wrap = int(root.get('pageWrap', '0'))
        if over_then_down:
            page_columns = min(wrap, len(pages)) if wrap > 0 else len(pages)
            page_rows = (len(pages) + page_columns - 1) // page_columns
        else:
            page_rows = min(wrap, len(pages)) if wrap > 0 else len(pages)
            page_columns = (len(pages) + page_rows - 1) // page_rows
    if len(reports) < 2:
        rows = [index for index in rows if index != VALUES]
        columns = [index for index in columns if index != VALUES]
    if len(reports) > 1 and VALUES not in rows and VALUES not in columns:
        columns.append(VALUES)
    fields, cache_records, ids, positions = normalized if normalized is not None else cache_layout(cache, records_)
    # Each field's items as the table lists them, carried over to the cache this plan renders from, and the
    # records the hidden items and the page fields' chosen items leave.
    on_axes = {index for index in (*rows, *columns, *pages) if index != VALUES}
    # Each field's items' captions and ShowDetail, by their places in the cache this plan renders from.
    kept: dict[int, dict[int, dict[str, str]]] = {}
    lists = carried(root, table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField'), list(fields), on_axes, kept,
                    positions)
    display = [[place for place, _, _ in listing] if listing is not None else positions[index]
               for index, listing in enumerate(lists)]
    chosen_pages = choices(root, lists)
    # Hidden items and label filters leave their records out first, then value filters, outer fields first.
    found = filter_rules(root)
    selected = passing(ids, lists, on_axes, chosen_pages, label_hidden(found, list(fields), on_axes))
    # Each data field's values, a calculated field's worked out from its formula.
    measures = Measures(list(fields), table.sheet)
    selected = value_filtered(selected, ids, records_, found, (rows, columns), reports, measures)
    # A field's calculated items add rows for each combination of the other fields' items, whose values their
    # formulas work out (tests/fixtures/pivot_calculated_items/).
    workout: Workout | None = None
    shown_fields = [index for index in (*rows, *columns) if index != VALUES]
    items_field = item_field(list(fields), cache, shown_fields)
    if items_field is not None:
        if found or any(calculation_of(report) != 'normal' for report in reports):
            raise VBAUnsupportedError('Filters and Show Values As beside calculated items are not implemented')
        if any(report.get('subtotal') in REFUSED_FUNCTIONS for report in reports):
            raise VBAUnsupportedError('Calculated items beside an average, a standard deviation or a variance are '
                                      'not implemented')
        # A calculated item's formula reaches its field's hidden items too (tests/fixtures/pivot_calculated_items/
        # hide_member).
        open_lists = [[(place, False, gone) for place, _, gone in listing] if index == items_field and listing
                      else listing for index, listing in enumerate(lists)]
        counted = passing(ids, open_lists, on_axes, chosen_pages)
        ids, records_ = [*ids], [*records_]
        workout = Workout(cache, list(fields), items_field, lists[items_field] or [], records_, ids, counted,
                          [index for index in shown_fields if index != items_field], measures, table.sheet)
        selected = workout.extend(selected)
    sorts = value_sorts(root.findall(f'{NS}pivotFields/{NS}pivotField'), on_axes, reports, records_, measures,
                        workout.value if workout is not None else None)
    layouts = {index: layout_of(root, index) for index in (*rows, *columns)}
    for index in layouts:
        if index != VALUES:
            layouts[index] = replace(layouts[index], collapsed=frozenset(
                place for place, extra in kept.get(index, {}).items() if extra.get('sd') in ('0', 'false')))
    # While ManualUpdate is on, the table shows its captions alone, an axis no items, the Values field its data
    # fields' places (tests/fixtures/pivot_options/ manual_pending, manual_two_data).
    skeleton = table.manual
    if skeleton:
        row_lines = ([Line(data=measure) for measure in range(len(reports))] if VALUES in rows
                     else [] if rows else [Line()])
        col_lines = ([Line(data=measure) for measure in range(len(reports))] if VALUES in columns
                     else [Line()])
    else:
        row_lines = _lines(rows, ids, display, len(reports), rows=True,
                           grand=root.get('colGrandTotals', '1') != '0', selected=selected, sorts=sorts,
                           layouts=layouts)
        col_lines = _lines(columns, ids, display, len(reports), rows=False,
                           grand=root.get('rowGrandTotals', '1') != '0', selected=selected, sorts=sorts,
                           layouts=layouts)
    location = root.find(NS + 'location')
    assert location is not None
    old = parse_area(location.get('ref', 'A1'))
    label_of, depth_of = label_columns(root, rows)
    # The Values row shows over the data fields' captions where the Values field is alone on the columns and the
    # table shows it (tests/fixtures/pivot_options/ values_row).
    values_row = columns == [VALUES] and shows_values_row(root)
    dr = len(columns) + 1 if any(index != VALUES for index in columns) or values_row else 1
    dc = label_of[-1] + 1 if rows else int(bool(reports) and bool(columns) and VALUES not in columns)
    if not rows and len(reports) == 1 and any(index != VALUES for index in columns):
        dc = 1
    if not rows and len(reports) > 1 and VALUES in columns:
        dc = 0
    height = dr + len(row_lines)
    width = dc + len(col_lines)
    if not reports:
        height -= int(not rows)
        width -= int(bool(rows) and not columns)
    empty = not rows and not columns and not reports
    if empty:
        dr, dc = (0, 0) if pages else (1, 0)
        height, width = (1, 1) if pages else (18, 3)
    area = Area(old.top, old.left, old.top + max(1, height) - 1, old.left + max(1, width) - 1, table.sheet.name)
    if area.bottom > MAX_ROWS or area.right > MAX_COLUMNS or (pages and area.top - page_rows - 1 < 1):
        raise error(1004)
    extent = Area(area.top - page_rows - 1 if pages else area.top, area.left, area.bottom,
                  max(area.right, area.left + 3 * page_columns - 2) if pages else area.right, table.sheet.name)
    if extent.right > MAX_COLUMNS:
        raise error(1004)
    if cell_limit is not None and extent.rows * extent.columns > cell_limit:
        raise VBAUnsupportedError("Pivot refresh exceeds the supported transaction output budget")
    for other in table.sheet.pivot_tables:
        other_location = other.root.find(NS + 'location')
        if other is not table and other_location is not None:
            block = parse_area(other_location.get('ref', 'A1'))
            if extent.top <= block.bottom and block.top <= extent.bottom and extent.left <= block.right and block.left <= extent.right:
                raise error(1004)
    cells: dict[tuple[int, int], object] = {}
    # The cells' formats, kept apart by kind: Excel formats a table's buttons, then its labels, then its values,
    # and so makes their formats in that order (tests/fixtures/pivot_layouts/ compact).
    buttons: dict[tuple[int, int], Style] = {}
    labels: dict[tuple[int, int], Style] = {}
    values_formats: dict[tuple[int, int], Style] = {}
    base = table.sheet.book.stylesheet.default

    def put(row: int, column: int, value: object) -> None:
        if isinstance(value, Typed):
            value = value.value
        # An empty name or caption leaves its cell empty (tests/fixtures/pivot_names/ data_empty, item_reset).
        if value is not EMPTY and value != '':
            cells[area.top + row, area.left + column] = value

    if empty and not pages:
        # An empty table draws its placeholder again, as a new one does (tests/fixtures/pivot_options/
        # clear_table).
        from pyopenvba.apps.excel._pivot_create import placeholder_border

        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                labels[row, column] = replace(base, border=placeholder_border(area, row, column),
                                              applied=base.applied | {'border'})

    def button(row: int, column: int, value: object) -> None:
        """A field's caption, which shows its button."""
        put(row, column, value)
        buttons[area.top + row, area.left + column] = replace(base, pivot_button=True)

    captions: dict[tuple[int, int], object] = {}
    codes = table.sheet.book.stylesheet.number_formats
    pivot_list = root.findall(f'{NS}pivotFields/{NS}pivotField')
    formats = [field_format(codes, pivot_list[index] if index < len(pivot_list) else None, field)
               for index, field in enumerate(fields)]
    # Every label is aligned left, indented by its depth in its column, where a row field is in compact form, the
    # innermost too (tests/fixtures/pivot_layouts/ compact_off_outer, tabular_compact_on, outline_form_compact).
    aligned = any(layouts[index].mode == 'compact' for index in rows)
    step = compact_indent(root)
    grand_name = decode_text(root.get('grandTotalCaption', 'Grand Total'))
    data_caption = decode_text(root.get('dataCaption', 'Values'))

    def caption(field: int, position: int) -> object:
        """An item's label cell: its caption where it was given one, else its value, a date or a formatted number in
        its field's format (tests/fixtures/pivot_names/ item_caption)."""
        if field == VALUES:
            return report_names[position]
        key = field, position
        if key not in captions:
            given = kept.get(field, {}).get(display[field][position], {}).get('n')
            shared = items_of(fields[field])
            assert shared is not None
            captions[key] = (decode_text(given) if given is not None
                             else item_label(shared[display[field][position]], formats[field]))
        return captions[key]

    def field_name(index: int) -> str:
        if index == VALUES:
            return data_caption
        own = pivot_list[index].get('name') if index < len(pivot_list) else None
        return decode_text(own if own is not None else fields[index].get('name', ''))

    def subtotal_label(field: int, position: int, kind: str, measure: int, values_below: bool) -> str:
        """A subtotal line's label: the field's SubtotalName for an Automatic subtotal, else the item's with Total,
        the data field's name where the Values field is inside it, or the function's (tests/fixtures/
        pivot_layouts/ sum_count, sum_count_values_rows, subtotal_name_bottom, subtotal_name_sum_count)."""
        layout = layouts.get(field, FieldLayout())
        if layout.caption and kind not in SUBTOTALS:
            return layout.caption
        text = label_text(caption(field, position))
        if kind not in SUBTOTALS:
            return f'{text} {report_names[measure]}' if values_below else f'{text} Total'
        name = SUBTOTALS[kind][2]
        if not values_below:
            return f'{text} {name}'
        source = int(reports[measure].get('fld', '0'))
        return f'{text} {name} of {decode_text(fields[source].get("name", ""))}'

    def label_format(code: str, depth: int | None) -> Style:
        """A label cell's format: its field's number format, aligned left and indented by its depth where the
        table aligns its labels (tests/fixtures/pivot_layouts/ indent, dates_compact)."""
        style = base
        if code != 'General':
            style = applying(style, 'number_format', number_format=code)
        if depth is not None:
            style = applying(style, 'alignment', alignment=Alignment(horizontal='left', indent=depth * step))
        return style

    def code_of(field: int, label: object = None) -> str:
        if isinstance(label, Typed):
            return label.number_format or 'General'
        return 'General' if field == VALUES else formats[field]

    def row_label(offset: int, position: int, value: object, code: str) -> None:
        column = label_of[position]
        put(dr + offset, column, value)
        labels[area.top + dr + offset, area.left + column] = label_format(
            code, depth_of[position] if aligned else None)

    def repeated(position: int) -> bool:
        """Whether a row field shows its item's label on every line inside it: where it repeats its labels and
        keeps the fields inside it in other columns (tests/fixtures/pivot_layouts/ repeat_compact)."""
        layout = layouts.get(rows[position], FieldLayout())
        return rows[position] != VALUES and layout.repeat and layout.mode != 'compact'

    captions_shown = on(root, 'showHeaders', True)
    compact_names = compact_captions(root, (*rows, *columns))
    if rows:
        if captions_shown:
            for column in sorted(set(label_of)):
                first = rows[label_of.index(column)]
                if column == 0 and compact_names and rows != [VALUES]:
                    button(dr - 1, 0, decode_text(root.get('rowHeaderCaption', 'Row Labels')))
                else:
                    button(dr - 1, column, field_name(first))
        if not columns and VALUES not in rows and reports:
            put(dr - 1, dc, report_names[0])
    elif not columns:
        for index, name in enumerate(report_names):
            put(0, index, name)
    if any(index != VALUES for index in columns):
        if captions_shown and compact_names:
            button(0, dc, decode_text(root.get('colHeaderCaption', 'Column Labels')))
        elif captions_shown:
            for position, index in enumerate(columns):
                button(0, dc + position, field_name(index))
        if len(reports) == 1:
            if rows:
                button(0, 0, report_names[0])
            else:
                put(dr, 0, report_names[0])
    elif values_row and captions_shown:
        button(0, dc, data_caption)
    previous: tuple[int, ...] = ()
    for offset, line in enumerate(col_lines):
        if columns:
            if line.kind == 'grand':
                label = report_totals[line.data] if VALUES in columns and reports else grand_name
                put(1 if dr > 1 else 0, dc + offset, label)
            elif line.kind != 'data':
                level = len(line.items) - 1
                put(len(line.items), dc + offset, subtotal_label(columns[level], line.items[-1], line.kind, line.data,
                                                                 VALUES in columns[level + 1:]))
            else:
                for level, item in enumerate(line.items):
                    if item == COLLAPSED:
                        continue
                    if level == len(columns) - 1 or previous[:level + 1] != line.items[:level + 1]:
                        label = caption(columns[level], item)
                        row = level + 1 if dr > 1 else 0
                        put(row, dc + offset, label)
                        if code_of(columns[level], label) != 'General':
                            labels[area.top + row, area.left + dc + offset] = label_format(
                                code_of(columns[level], label), None)
                previous = line.items
    previous = ()
    for offset, line in enumerate(row_lines):
        if not rows:
            break
        last = len(line.items) - 1
        if line.kind == 'grand':
            label = report_totals[line.data] if VALUES in rows and reports else grand_name
            row_label(offset, 0, label, code_of(rows[0]))
        elif line.kind == 'blank':
            # A blank line's label cell keeps its field's format, empty (tests/fixtures/pivot_layouts/ blank).
            labels[area.top + dr + offset, area.left + label_of[last]] = label_format(
                code_of(rows[last]), depth_of[last] if aligned else None)
        elif line.kind != 'data':
            for outer in range(last):
                if repeated(outer) and line.items[outer] != ALL_ITEMS:
                    row_label(offset, outer, caption(rows[outer], line.items[outer]),
                              code_of(rows[outer], caption(rows[outer], line.items[outer])))
            row_label(offset, last, subtotal_label(rows[last], line.items[last], line.kind, line.data,
                                                   VALUES in rows[last + 1:]), code_of(rows[last]))
        else:
            for position, item in enumerate(line.items):
                if item == COLLAPSED:
                    continue
                if previous[:position + 1] != line.items[:position + 1] or (repeated(position) and position < last):
                    label = caption(rows[position], item)
                    row_label(offset, position, label, code_of(rows[position], label))
        previous = line.items
    # Populate only observed intersections and their prefix totals. A sparse
    # diagonal must not scan every source row for every blank grid cell.
    row_lookup: dict[tuple[tuple[int, int], ...], list[int]] = {}
    col_lookup: dict[tuple[tuple[int, int], ...], list[int]] = {}
    for offset, line in enumerate(row_lines):
        if not line.header:
            row_lookup.setdefault(line.keys, []).append(offset)
    for offset, line in enumerate(col_lines):
        col_lookup.setdefault(line.keys, []).append(offset)
    real_rows = [index for index in rows if index != VALUES]
    real_cols = [index for index in columns if index != VALUES]
    buckets: dict[tuple[tuple[tuple[int, int], ...], tuple[tuple[int, int], ...]], list[int]] = {}

    def valued(report: XmlElement, picked: list[int], other: str) -> object:
        """A data field's value over a set of the table's rows, calculated items' worked out where it has any."""
        if workout is not None:
            return workout.value(report, picked, other)
        return measures.value(report, [records_[row] for row in picked], other)

    def function_of(line: Line) -> str:
        """The function a line's subtotal takes its data field's values through, where not the field's own."""
        kind = line.kind if line.kind in SUBTOTALS else line.function
        return SUBTOTALS[kind][1] if kind in SUBTOTALS else ''

    def value_format(measure: int) -> Style:
        identifier = reports[measure].get('numFmtId', '0') if measure < len(reports) else '0'
        code = codes.get(int(identifier), 'General') if identifier.isdecimal() and identifier != '0' else 'General'
        return applying(base, 'number_format', number_format=code)

    if reports and not skeleton:
        # Every cell of the values takes its data field's format, an empty one too (tests/fixtures/pivot_layouts/).
        for offset, line in enumerate(row_lines):
            for column, column_line in enumerate(col_lines):
                values_formats[area.top + dr + offset, area.left + dc + column] = value_format(
                    line.data if VALUES in rows else column_line.data)
        rows_over = {line.keys for line in row_lines if line.items and line.items[0] == ALL_ITEMS}
        columns_over = {line.keys for line in col_lines if line.items and line.items[0] == ALL_ITEMS}
        for row in selected:
            indices = ids[row]
            row_keys = [tuple((index, indices[index]) for index in real_rows[:depth]) for depth in range(len(real_rows) + 1)]
            if rows_over and ((real_rows[-1], indices[real_rows[-1]]),) in rows_over:
                row_keys.append(((real_rows[-1], indices[real_rows[-1]]),))
            col_keys = [tuple((index, indices[index]) for index in real_cols[:depth]) for depth in range(len(real_cols) + 1)]
            if columns_over and ((real_cols[-1], indices[real_cols[-1]]),) in columns_over:
                col_keys.append(((real_cols[-1], indices[real_cols[-1]]),))
            for row_key in row_keys:
                if row_key not in row_lookup:
                    continue
                for col_key in col_keys:
                    if col_key in col_lookup:
                        buckets.setdefault((row_key, col_key), []).append(row)
        for (row_key, col_key), bucket in buckets.items():
            for offset in row_lookup[row_key]:
                line = row_lines[offset]
                for column in col_lookup[col_key]:
                    measure = line.data if VALUES in rows else col_lines[column].data
                    report = reports[measure]
                    if not measures.calculated(int(report.get('fld', '0'))):
                        put(dr + offset, dc + column, valued(
                            report, bucket, function_of(line) or function_of(col_lines[column])))
        if any(measures.calculated(int(report.get('fld', '0'))) for report in reports):
            # A calculated field works its formula out in every cell, one no record reaches over zeros
            # (tests/fixtures/pivot_calculated_fields/ margin, ratio).
            chosen = list(selected)
            places: dict[tuple[int, int], set[int]] = {}
            for position, row in enumerate(selected):
                for index in (*real_rows, *real_cols):
                    places.setdefault((index, ids[row][index]), set()).add(position)
            for offset, line in enumerate(row_lines):
                if line.header:
                    continue
                for column, column_line in enumerate(col_lines):
                    report = reports[line.data if VALUES in rows else column_line.data]
                    if not measures.calculated(int(report.get('fld', '0'))):
                        continue
                    reached: set[int] | None = None
                    for pair in (*line.keys, *column_line.keys):
                        reached = set(places.get(pair, set())) if reached is None else reached & places.get(pair, set())
                    found = range(len(chosen)) if reached is None else sorted(reached)
                    value = valued(report, [chosen[place] for place in found],
                                   function_of(line) or function_of(column_line))
                    key = area.top + dr + offset, area.left + dc + column
                    if value is EMPTY:
                        cells.pop(key, None)
                    else:
                        cells[key] = value
        # The cells a calculation works out, which show no null string where it leaves them empty (tests/fixtures/
        # pivot_calculations/ null_string).
        worked_out: set[tuple[int, int]] = set()
        if any(calculation_of(report) != 'normal' for report in reports):
            if rows_over or columns_over:
                raise VBAUnsupportedError("Show Values As beside an inner field's subtotals by another function")
            shown = calculated(reports, real_rows, real_cols, row_lines, col_lines, display,
                               [(records_[row], ids[row]) for row in selected], function_of, measures,
                               measure_of=lambda line, column_line: line.data if VALUES in rows else column_line.data)
            for (offset, column), value in shown.items():
                key = area.top + dr + offset, area.left + dc + column
                worked_out.add(key)
                if value is EMPTY:
                    cells.pop(key, None)
                else:
                    cells[key] = value
        # A value cell no record reaches shows the table's null string, or 0 where it displays none, and one that
        # works out an error its error string where it displays that (tests/fixtures/pivot_options/ null_letter,
        # null_display_off, error_string).
        missing = decode_text(root.get('missingCaption', ''))
        filler = (shown_text(missing) if missing else EMPTY) if on(root, 'showMissing', True) else 0.0
        failed = shown_text(decode_text(root.get('errorCaption', ''))) if on(root, 'showError', False) else None
        for offset, line in enumerate(row_lines):
            if line.header:
                continue
            for column in range(len(col_lines)):
                key = area.top + dr + offset, area.left + dc + column
                if key not in cells and filler is not EMPTY and key not in worked_out:
                    cells[key] = filler
                elif failed is not None and isinstance(cells.get(key), ExcelError):
                    if failed == '':
                        del cells[key]
                    else:
                        cells[key] = failed
    pivot_fields = root.find(NS + 'pivotFields')
    planned_fields = deepcopy(pivot_fields)
    assert pivot_fields is not None and planned_fields is not None
    for offset, field in enumerate(reversed(pages)):
        if over_then_down:
            row_offset, column_offset = divmod(offset, page_columns)
        else:
            column_offset, row_offset = divmod(offset, page_rows)
        # A page field shows by its caption, and its choice by the item's (tests/fixtures/pivot_names/ field_page,
        # item_page).
        button(-page_rows - 1 + row_offset, 3 * column_offset, field_name(field))
        showing = page_caption(lists[field] or [], chosen_pages.get(field), flag(pivot_fields[field], MULTIPLE))
        if isinstance(showing, int):
            values = items_of(fields[field])
            assert values is not None
            given = kept.get(field, {}).get(showing, {}).get('n')
            label = decode_text(given) if given is not None else item_label(values[showing], formats[field])
            put(-page_rows - 1 + row_offset, 3 * column_offset + 1, label)
            if isinstance(label, Typed):
                labels[area.top - page_rows - 1 + row_offset, area.left + 3 * column_offset + 1] = label_format(
                    code_of(field, label), None)
        else:
            put(-page_rows - 1 + row_offset, 3 * column_offset + 1, showing)
    previous_areas = [old, extent]
    # A table moved elsewhere leaves its old place empty (tests/fixtures/pivot_options/ location).
    current = table.root.find(NS + 'location')
    if current is not None and current.get('ref', 'A1') != location.get('ref', 'A1'):
        previous_areas.append(parse_area(current.get('ref', 'A1')))
    if table.root.findall(f'{NS}pageFields/{NS}pageField'):
        from pyopenvba.apps.excel._pivot_ranges import page_range

        previous_areas.extend(page_range(table).areas)
    for index, (original, planned) in enumerate(zip(pivot_fields, planned_fields, strict=True)):
        attributes = dict(planned.attrib)
        attributes['dataField'] = str(int(any(int(entry.get('fld', '0')) == index for entry in reports)))
        original.attrib.clear()
        original.attrib.update(field_attributes(attributes))
        old_items = original.find(NS + 'items')
        # A subtotal's entry keeps the ShowDetail a whole field's took (tests/fixtures/pivot_show_detail/
        # collapse_field).
        entries: list[XmlElement] = list(old_items) if old_items is not None else []
        shown_kinds = {item.get('t'): item.get('sd') for item in entries if item.get('t') not in (None, 'data')}
        if old_items is not None:
            original.remove(old_items)
        listing = lists[index]
        if listing is not None:
            # A field lists an entry for each of its subtotals' types after its items, none where it has no
            # subtotals (tests/fixtures/pivot_layouts/ no_subtotals, all_functions).
            kinds = functions_of(original)
            items = node('items', {'count': str(len(listing) + len(kinds))})
            own = items_of(fields[index]) if index < len(fields) else None
            for place, hide, gone in listing:
                extra = kept.get(index, {}).get(place, {})
                formula = own is not None and place < len(own) and own[place].get('f') in ('1', 'true')
                items.append(item_node(place, hide=hide, gone=gone, caption=extra.get('n'),
                                       collapsed=extra.get('sd') in ('0', 'false'), formula=formula))
            for kind in kinds:
                items.append(node('item', {'t': kind, **({'sd': '0'} if shown_kinds.get(kind) in ('0', 'false')
                                                         else {})}))
            original.insert(0, items)
    for child in list(root):
        if child.tag in {NS + name for name in ('pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields')}:
            root.remove(child)
    children = [pivot_fields]
    for axis, name in ((rows, 'rowFields'), (columns, 'colFields')):
        if axis:
            group = node(name, {'count': str(len(axis))})
            for index in axis:
                group.append(node('field', {'x': str(index)}))
            children.append(group)
        if not empty and not skeleton:
            children.append(_line_xml('rowItems' if name == 'rowFields' else 'colItems', row_lines if name == 'rowFields' else col_lines))
    if pages:
        group = node('pageFields', {'count': str(len(pages))})
        for index in reversed(pages):
            page = node('pageField', {'fld': str(index)})
            choose(page, chosen_pages.get(index))
            group.append(page)
        children.append(group)
    if reports:
        group = node('dataFields', {'count': str(len(reports))})
        group.extend(reports)
        children.append(group)
    for index, child in enumerate(children):
        root.insert(1 + index, child)
    location.attrib.update(ref=area.address(absolute=False), firstHeaderRow=str(0 if empty and pages else int(VALUES not in columns or len(columns) > 1 or values_row)),
                           firstDataRow=str(dr), firstDataCol=str(dc))
    if pages:
        location.attrib.update(rowPageCount=str(page_rows), colPageCount=str(page_columns))
    else:
        location.attrib.pop('rowPageCount', None)
        location.attrib.pop('colPageCount', None)
    set_flag(root, 'dataOnRows', VALUES in rows, False, TABLE_ORDER)
    styles = [*buttons.items(), *labels.items(), *values_formats.items()]
    return RenderPlan(table, root, fields, cache_records, extent, previous_areas, cells, styles)


def unformat(sheet: Worksheet, areas: list[Area]) -> None:
    """Take every cell in a table's old place back to the default format, as Excel does before it lays the table
    out again (tests/fixtures/pivot_layouts/ compact_again)."""
    default = sheet.book.stylesheet.default
    if sum(area.rows * area.columns for area in areas) < len(sheet.cells_):
        keys = {(row, column) for area in areas for row in range(area.top, area.bottom + 1)
                for column in range(area.left, area.right + 1) if (row, column) in sheet.cells_}
    else:
        keys = {key for key in sheet.cells_
                if any(area.top <= key[0] <= area.bottom and area.left <= key[1] <= area.right for area in areas)}
    for row, column in sorted(keys):
        cell = sheet.cells_.get((row, column))
        if cell is not None and cell.style is not None:
            sheet.restyle(row, column, default)


def apply_render(plan: RenderPlan, *, cache: PivotCache | None = None) -> None:
    """Apply a precomputed layout without emitting callbacks."""
    table, root = plan.table, plan.root
    cache = cache if cache is not None else table.cache
    Range(table.sheet, plan.previous_areas).clear_contents_checked()
    unformat(table.sheet, plan.previous_areas)
    # Preserve retained source-field and data-field nodes, not just their values.
    original_fields = table.root.find(NS + 'pivotFields')
    planned_fields = root.find(NS + 'pivotFields')
    assert original_fields is not None and planned_fields is not None
    for original, planned in zip(original_fields, planned_fields, strict=True):
        original.attrib.clear()
        original.attrib.update(planned.attrib)
        original[:] = list(planned)
    root[list(root).index(planned_fields)] = original_fields
    table.root.attrib.clear()
    table.root.attrib.update(root.attrib)
    table.root[:] = list(root)
    old_fields = cache.root.find(NS + 'cacheFields')
    assert old_fields is not None
    index = list(cache.root).index(old_fields)
    cache.root.remove(old_fields)
    cache.root.insert(index, plan.fields)
    cache.records = plan.records
    cache.changed = True
    output = Range(table.sheet, [plan.extent])
    for (row, column), value in plan.cells.items():
        output.store_typed(row, column, value if isinstance(value, Typed) else Typed(value))
    for (row, column), style in plan.styles:
        table.sheet.restyle(row, column, style)
    table.changed = True
    table.report_loaded = False
    table.report_metadata = None
    table.sheet.touched()


def render(table: PivotTable, definition: XmlElement, data: list[XmlElement] | None = None) -> None:
    plan = plan_render(table, definition, data)
    apply_render(plan)
    from pyopenvba.apps.excel import _events

    _events.after_edit(Range(plan.table.sheet, [plan.extent]))
    _events.pivot_updated(plan.table)


def orient(field: object, value: object) -> None:
    from pyopenvba.apps.excel._pivots import PivotField

    assert isinstance(field, PivotField)
    table = field.table.entry
    try:
        # This setter truncates fractional values and treats True as 1.
        chosen = int(value) if isinstance(value, bool) else int(to_number(value))
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(1004) from None
    if chosen not in (0, 1, 2, 3, 4):
        raise error(1004)
    reports = list(table.root.findall(f'{NS}dataFields/{NS}dataField'))
    if field.data and chosen == 4:
        return
    if chosen in (1, 2, 3) and calculated_view(field):
        # A calculated field goes to the data area alone (tests/fixtures/pivot_calculated_fields/ errors).
        raise error(1004)
    if chosen == 3 and field.index >= 0 and has_items(table.cache, field.index):
        # A field with calculated items keeps off the page area (tests/fixtures/pivot_calculated_items/ page_rows).
        raise error(1004)
    if chosen == 4 and field.index != VALUES:
        from pyopenvba.apps.excel._pivot_aggregate import add_data_field

        add_data_field(table, field, MISSING, MISSING, keep_source_axis=True)
        return
    if field.index == VALUES:
        if len(reports) < 2 or chosen not in (0, 1, 2):
            raise error(1004)
        if chosen == 0:
            reports = []
    elif field.data:
        reports = [entry for entry in reports if entry is not field.root]
    elif chosen == int(field.Orientation()):
        # Reassigning the same visible axis preserves its field position.
        if chosen == 0:
            # A field on no axis lists its items first (tests/fixtures/pivot_calculated_items/ hidden_field).
            from pyopenvba.apps.excel._pivot_filters import list_on_reading

            list_on_reading(field)
            raise error(1004)
        return
    if chosen == 3 and any(rule.field == field.index for rule in filter_rules(table.root)):
        raise VBAUnsupportedError('Moving a field with label or value filters to the page area is not implemented')
    root = deepcopy(table.root)
    for name, attribute in (('rowFields', 'x'), ('colFields', 'x'), ('pageFields', 'fld')):
        container = root.find(NS + name)
        if container is not None:
            for entry in list(container):
                if int(entry.get(attribute, '0')) == field.index:
                    container.remove(entry)
            container.set('count', str(len(container)))
    if field.index >= 0:
        fields = root.find(NS + 'pivotFields')
        assert fields is not None
        fields[field.index].attrib.pop('axis', None)
        if chosen:
            fields[field.index].set('axis', {1: 'axisRow', 2: 'axisCol', 3: 'axisPage'}[chosen])
        # A field hiding items comes to the page area showing the rest (tests/fixtures/pivot_filters/ row_to_page).
        if chosen == 3 and any(flag(item, 'h') for item in listed(fields[field.index])):
            fields[field.index].set(MULTIPLE, '1')
    if chosen:
        name = {1: 'rowFields', 2: 'colFields', 3: 'pageFields'}[chosen]
        container = root.find(NS + name)
        if container is None:
            container = node(name)
            root.append(container)
        child = node('pageField', {'fld': str(field.index), 'hier': '-1'}) if chosen == 3 else node('field', {'x': str(field.index)})
        if chosen == 3:
            container.insert(0, child)
        else:
            container.append(child)
        container.set('count', str(len(container)))
    render(table, root, reports)
