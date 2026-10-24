"""Cached-record PivotTable axes, compact rows and column subtotal layouts."""
from __future__ import annotations

import datetime as dt
from collections.abc import Sequence
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, MAX_COLUMNS, MAX_ROWS, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._pivot_create import node, part_root
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
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.apps.excel._typing import Typed
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.formula._calc import collate
from pyopenvba.interpreter._values import EMPTY, MISSING, error, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable

VALUES = -2
_MONTHS: tuple[Sequence[str], ...] = ('Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.lower().split(),
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


def order(items: list[XmlElement]) -> list[int]:
    text = [str(scalar(entry)).lower() for entry in items if entry.tag == NS + 's']
    custom = next((sequence for sequence in _MONTHS if text and all(value in sequence for value in text)), None)

    def key(index: int) -> tuple[int, object]:
        entry = items[index]
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
    kind: str = 'data'
    header: bool = False


def _lines(axis: list[int], ids: list[list[int]], positions: list[list[int]], data_count: int,
           *, rows: bool, grand: bool, selected: list[int] | None = None) -> list[Line]:
    if not axis:
        return [Line()]
    output: list[Line] = []
    ranks = [{item: rank for rank, item in enumerate(field)} for field in positions]

    def visit(depth: int, selected: list[int], items: tuple[int, ...], keys: tuple[tuple[int, int], ...],
              data: int = 0) -> None:
        field = axis[depth]
        grouped: dict[int, list[int]] = {}
        if field != VALUES:
            for row in selected:
                grouped.setdefault(ids[row][field], []).append(row)
        choices = list(range(data_count)) if field == VALUES else sorted(grouped, key=ranks[field].__getitem__)
        for choice in choices:
            inner = selected if field == VALUES else grouped[choice]
            assigned = items + ((choice if field == VALUES else ranks[field][choice]),)
            constraints = keys if field == VALUES else keys + ((field, choice),)
            measure = choice if field == VALUES else data
            if depth + 1 == len(axis):
                output.append(Line(assigned, constraints, measure))
            else:
                if rows and field != VALUES:
                    output.append(Line(assigned, constraints, measure, header=VALUES in axis[depth + 1:]))
                visit(depth + 1, inner, assigned, constraints, measure)
                if not rows and field != VALUES and any(index != VALUES for index in axis[depth + 1:]):
                    measures = range(data_count) if VALUES in axis[depth + 1:] else [measure]
                    for subtotal in measures:
                        output.append(Line(assigned, constraints, subtotal, 'default'))

    visit(0, selected if selected is not None else list(range(len(ids))), (), ())
    if grand and any(field != VALUES for field in axis):
        for measure in range(data_count) if VALUES in axis else [0]:
            output.append(Line(data=measure, kind='grand'))
    return output


def _line_xml(name: str, lines: list[Line]) -> XmlElement:
    container = node(name, {'count': str(len(lines))})
    for line in lines:
        attributes = {'i': str(line.data)}
        if line.kind != 'data':
            attributes['t'] = line.kind
        item = node('i', attributes)
        for value in line.items:
            item.append(node('x', {'v': str(value)}))
        if not line.items and line.kind == 'grand':
            item.append(node('x'))
        container.append(item)
    return container


def cache_layout(cache: PivotCache, records: list[list[XmlElement]]) -> tuple[XmlElement, XmlElement, list[list[int]], list[list[int]]]:
    from pyopenvba.apps.excel._pivots import refuse_1904_dates

    refuse_1904_dates(cache)
    fields = deepcopy(cache.root.find(NS + 'cacheFields'))
    assert fields is not None
    dictionaries: list[dict[tuple[str, object], int]] = []
    for field in fields:
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
    for field in fields:
        shared = field.find(NS + 'sharedItems')
        assert shared is not None
        types = {entry.tag[len(NS):] for entry in shared}
        # Shared-item summaries classify logicals and errors with strings.
        # Declaring Boolean+text as mixed makes native Excel reject the cache.
        flags = {'s' if kind in ('e', 'b') else kind for kind in types}
        for attribute in ('minValue', 'maxValue', 'minDate', 'maxDate', 'containsInteger'):
            shared.attrib.pop(attribute, None)
        shared.attrib.update(count=str(len(shared)), containsString=str(int('s' in flags)),
                             containsNumber=str(int('n' in types and 'd' not in types)),
                             containsDate=str(int('d' in types)), containsBlank=str(int('m' in types)),
                             containsMixedTypes=str(int(len(flags - {'m'}) > 1)),
                             containsSemiMixedTypes=str(int(bool(flags & {'s', 'b'}))),
                             containsNonDate=str(int(bool(types - {'d', 'm'}))))
        numeric = [float(entry.get('v', '0')) for entry in shared if entry.tag == NS + 'n']
        if numeric and 'd' not in types:
            shared.attrib.update(minValue=repr(min(numeric)), maxValue=repr(max(numeric)),
                                 containsInteger=str(int(all(value.is_integer() for value in numeric))))
        dates = [dt.datetime.fromisoformat(entry.get('v', '')) for entry in shared if entry.tag == NS + 'd']
        if dates:
            shared.attrib.update(minDate=min(dates).isoformat(), maxDate=max(dates).isoformat())
        positions.append(order(list(shared)))
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


def plan_render(table: PivotTable, definition: XmlElement, data: list[XmlElement] | None = None,
                *, cache: PivotCache | None = None, normalized: CacheLayout | None = None,
                cell_limit: int | None = None) -> RenderPlan:
    from pyopenvba.apps.excel._pivot_aggregate import aggregate, hide_values_row, records

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
    lists = carried(root, table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField'), list(fields), on_axes)
    display = [[place for place, _, _ in listing] if listing is not None else positions[index]
               for index, listing in enumerate(lists)]
    chosen_pages = choices(root, lists)
    selected = passing(ids, lists, on_axes, chosen_pages)
    row_lines = _lines(rows, ids, display, len(reports), rows=True, grand=root.get('colGrandTotals', '1') != '0',
                       selected=selected)
    col_lines = _lines(columns, ids, display, len(reports), rows=False, grand=root.get('rowGrandTotals', '1') != '0',
                       selected=selected)
    location = root.find(NS + 'location')
    assert location is not None
    old = parse_area(location.get('ref', 'A1'))
    dr = len(columns) + 1 if any(index != VALUES for index in columns) else 1
    dc = int(bool(rows) or (bool(reports) and bool(columns) and VALUES not in columns))
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

    def put(row: int, column: int, value: object) -> None:
        if value is not EMPTY:
            cells[area.top + row, area.left + column] = value

    captions: dict[tuple[int, int], object] = {}
    subtotals: dict[tuple[int, int], str] = {}

    def caption(field: int, position: int) -> object:
        if field == VALUES:
            return report_names[position]
        key = field, position
        if key not in captions:
            shared = fields[field].find(NS + 'sharedItems')
            assert shared is not None
            captions[key] = shown(shared[display[field][position]])
        return captions[key]

    def subtotal_caption(field: int, position: int) -> str:
        key = field, position
        if key not in subtotals:
            subtotals[key] = str(caption(field, position)) + ' Total'
        return subtotals[key]

    if rows:
        put(dr - 1, 0, 'Values' if rows == [VALUES] else 'Row Labels')
        if not columns and VALUES not in rows and reports:
            put(dr - 1, dc, report_names[0])
    elif not columns:
        for index, name in enumerate(report_names):
            put(0, index, name)
    if any(index != VALUES for index in columns):
        put(0, dc, 'Column Labels')
        if len(reports) == 1:
            put(0 if rows else dr, 0, report_names[0])
    previous: tuple[int, ...] = ()
    for offset, line in enumerate(col_lines):
        if columns:
            if line.kind == 'grand':
                label = report_totals[line.data] if VALUES in columns and reports else 'Grand Total'
                put(1 if dr > 1 else 0, dc + offset, label)
            elif line.kind == 'default':
                label = subtotal_caption(columns[len(line.items) - 1], line.items[-1])
                put(len(line.items), dc + offset, label)
            else:
                for level, item in enumerate(line.items):
                    if level == len(columns) - 1 or previous[:level + 1] != line.items[:level + 1]:
                        put(level + 1 if dr > 1 else 0, dc + offset, caption(columns[level], item))
                previous = line.items
    for offset, line in enumerate(row_lines):
        if rows:
            if line.kind == 'grand':
                label = report_totals[line.data] if VALUES in rows and reports else 'Grand Total'
            else:
                label = caption(rows[len(line.items) - 1], line.items[-1])
            put(dr + offset, 0, label)
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
    buckets: dict[tuple[tuple[tuple[int, int], ...], tuple[tuple[int, int], ...]], list[list[XmlElement]]] = {}
    if reports:
        for record, indices in ((records_[row], ids[row]) for row in selected):
            row_keys = [tuple((index, indices[index]) for index in real_rows[:depth]) for depth in range(len(real_rows) + 1)]
            col_keys = [tuple((index, indices[index]) for index in real_cols[:depth]) for depth in range(len(real_cols) + 1)]
            for row_key in row_keys:
                if row_key not in row_lookup:
                    continue
                for col_key in col_keys:
                    if col_key in col_lookup:
                        buckets.setdefault((row_key, col_key), []).append(record)
        for (row_key, col_key), selected in buckets.items():
            for offset in row_lookup[row_key]:
                line = row_lines[offset]
                for column in col_lookup[col_key]:
                    measure = line.data if VALUES in rows else col_lines[column].data
                    report = reports[measure]
                    value = aggregate([record[int(report.get('fld', '0'))] for record in selected], report.get('subtotal', 'sum'))
                    put(dr + offset, dc + column, value)
    pivot_fields = root.find(NS + 'pivotFields')
    planned_fields = deepcopy(pivot_fields)
    assert pivot_fields is not None and planned_fields is not None
    for offset, field in enumerate(reversed(pages)):
        if over_then_down:
            row_offset, column_offset = divmod(offset, page_columns)
        else:
            column_offset, row_offset = divmod(offset, page_rows)
        put(-page_rows - 1 + row_offset, 3 * column_offset, decode_text(fields[field].get('name', '')))
        showing = page_caption(lists[field] or [], chosen_pages.get(field), flag(pivot_fields[field], MULTIPLE))
        if isinstance(showing, int):
            values = fields[field].find(NS + 'sharedItems')
            assert values is not None
            put(-page_rows - 1 + row_offset, 3 * column_offset + 1, shown(values[showing]))
        else:
            put(-page_rows - 1 + row_offset, 3 * column_offset + 1, showing)
    previous_areas = [old, extent]
    if table.root.findall(f'{NS}pageFields/{NS}pageField'):
        from pyopenvba.apps.excel._pivot_ranges import page_range

        previous_areas.extend(page_range(table).areas)
    for index, (original, planned) in enumerate(zip(pivot_fields, planned_fields, strict=True)):
        attributes = dict(planned.attrib)
        attributes['dataField'] = str(int(any(int(entry.get('fld', '0')) == index for entry in reports)))
        original.attrib.clear()
        original.attrib.update(field_attributes(attributes))
        old_items = original.find(NS + 'items')
        if old_items is not None:
            original.remove(old_items)
        listing = lists[index]
        if listing is not None:
            items = node('items', {'count': str(len(listing) + 1)})
            for place, hide, gone in listing:
                items.append(item_node(place, hide=hide, gone=gone))
            items.append(node('item', {'t': 'default'}))
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
        if not empty:
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
    location.attrib.update(ref=area.address(absolute=False), firstHeaderRow=str(0 if empty and pages else int(VALUES not in columns or len(columns) > 1)),
                           firstDataRow=str(dr), firstDataCol=str(dc))
    if pages:
        location.attrib.update(rowPageCount=str(page_rows), colPageCount=str(page_columns))
    else:
        location.attrib.pop('rowPageCount', None)
        location.attrib.pop('colPageCount', None)
    root.set('dataOnRows', str(int(VALUES in rows)))
    hide_values_row(root)
    return RenderPlan(table, root, fields, cache_records, extent, previous_areas, cells)


def apply_render(plan: RenderPlan, *, cache: PivotCache | None = None) -> None:
    """Apply a precomputed layout without emitting callbacks."""
    table, root = plan.table, plan.root
    cache = cache if cache is not None else table.cache
    Range(table.sheet, plan.previous_areas).clear_contents_checked()
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
        output.store_typed(row, column, Typed(value))
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
            raise error(1004)
        return
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
