"""Worksheet PivotTable record aggregation and populated report construction."""
from __future__ import annotations

import math
import datetime as dt
from copy import deepcopy
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, MAX_COLUMNS, MAX_ROWS, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._typing import Typed
from pyopenvba.apps.excel._pivot_create import encoded, node
from pyopenvba.apps.excel._pivots import NS, PivotField, decode_text, related_parts
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable

FUNCTIONS = {-4157: ('sum', 'Sum'), -4112: ('count', 'Count'), -4113: ('countNums', 'Count'),
             -4106: ('average', 'Average'), -4139: ('min', 'Min'), -4136: ('max', 'Max'),
             -4149: ('product', 'Product'), -4155: ('stdDev', 'StdDev'), -4156: ('stdDevp', 'StdDevp'),
             -4164: ('var', 'Var'), -4165: ('varp', 'Varp')}


def hide_values_row(root: XmlElement) -> None:
    """Preserve the native header geometry when Excel reopens the report."""
    extensions = root.find(NS + 'extLst')
    if extensions is None:
        extensions = node('extLst')
        root.append(extensions)
    identifier = '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}'
    extension = next((entry for entry in extensions if entry.get('uri') == identifier), None)
    if extension is None:
        extension = node('ext', {'uri': identifier})
        extensions.append(extension)
    tag = '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}pivotTableDefinition'
    options = extension.find(tag)
    if options is None:
        options = XmlElement(tag)
        extension.append(options)
    # Without this native default, Excel infers an extra Values-header row for
    # data-field ranges after reopening even though cached cells use two rows.
    options.set('hideValuesRow', '1')


def records(cache: PivotCache) -> list[list[XmlElement]]:
    """Resolve cached shared-item references, never rereading worksheet source cells."""
    root = cache.records
    if root is None:
        package = cache.book.package
        parts = related_parts(package, cache.part, 'richPivotRecords' if cache.rich_records else 'pivotCacheRecords')
        if not parts:
            raise VBAUnsupportedError('Pivot aggregation requires saved cache records')
        cache.records_part = parts[0][1]
        root = xml_tree(package.read(cache.records_part))
        cache.records = root
    fields = cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    shared: list[list[XmlElement]] = []
    for field in fields:
        items = field.find(NS + 'sharedItems')
        shared.append(list(items) if items is not None else [])
    rows: list[list[XmlElement]] = []
    for record in root.findall(NS + 'r'):
        if len(record) != len(fields):
            raise VBAUnsupportedError('This pivot cache record shape is not implemented')
        row: list[XmlElement] = []
        for column, entry in enumerate(record):
            if entry.tag == NS + 'x':
                index = int(entry.get('v', '-1'))
                if not 0 <= index < len(shared[column]):
                    raise VBAUnsupportedError('Invalid pivot shared-item reference')
                entry = shared[column][index]
            if entry.tag[len(NS):] not in ('n', 'm', 'b', 'd', 'e', 's'):
                raise VBAUnsupportedError('This pivot cached value type is not implemented')
            row.append(entry)
        rows.append(row)
    return rows


def aggregate(values: list[XmlElement], function: str) -> object:
    present = [entry for entry in values if entry.tag != NS + 'm']
    if not present:
        return EMPTY
    numbers: list[float] = []
    for entry in present:
        if entry.tag == NS + 'n':
            numbers.append(float(entry.get('v', '0')))
        elif entry.tag == NS + 'd':
            clock = dt.datetime.fromisoformat(entry.get('v', ''))
            numbers.append((clock - dt.datetime(1899, 12, 30)).total_seconds() / 86400)
    if function == 'count':
        return float(len(present))
    if function == 'countNums':
        return float(len(numbers))
    for entry in present:
        if entry.tag == NS + 'e':
            return ExcelError(entry.get('v', '#VALUE!'))
    result = 0.0
    if function in ('sum', 'average'):
        # Python 3.12+ sum compensates cancellation; native pivots accumulate
        # doubles in record order (the 1E16, 1, -1E16 control returns zero).
        for value in numbers:
            result += value
        if function == 'average':
            if not numbers:
                return ExcelError('#DIV/0!')
            result /= len(numbers)
    elif function in ('min', 'max'):
        result = (min(numbers) if function == 'min' else max(numbers)) if numbers else 0.0
    elif function == 'product':
        result = 1.0 if numbers else 0.0
        for value in numbers:
            result *= value
            if not math.isfinite(result):
                return ExcelError('#NUM!')
    else:
        sample = function in ('stdDev', 'var')
        if len(numbers) < (2 if sample else 1):
            return ExcelError('#DIV/0!')
        mean = squares = 0.0
        for count, value in enumerate(numbers, 1):
            delta = value - mean
            mean += delta / count
            squares += delta * (value - mean)
        result = squares / (len(numbers) - int(sample))
        if not math.isfinite(result):
            return ExcelError('#NUM!')
        if function in ('stdDev', 'stdDevp'):
            result = math.sqrt(max(0.0, result))
    return (0.0 if result == 0 else result) if math.isfinite(result) else ExcelError('#NUM!')


def add_data_field(table: PivotTable, source: object, caption: object, function: object,
                   *, keep_source_axis: bool = False) -> PivotField:
    table = table.entry
    if int(table.cache.SourceType()) != 1:
        raise VBAUnsupportedError('Aggregating external or consolidated pivot sources is not implemented')
    axes = table.root.findall(f'{NS}rowFields/{NS}field') + table.root.findall(f'{NS}pageFields/{NS}pageField')
    axes += [entry for entry in table.root.findall(f'{NS}colFields/{NS}field') if entry.get('x') != '-2']
    if isinstance(source, Range):
        index = 0
    elif isinstance(source, PivotField):
        if source.table.cache is not table.cache or source.index < 0:
            raise error(1004)
        index = source.index
    else:
        raise error(13)
    rows = records(table.cache)
    fields = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    if not 0 <= index < len(fields):
        raise error(1004)
    values = [row[index] for row in rows]
    if function is MISSING or function is NULL:
        numeric = any(entry.tag == NS + 'n' for entry in values)
        numeric = numeric and all(entry.tag in (NS + 'n', NS + 'm') for entry in values)
        chosen = -4157 if numeric else -4112
    elif not isinstance(function, (int, float, Decimal)) or isinstance(function, bool):
        raise error(5)
    elif function not in FUNCTIONS:
        raise error(9)
    else:
        chosen = int(function)
    kind, label = FUNCTIONS[chosen]
    data = table.root.find(NS + 'dataFields')
    previous = list(data) if data is not None else []
    source_name = decode_text(fields[index].get('name', ''))
    used = {decode_text(entry.get('name', '')).casefold() for entry in [*fields, *previous]}
    if caption is MISSING or caption is NULL:
        base = label + ' of ' + source_name
        name, suffix = base, 2
        while name.casefold() in used:
            name, suffix = base + str(suffix), suffix + 1
    elif not isinstance(caption, str):
        raise error(5)
    else:
        name = caption
        if name.casefold() in used:
            raise error(1004)
    if len(name.encode('utf-16-le', 'surrogatepass')) // 2 > 255:
        raise error(5)
    new = node('dataField', {'name': encoded(name), 'fld': str(index), 'subtotal': kind,
                             'baseField': '0', 'baseItem': '0'})
    reports = [*previous, new]
    if axes:
        from pyopenvba.apps.excel._pivot_layout import render

        definition = deepcopy(table.root)
        if not keep_source_axis:
            # AddDataField moves its source off a label/filter axis. Assigning
            # Orientation=xlDataField instead leaves that source axis intact.
            for axis, attribute in (('rowFields', 'x'), ('colFields', 'x'), ('pageFields', 'fld')):
                container = definition.find(NS + axis)
                if container is not None:
                    container[:] = [entry for entry in container if int(entry.get(attribute, '0')) != index]
                    container.set('count', str(len(container)))
            pivot_fields = definition.find(NS + 'pivotFields')
            assert pivot_fields is not None
            pivot_fields[index].attrib.pop('axis', None)
        render(table, definition, reports)
        return PivotField(table, index, new, True)
    results = [aggregate([row[int(entry.get('fld', '0'))] for row in rows], entry.get('subtotal', 'sum'))
               for entry in reports]
    location = table.root.find(NS + 'location')
    assert location is not None
    old = parse_area(location.get('ref', 'A1'))
    area = Area(old.top, old.left, old.top + 1, old.left + len(reports) - 1, table.sheet.name)
    if area.bottom > MAX_ROWS or area.right > MAX_COLUMNS:
        raise error(1004)
    for other in table.sheet.pivot_tables:
        other_location = other.root.find(NS + 'location')
        if other is not table and other_location is not None:
            block = parse_area(other_location.get('ref', 'A1'))
            if area.top <= block.bottom and block.top <= area.bottom and area.left <= block.right and block.left <= area.right:
                raise error(1004)
    cleared = Range(table.sheet, [old, area])
    cleared.clear_contents_checked()
    for child in list(table.root):
        if child.tag in {NS + part for part in ('rowItems', 'colFields', 'colItems', 'dataFields')}:
            table.root.remove(child)
    pivot_fields = table.root.find(NS + 'pivotFields')
    assert pivot_fields is not None
    data_indexes = {int(entry.get('fld', '0')) for entry in reports}
    for column, field in enumerate(pivot_fields):
        field.set('dataField', str(int(column in data_indexes)))
    children = [node('rowItems', {'count': '1'})]
    children[0].append(node('i'))
    if len(reports) > 1:
        columns = node('colFields', {'count': '1'})
        columns.append(node('field', {'x': '-2'}))
        children.append(columns)
    columns = node('colItems', {'count': str(len(reports))})
    for value_index in range(len(reports)):
        item = node('i', {'i': str(value_index)})
        if len(reports) > 1:
            item.append(node('x', {'v': str(value_index)}))
        columns.append(item)
    children.append(columns)
    data = node('dataFields', {'count': str(len(reports))})
    data.extend(reports)
    children.append(data)
    position = list(table.root).index(pivot_fields) + 1
    for offset, child in enumerate(children):
        table.root.insert(position + offset, child)
    hide_values_row(table.root)
    location.attrib.update(ref=area.address(absolute=False), firstHeaderRow=str(int(len(reports) == 1)),
                           firstDataRow='1', firstDataCol='0')
    output = Range(table.sheet, [area])
    for offset, (entry, result) in enumerate(zip(reports, results)):
        heading = decode_text(entry.get('name', ''))
        output.store_typed(area.top, area.left + offset, Typed(heading if heading else EMPTY))
        output.put_as_value(area.bottom, area.left + offset, result)
    table.changed = True
    table.report_loaded = False
    table.report_metadata = None
    table.sheet.touched()
    from pyopenvba.apps.excel import _events

    _events.after_edit(output)
    _events.pivot_updated(table)
    return PivotField(table, index, new, True)
