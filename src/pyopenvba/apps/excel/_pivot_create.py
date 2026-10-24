"""Worksheet-source PivotCache creation, materialization and empty report parts."""
from __future__ import annotations

import datetime as dt
import math
import posixpath
import re
import uuid
from collections.abc import Callable
from copy import copy, deepcopy
from dataclasses import replace
from decimal import Decimal
from html import escape
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, split_sheet
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._pivots import NS, REL, PivotCache, PivotTable
from pyopenvba.apps.excel._xstring import encode_attribute as encoded
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._display import format_value
from pyopenvba.formula._r1c1 import from_a1, to_a1
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBADate,
    VBAErrorValue,
    error,
    to_bool,
    to_integer,
    to_text,
)
from pyopenvba.powerquery._sheets import add_content_type, add_relationship

if TYPE_CHECKING:
    from pyopenvba.apps.excel._dynamic_file import RichError
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._tables import Table
    from pyopenvba.powerquery._opc import OpcFile

MAIN = NS[1:-1]
LINK = REL[1:-1] + '/'


def node(name: str, attributes: dict[str, str] | None = None) -> XmlElement:
    return XmlElement(NS + name, attributes or {})


def part_root(name: str, attributes: dict[str, str] | None = None, *, relationships: bool = False) -> XmlElement:
    """A part's root, declaring SpreadsheetML's namespace its default as Excel writes it, and the relationships'
    prefix where Excel declares it, a pivot cache's definition and records (tests/fixtures/dxf_table/order.xlsx)."""
    declared = {'xmlns': MAIN, **({'xmlns:r': REL[1:-1]} if relationships else {})}
    return node(name, {**declared, **(attributes or {})})


def table_source(book: Workbook, text: str) -> Table | None:
    owner, bare = split_sheet(text)
    name = bare.split('[', 1)[0]
    return next((table for sheet in book.sheets_ for table in sheet.tables
                 if table.name.casefold() == name.casefold()
                 and (not owner or owner.casefold() == sheet.name.casefold())), None)


def a1_text(text: str) -> str:
    """A source's text with an R1C1 reference, as a recorded macro gives one, spelled in A1."""
    if re.search(r'(?:^|!)R(?:\d+|\[-?\d+\])C(?:\d+|\[-?\d+\])', text, re.I):
        return to_a1('=' + text, 1, 1)[1:]
    return text


def source_range(book: Workbook, source: object, *, cache_source: bool = False) -> Range:
    if isinstance(source, Range):
        if source.sheet.book is not book:
            raise VBAUnsupportedError('Pivot caches from another workbook are not implemented')
        return source
    text = a1_text(to_text(source))
    sheet = book.active_sheet
    if sheet is None:
        raise error(1004)
    owner, bare = split_sheet(text)
    context = next((one for one in book.sheets_ if owner and one.name.casefold() == owner.casefold()), None)
    # A source naming its sheet is read there, whichever sheet is active.
    if context is not None:
        sheet = context
    if cache_source:
        if context is not None:
            named = book.names_.find(bare, scope=context)
            if named is not None and '!' not in to_text(named.Name()):
                sheet, text = context, bare
    result = sheet.Range(text)
    if not isinstance(result, Range):
        raise error(1004)
    table = table_source(book, text) if cache_source else None
    if table is not None and table.headers and len(result.areas) == 1:
        # Pivot sources include column captions even when a structured reference
        # names the data body. Table totals never become cache records.
        selected = result.first
        headers_only = '#headers' in text.casefold() and '#data' not in text.casefold() and '#all' not in text.casefold()
        totals_only = '#totals' in text.casefold() and '#data' not in text.casefold() and '#all' not in text.casefold()
        if not headers_only and not totals_only:
            result = Range(table.sheet, [Area(table.area.top, selected.left,
                           min(selected.bottom, table.area.bottom - int(table.totals)), selected.right, table.sheet.name)])
    return result


def create_cache(book: Workbook, kind: object, source: object, version: object) -> PivotCache:
    source_type = int(to_integer(kind, 'Long'))
    if source_type in (2, 3):
        raise VBAUnsupportedError('PivotCaches.Create from external/consolidated sources is not implemented')
    if source_type != 1 or source is MISSING or source is EMPTY or source is NULL or source is NOTHING:
        raise error(5)
    chosen = 8 if version is MISSING else int(to_integer(version, 'Long'))
    if not 0 <= chosen <= 8:
        raise error(5)
    if not isinstance(source, (str, Range)):
        raise error(5)
    root = part_root('pivotCacheDefinition', {'createdVersion': str(chosen), 'refreshedVersion': str(chosen),
                                           'minRefreshableVersion': str(min(chosen, 3)), 'recordCount': '0'},
                     relationships=True)
    origin = node('cacheSource', {'type': 'worksheet'})
    root.append(origin)
    raw = to_text(source) if isinstance(source, str) else ''
    try:
        selected = source_range(book, source, cache_source=True)
    except VBARuntimeError:
        selected = None
    if selected is not None and len(selected.areas) == 1:
        area = selected.first
        origin.append(node('worksheetSource', {'ref': area.address(absolute=False), 'sheet': selected.sheet.name}))
        raw = from_a1('=' + area.address(absolute=True), 1, 1)[1:]
        sheet = selected.sheet.name
        if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*', sheet):
            sheet = "'" + sheet.replace("'", "''") + "'"
        raw = sheet + '!' + raw
        if isinstance(source, str):
            table = table_source(book, source)
            owner, bare = split_sheet(source)
            named = book.names_.find(source, scope=selected.sheet)
            if named is None and owner:
                named = book.names_.find(bare, scope=selected.sheet)
            if table is not None or named is not None:
                raw = source
                if table is None and named is not None and owner:
                    raw = to_text(named.Name())
                stored = table.name if table is not None and ('#all' in source.casefold() or '#data' in source.casefold()) else raw
                origin[0].attrib.clear()
                origin[0].set('name', stored)
    elif isinstance(source, Range):
        areas = source.areas
        raw = areas[0].address(absolute=True) + ',' + ','.join(
            f'[{book.name}]{source.sheet.name}!{area.address(absolute=True)}' for area in areas[1:])
    cache = PivotCache(book, '', root, 0)
    cache.pending_source = source
    cache.source_text = raw
    return cache


def _materialize(cache: PivotCache, charge_text: Callable[[object], None] | None = None) -> None:
    source = cache.pending_source
    if source is None:
        origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
        if origin is None:
            raise VBAUnsupportedError('Refreshing this pivot source is not implemented')
        source = origin.get('name') or (origin.get('sheet', '') + '!' + origin.get('ref', ''))
    try:
        selected = source_range(cache.book, source, cache_source=True)
    except VBARuntimeError:
        raise error(1004) from None
    if len(selected.areas) != 1 or selected.first.rows < 2:
        raise error(1004)
    area = selected.first
    captions: list[str] = []
    for column in range(area.left, area.right + 1):
        cell = Range(selected.sheet, [Area(area.top, column, area.top, column, selected.sheet.name)])
        value = cell.Value2()
        if charge_text is not None:
            charge_text(value)
        if value is EMPTY:
            caption = ''
        elif isinstance(value, VBAErrorValue):
            from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS
            caption = next((name for name, number in ERROR_NUMBERS.items() if number == value.number), '#VALUE!')
        else:
            caption = format_value(value, selected.sheet.style_at(area.top, column).number_format,
                                   epoch_1904=selected.sheet.book.epoch_1904)
        captions.append(caption)
    reserved = {caption.casefold() for caption in captions if caption}
    used: set[str] = set()
    headers: list[str] = []
    suffixes: dict[str, int] = {}
    for index, caption in enumerate(captions, 1):
        base = caption or f'Field{index}'
        name = base
        suffix = suffixes.get(base.casefold(), 2)
        while name.casefold() in used or (not caption and name.casefold() in reserved):
            name = base + str(suffix)
            suffix += 1
            while name.casefold() in reserved:
                name = base + str(suffix)
                suffix += 1
        suffixes[base.casefold()] = suffix
        used.add(name.casefold())
        headers.append(name)
    fields = node('cacheFields', {'count': str(len(headers))})
    for offset, header in enumerate(headers):
        # A field takes the format its source's cells all share, else General (tests/fixtures/pivot_date_filters/
        # money, mixed_format); the cells' formats also decide the typed records below.
        codes = {selected.sheet.style_at(row, area.left + offset).number_format
                 for row in range(area.top + 1, area.bottom + 1)}
        shared_code = next(iter(codes)) if len(codes) == 1 else 'General'
        identifier = cache.book.stylesheet.format_id(shared_code) if shared_code != 'General' else 0
        field = node('cacheField', {'name': encoded(header), 'numFmtId': str(identifier)})
        field.append(node('sharedItems'))
        fields.append(field)
    records = part_root('pivotCacheRecords', {'count': str(area.rows - 1)}, relationships=True)
    kinds: list[set[str]] = [set() for _ in headers]
    numbers: list[list[float]] = [[] for _ in headers]
    cache.rich_records = False
    rich_columns: dict[str, list[str | None]] = {}
    for row in range(area.top + 1, area.bottom + 1):
        entry = node('r')
        for column in range(area.left, area.right + 1):
            cell = selected.sheet.cell(row, column)
            value = (cache.book.calculator.value_of(selected.sheet.name, row, column) if cell is not None and cell.formula
                     else cell.value if cell is not None else EMPTY)
            if charge_text is not None:
                charge_text(value)
            stored = record_value(value, selected.sheet.style_at(row, column).number_format)
            if stored.tag == NS + 's' and isinstance(value, (VBAErrorValue, ExcelError)):
                cache.rich_records = True
                rich_column = rich_columns.setdefault(headers[column - area.left], [])
                rich_column.extend([None] * (row - area.top - 1 - len(rich_column)))
                rich_column.append(stored.get('v'))
            entry.append(stored)
            kind = stored.tag[len(NS):]
            kinds[column - area.left].add(kind)
            if kind == 'n':
                numbers[column - area.left].append(float(stored.get('v', '0')))
        records.append(entry)
    for index, field in enumerate(fields):
        shared = field[0]
        types = kinds[index]
        shared.attrib.update({'containsString': str(int('s' in types)), 'containsNumber': str(int('n' in types)),
                              'containsDate': str(int('d' in types)), 'containsBlank': str(int('m' in types)),
                              'containsNonDate': str(int(bool(types - {'d', 'm'}))),
                              'containsMixedTypes': str(int(len(types - {'m'}) > 1)),
                              'containsSemiMixedTypes': str(int(bool(types & {'s', 'b', 'e'})))})
        # Numeric bounds are invalid for a field containing date records: Excel
        # refuses the workbook rather than ignoring the contradictory summary.
        if numbers[index] and 'd' not in types:
            shared.set('containsInteger', str(int(all(value.is_integer() for value in numbers[index]))))
            shared.set('minValue', repr(min(numbers[index])))
            shared.set('maxValue', repr(max(numbers[index])))
    old = cache.root.find(NS + 'cacheFields')
    position = list(cache.root).index(old) if old is not None else len(cache.root)
    if old is not None:
        cache.root.remove(old)
    cache.root.insert(position, fields)
    cache.root.set('recordCount', str(area.rows - 1))
    cache.records = records
    if cache.rich_records:
        cache.rich_guid = '{' + str(uuid.uuid4()).upper() + '}'
        for values in rich_columns.values():
            values.extend([None] * (area.rows - 1 - len(values)))
        cache.rich_columns = rich_columns
    origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
    if origin is not None and origin.get('name') is not None:
        cache.source_text = origin.get('name')
    cache.changed = True


def plan_materialize(cache: PivotCache, *, source: Range | None = None, charge_text: Callable[[object], None] | None = None) -> PivotCache:
    """Read the source into a detached cache snapshot, without publishing it."""
    from pyopenvba.apps.excel._pivots import refuse_1904_dates

    planned = copy(cache)
    planned.root = deepcopy(cache.root)
    planned.records = None
    if source is not None:
        planned.pending_source = source
    _materialize(planned, charge_text)
    refuse_1904_dates(planned)
    return planned


def apply_materialized(cache: PivotCache, planned: PivotCache) -> None:
    """Publish a fully admitted snapshot while retaining the cache/root identity."""
    cache.root.attrib.clear()
    cache.root.attrib.update(planned.root.attrib)
    cache.root[:] = list(planned.root)
    cache.records = planned.records
    cache.records_part = planned.records_part
    cache.rich_records = planned.rich_records
    cache.rich_guid = planned.rich_guid
    cache.rich_columns = planned.rich_columns
    cache.source_text = planned.source_text
    cache.changed = True


def materialize(cache: PivotCache) -> None:
    apply_materialized(cache, plan_materialize(cache))


def record_value(value: object, code: str = 'General') -> XmlElement:
    if value is EMPTY:
        return node('m')
    if isinstance(value, bool):
        return node('b', {'v': '1' if value else '0'})
    if isinstance(value, (VBAErrorValue, ExcelError)):
        from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS
        text = value.name if isinstance(value, ExcelError) else next(
            (name for name, number in ERROR_NUMBERS.items() if number == value.number), '#VALUE!')
        return node('s' if text in ('#SPILL!', '#CALC!') else 'e', {'v': text})
    if isinstance(value, VBADate):
        value = value.serial
    if isinstance(value, (int, float, Decimal)):
        value = float(value)
        if not math.isfinite(value):
            raise error(1004)
        from pyopenvba.apps.excel._typing import date_time_codes, is_date_format
        if date_time_codes(code):
            if value < 0:
                if is_date_format(code):
                    raise error(5)
                return node('n', {'v': repr(value)})
            try:
                VBADate(value)  # The cache's OLE-date conversion has the VBA Date bounds.
                days = math.trunc(value)
                fraction = abs(value - days)
                clock = dt.datetime(1899, 12, 30) + dt.timedelta(days=days, milliseconds=round(fraction * 86400000))
            except (ValueError, OverflowError, VBARuntimeError):
                raise error(440) from None
            # Excel spells a cached moment to the second (tests/fixtures/pivot_date_filters/).
            return node('d', {'v': clock.isoformat(timespec='milliseconds' if clock.microsecond else 'seconds')})
        return node('n', {'v': repr(float(value))})
    return node('s', {'v': encoded(to_text(value))})


def create_table(cache: PivotCache, destination: object, name: object, read: object, version: object) -> PivotTable:
    if read is not MISSING and not to_bool(read):
        raise error(1004)
    if version is not MISSING:
        chosen = int(to_integer(version, 'Long'))
        if chosen != int(cache.Version()):
            raise error(5)
    try:
        target = source_range(cache.book, destination)
    except VBARuntimeError:
        raise error(1004) from None
    if target.sheet.book is not cache.book or not target.areas:
        raise error(1004)
    if name is MISSING:
        number = 1
        used = {str(table.Name()).casefold() for sheet in cache.book.sheets_ for table in sheet.pivot_tables}
        while f'pivottable{number}' in used:
            number += 1
        caption = f'PivotTable{number}'
    else:
        caption = to_text(name)
    if not caption or len(caption) > 255 or any(ord(char) < 32 for char in caption) or any(
            str(table.Name()).casefold() == caption.casefold() for sheet in cache.book.sheets_ for table in sheet.pivot_tables):
        raise error(1004)
    first = target.first
    area = Area(first.top, first.left, first.top + 17, first.left + 2, target.sheet.name)
    if any(merged.contains(first.top, first.left) for merged in target.sheet.merged_areas):
        raise error(1004)
    for table in target.sheet.pivot_tables:
        location = table.root.find(NS + 'location')
        if location is not None:
            from pyopenvba._a1 import parse_area
            other = parse_area(location.get('ref', 'A1'))
            if area.top <= other.bottom and other.top <= area.bottom and area.left <= other.right and other.left <= area.right:
                raise error(1004)
    blocked = area.bottom > MAX_ROWS or area.right > MAX_COLUMNS or any(
        area.top <= merged.bottom and merged.top <= area.bottom and area.left <= merged.right and merged.left <= area.right
        for merged in target.sheet.merged_areas)
    if blocked:
        area = Area(first.top, first.left, first.top, first.left, target.sheet.name)
    # Further reports share the existing snapshot, even after the source changes.
    if cache.index == 0:
        materialize(cache)
    current = str(cache.Version())
    root = part_root('pivotTableDefinition', {'name': encoded(caption), 'cacheId': str(cache.cache_id),
                                       'dataCaption': 'Values', 'createdVersion': current, 'updatedVersion': current,
                                       'minRefreshableVersion': str(min(int(current), 3)),
                                       'useAutoFormatting': '1', 'rowGrandTotals': '1', 'colGrandTotals': '1',
                                       'outline': '1', 'outlineData': '1', 'compact': '1', 'compactData': '1',
                                       'multipleFieldFilters': '0'})
    root.append(node('location', {'ref': area.address(absolute=False), 'firstHeaderRow': str(int(not blocked)),
                                 'firstDataRow': str(int(not blocked)), 'firstDataCol': '0'}))
    fields = cache.root.find(NS + 'cacheFields')
    assert fields is not None
    pivots = node('pivotFields', {'count': str(len(fields))})
    for field in fields:
        # Each pivot field shows its items in its cache field's format (tests/fixtures/pivot_date_filters/).
        identifier = field.get('numFmtId', '0')
        pivots.append(node('pivotField', {'numFmtId': identifier, 'showAll': '0'} if identifier != '0'
                           else {'showAll': '0'}))
    root.append(pivots)
    root.append(node('pivotTableStyleInfo', {'name': 'PivotStyleLight16', 'showRowHeaders': '1',
                                         'showColHeaders': '1', 'showRowStripes': '0', 'showColStripes': '0', 'showLastColumn': '1'}))
    table = PivotTable(target.sheet, cache, root)
    table.changed = True
    cleared = Range(target.sheet, [area])
    cleared.clear_contents_checked()
    if blocked:
        cleared.put_as_value(first.top, first.left, VBAErrorValue(2045))
    if cache.index == 0:
        cache.book.pivot_caches.append(cache)
        cache.index = len(cache.book.pivot_caches)
    target.sheet.pivot_tables.insert(0, table)
    # Even blank report cells carry formatting and count toward UsedRange.
    from pyopenvba.apps.excel._styles import NO_SIDE, Border, Color, Side

    edge = Side('thin', Color('rgb', 'FF999999'))
    grid = Side('thin', Color('indexed', 65))
    for row in range(area.top, area.bottom + 1):
        for column in range(area.left, area.right + 1):
            old_style = target.sheet.style_at(row, column)
            border = Border(left=edge if column == area.left else grid,
                            right=edge if column == area.right else NO_SIDE,
                            top=edge if row == area.top else grid,
                            bottom=edge if row == area.bottom else NO_SIDE)
            style = replace(old_style, border=border, applied=old_style.applied | {'border'})
            target.sheet.restyle(row, column, style)
    target.sheet.touched()
    from pyopenvba.apps.excel import _events

    _events.after_edit(cleared)
    _events.pivot_updated(table)
    return table


def rich_errors(sheet: Worksheet) -> dict[tuple[int, int], RichError]:
    """Rich values for a blocked empty report's #SPILL! anchor."""
    from pyopenvba._a1 import parse_area
    from pyopenvba.apps.excel._dynamic_file import RichError

    result: dict[tuple[int, int], RichError] = {}
    for table in sheet.pivot_tables:
        location = table.root.find(NS + 'location')
        if location is None or table.root.find(NS + 'dataFields') is not None:
            continue
        area = parse_area(location.get('ref', 'A1'))
        cell = sheet.cell(area.top, area.left)
        if area.rows != 1 or area.columns != 1 or cell is None or cell.formula or cell.value != ExcelError('#SPILL!'):
            continue
        rows = min(17, MAX_ROWS - area.top)
        columns = min(2, MAX_COLUMNS - area.left)
        # Native pivot placeholders use subtype 7 for blocked spill areas.
        why = 3 if rows < 17 or columns < 2 else 7
        result[area.top, area.left] = RichError('#SPILL!', why, rows, columns)
    return result


def write(book: Workbook, package: OpcFile) -> None:
    """Write new caches and new or explicitly edited report definitions."""
    from pyopenvba.apps.excel._pivot_filters import savable

    for cache in book.pivot_caches:
        if not cache.changed:
            continue
        if cache.part:
            if not cache.records_part:
                from pyopenvba.apps.excel._pivot_refresh import existing_record_part

                cache.records_part = existing_record_part(cache)
            package.write(cache.part, xml_bytes(cache.root))
            # Metadata-only setters preserve lazy, untouched records byte for byte.
            if cache.records is not None:
                package.write(cache.records_part, xml_bytes(cache.records))
            cache.changed = False
            continue
        number = 1
        while any(package.has(part) for part in (
                f'xl/pivotCache/pivotCacheDefinition{number}.xml',
                f'xl/pivotCache/pivotCacheRecords{number}.xml',
                f'xl/pivotCache/richPivotRecords{number}.xml',
                f'xl/pivotCache/_rels/pivotCacheDefinition{number}.xml.rels',
                f'xl/pivotCache/_rels/pivotCacheRecords{number}.xml.rels')):
            number += 1
        cache.part = f'xl/pivotCache/pivotCacheDefinition{number}.xml'
        records_kind = 'richPivotRecords' if cache.rich_records else 'pivotCacheRecords'
        cache.records_part = f'xl/pivotCache/{records_kind}{number}.xml'
        cache.cache_id = max([0, *[other.cache_id for other in book.pivot_caches]]) + 1
        rid = add_relationship(package, 'xl/_rels/workbook.xml.rels', LINK + 'pivotCacheDefinition',
                               f'pivotCache/pivotCacheDefinition{number}.xml')
        text = package.read('xl/workbook.xml').decode('utf-8')
        entry = f'<pivotCache cacheId="{cache.cache_id}" r:id="{rid}"/>'
        if '</pivotCaches>' in text:
            text = text.replace('</pivotCaches>', entry + '</pivotCaches>', 1)
        else:
            match = re.search(r'<(?:smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|</workbook>', text)
            assert match is not None
            text = text[:match.start()] + '<pivotCaches>' + entry + '</pivotCaches>' + text[match.start():]
        package.write('xl/workbook.xml', text.encode('utf-8'))
        records_id = add_relationship(package, f'xl/pivotCache/_rels/pivotCacheDefinition{number}.xml.rels',
                                      LINK + records_kind, f'{records_kind}{number}.xml')
        if cache.rich_records:
            cache.root.set('invalid', '1')
            guid = cache.rich_guid
            cache.root.set('{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid', guid)
            extensions = node('extLst')
            version = node('ext', {'uri': '{9F748A41-CAEA-4470-BF7A-CE61E8FFA7F9}'})
            version_ns = '{http://schemas.microsoft.com/office/spreadsheetml/2022/pivotVersionInfo}'
            info = XmlElement(version_ns + 'cacheVersionInfo')
            feature = XmlElement(version_ns + 'lastRefreshFeature')
            feature.text = 'RichData'
            info.append(feature)
            version.append(info)
            extensions.append(version)
            rich = node('ext', {'uri': '{2C874A73-7782-4A18-856F-96AC7E287872}'})
            rich.append(XmlElement('{http://schemas.microsoft.com/office/spreadsheetml/2022/pivotRichData}richInfo',
                                   {'pivotCacheGuid': guid, 'pivotIgnoreInvalidCache': '1', REL + 'id': records_id}))
            extensions.append(rich)
            cache.root.append(extensions)
        else:
            cache.root.set(REL + 'id', records_id)
        package.write(cache.part, xml_bytes(cache.root))
        assert cache.records is not None
        package.write(cache.records_part, xml_bytes(cache.records))
        add_content_type(package, cache.part,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml')
        add_content_type(package, cache.records_part,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.' + records_kind + '+xml')
        cache.changed = False
    for sheet in book.sheets_:
        added = False
        for table in sheet.pivot_tables:
            if not table.changed:
                continue
            if table.part:
                package.write(table.part, xml_bytes(savable(table.root)))
                table.changed = False
                continue
            number = 1
            while package.has(f'xl/pivotTables/pivotTable{number}.xml') or package.has(
                    f'xl/pivotTables/_rels/pivotTable{number}.xml.rels'):
                number += 1
            table.part = f'xl/pivotTables/pivotTable{number}.xml'
            table.root.set('cacheId', str(table.cache.cache_id))
            folder, filename = posixpath.split(sheet.part_name)
            add_relationship(package, f'{folder}/_rels/{filename}.rels', LINK + 'pivotTable',
                             f'../pivotTables/pivotTable{number}.xml')
            add_relationship(package, f'xl/pivotTables/_rels/pivotTable{number}.xml.rels', LINK + 'pivotCacheDefinition',
                             '../pivotCache/' + posixpath.basename(table.cache.part))
            package.write(table.part, xml_bytes(savable(table.root)))
            add_content_type(package, table.part, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml')
            table.changed = False
            added = True
        if added:
            folder, filename = posixpath.split(sheet.part_name)
            part = f'{folder}/_rels/{filename}.rels'
            text = package.read(part).decode('utf-8')
            known = {table.part for table in sheet.pivot_tables}
            reports = [entry for entry in xml_tree(package.read(part)) if entry.get('Type') == LINK + 'pivotTable'
                       and entry.get('TargetMode') != 'External'
                       and posixpath.normpath(posixpath.join(folder, entry.get('Target', ''))).lstrip('/') in known]
            if all(re.fullmatch(r'rId[0-9]+', entry.get('Id', '')) for entry in reports):
                reports.sort(key=lambda entry: int(entry.get('Id', '')[3:]))
            # Preserve the model's deterministic report order across its own
            # saves; native Excel enumeration can vary across repeated opens.
            # Reassign only known targets, preserving unrelated worksheet ids.
            targets = {entry.get('Id', ''): posixpath.relpath(table.part, folder)
                       for entry, table in zip(reports, sheet.pivot_tables)}

            def reorder(match: re.Match[str]) -> str:
                tag = match.group()
                identifier = re.search(r'\bId="([^"]*)"', tag)
                target = targets.get(identifier[1]) if identifier is not None else None
                if target is not None:
                    return re.sub(r'\bTarget="[^"]*"', lambda _: f'Target="{escape(target, quote=True)}"', tag)
                return tag

            package.write(part, re.sub(r'<Relationship\b[^>]*/>', reorder, text).encode('utf-8'))

    from pyopenvba.apps.excel._pivot_refresh import persist_cache_order

    persist_cache_order(book, package)
