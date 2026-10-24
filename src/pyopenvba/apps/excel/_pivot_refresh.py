"""Fixed-schema local cache refresh, planned across all linked reports."""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement, XmlParseError, xml_tree
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._pivot_layout import RenderPlan
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable
    from pyopenvba.powerquery._opc import OpcFile


X14 = '{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}'
X16 = '{http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
XR = '{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


MAX_REPORTS = 64
MAX_WORK = 1_000_000
MAX_STAGED_CELLS = 1_000_000
MAX_TEXT = 8_000_000
MAX_METADATA_NODES = 200_000


class TextBudget:
    def __init__(self, copies: int) -> None:
        self.remaining = MAX_TEXT
        self.copies = copies
        self.nodes = 0

    def charge(self, value: object, copies: int = 1) -> None:
        if isinstance(value, str):
            self.remaining -= len(value) * copies
            if self.remaining < 0:
                raise unsupported('beyond the supported transaction text budget')

    def source_value(self, value: object) -> None:
        self.charge(value, self.copies)

    def metadata(self, root: XmlElement) -> None:
        for node in root.iter():
            self.nodes += 1
            if self.nodes > MAX_METADATA_NODES:
                raise unsupported('beyond the supported transaction metadata budget')
            self.charge(node.tag)
            self.charge(node.text)
            self.charge(node.tail)
            for name, value in node.attrib.items():
                self.charge(name)
                self.charge(value)


def admit_work(reports: list[PivotTable], record_count: int, field_count: int) -> None:
    work = 0
    for table in reports:
        root = table.root
        rows = len(root.findall(f'{NS}rowFields/*'))
        columns = len(root.findall(f'{NS}colFields/*'))
        pages = len(root.findall(f'{NS}pageFields/*'))
        measures = len(root.findall(f'{NS}dataFields/*'))
        axis_work = (rows + 1) * (columns + 1) * max(1, measures) * max(1, rows + columns + pages)
        work += record_count * max(field_count, axis_work)
        if work > MAX_WORK:
            raise unsupported('beyond the supported transaction planning budget')


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Pivot refresh ' + detail + ' is not implemented')


def attributes(root: XmlElement, allowed: str, *, provenance: bool = False) -> None:
    names = set(allowed.split())
    if provenance:
        names.update((MC + 'Ignorable', XR + 'uid', REL + 'id'))
    if any(name not in names and name != 'xmlns' and not name.startswith('xmlns:') for name in root.attrib):
        raise unsupported('with unknown metadata attributes')


def children(root: XmlElement, allowed: str) -> None:
    tags = {NS + name for name in allowed.split()}
    if any(child.tag not in tags for child in root):
        raise unsupported('with unknown metadata children')


def extensions(root: XmlElement, *, cache: bool) -> None:
    known = ({'{725AE2AE-9491-48be-B2B4-4EB974FC3084}': (X14 + 'pivotCacheDefinition', '')} if cache else {
        '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}': (X14 + 'pivotTableDefinition', 'hideValuesRow'),
        '{747A6164-185A-40DC-8AA5-F01512510D54}': (X16 + 'pivotTableDefinition16', '')})
    for group in root.findall(NS + 'extLst'):
        attributes(group, '')
        for extension in group:
            attributes(extension, 'uri')
            found = known.get(extension.get('uri', ''))
            if extension.tag != NS + 'ext' or found is None or len(extension) != 1:
                raise unsupported('with unknown extensions')
            node = extension[0]
            if node.tag != found[0] or len(node):
                raise unsupported('with unknown extension content')
            attributes(node, found[1])
            if node.get('hideValuesRow', '1') != '1':
                raise unsupported('with a custom Values header')


def integer(root: XmlElement, name: str, default: int = 0) -> int:
    try:
        return int(root.get(name, str(default)))
    except (ValueError, OverflowError):
        raise unsupported('with malformed numeric metadata') from None


def distinct_children(root: XmlElement) -> None:
    tags = [child.tag for child in root]
    if len(set(tags)) != len(tags):
        raise unsupported('with repeated metadata containers')


def existing_record_part(cache: PivotCache) -> str:
    """Resolve an ordinary existing part without materializing or changing state."""
    from pyopenvba.apps.excel._pivots import related_parts

    if not cache.part or int(cache.SourceType()) != 1 or cache.rich_records:
        raise unsupported('writing this cache without its known ordinary record part')
    found = related_parts(cache.book.package, cache.part, 'pivotCacheRecords')
    if len(found) != 1 or (cache.records_part and cache.records_part != found[0][1]):
        raise unsupported('with missing or ambiguous cache record relationships')
    if not cache.root.get(REL + 'id') or cache.root.get(REL + 'id') != found[0][0]:
        raise unsupported('with an inconsistent cache record relationship identity')
    part = found[0][1]
    for other in cache.book.pivot_caches:
        if other is cache or not other.part:
            continue
        if other.part == cache.part or any(target == part for _, target in related_parts(
                cache.book.package, other.part, 'pivotCacheRecords')):
            raise unsupported('with shared physical cache definitions or records')
    return part


def admit_records(cache: PivotCache, part: str, budget: TextBudget) -> XmlElement:
    try:
        if cache.records is None:
            raw = cache.book.package.read(part)
            if len(raw) > MAX_TEXT * 4:
                raise unsupported('beyond the supported cached-record byte budget')
            records = xml_tree(raw)
        else:
            records = cache.records
        budget.metadata(records)
    except (ValueError, XmlParseError):
        raise unsupported('reading malformed ordinary cache records') from None
    if records.tag != NS + 'pivotCacheRecords':
        raise unsupported('with an unknown cache record root')
    attributes(records, 'count', provenance=True)
    children(records, 'r')
    if len(records) > 1_000_000:
        raise unsupported('beyond the supported record budget')
    fields = cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    for record in records:
        attributes(record, '')
        children(record, 'n s d b e m x')
        if len(record) != len(fields):
            raise unsupported('with an inconsistent cached record schema')
        for index, item in enumerate(record):
            attributes(item, 'v')
            if len(item):
                raise unsupported('with unknown record extensions')
            if item.tag == NS + 'x':
                shared = fields[index].find(NS + 'sharedItems')
                value = integer(item, 'v', -1)
                if shared is None or not 0 <= value < len(shared):
                    raise unsupported('with an invalid cached shared-item reference')
    return records


def admit_cache(cache: PivotCache) -> None:
    if cache.root.tag != NS + 'pivotCacheDefinition':
        raise unsupported('with an unknown cache definition')
    if int(cache.SourceType()) != 1 or cache.rich_records or bool(cache.OLAP()):
        raise unsupported('for external, OLAP or rich sources')
    attributes(cache.root, 'createdVersion refreshedVersion minRefreshableVersion recordCount enableRefresh '
               'refreshedBy refreshedDate refreshedDateIso saveData refreshOnLoad backgroundQuery', provenance=True)
    children(cache.root, 'cacheSource cacheFields extLst')
    distinct_children(cache.root)
    extensions(cache.root, cache=True)
    source = cache.root.find(NS + 'cacheSource')
    origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
    if source is None or origin is None:
        raise unsupported('without a local rectangular source')
    attributes(source, 'type')
    children(source, 'worksheetSource')
    attributes(origin, 'ref sheet')
    if len(origin) or not origin.get('ref') or not origin.get('sheet'):
        raise unsupported('for named or nonrectangular sources')
    for field in cache.root.findall(f'{NS}cacheFields/{NS}cacheField'):
        attributes(field, 'name numFmtId')
        children(field, 'sharedItems')
        distinct_children(field)
        for shared in field:
            attributes(shared, 'count containsString containsNumber containsDate containsBlank containsNonDate '
                       'containsMixedTypes containsSemiMixedTypes containsInteger minValue maxValue minDate maxDate '
                       'longText')
            children(shared, 'n s b d e m')
            for item in shared:
                attributes(item, 'v u')
                if len(item):
                    raise unsupported('with custom shared-item metadata')
    fields = cache.root.find(NS + 'cacheFields')
    if fields is None or not len(fields):
        raise unsupported('without a populated schema')
    attributes(fields, 'count')
    children(fields, 'cacheField')


def admit_report(table: PivotTable, field_count: int, record_count: int) -> None:
    from pyopenvba.apps.excel._pivot_criteria import readable
    from pyopenvba.apps.excel._pivot_sorts import sort_scope

    root = table.root
    if root.tag != NS + 'pivotTableDefinition':
        raise unsupported('with an unknown report definition')
    attributes(root, 'name cacheId dataCaption createdVersion updatedVersion minRefreshableVersion useAutoFormatting '
               'rowGrandTotals colGrandTotals outline outlineData compact compactData dataOnRows pageWrap pageOverThenDown '
               'applyNumberFormats applyBorderFormats applyFontFormats applyPatternFormats applyAlignmentFormats '
               'applyWidthHeightFormats itemPrintTitles indent multipleFieldFilters', provenance=True)
    children(root, 'location pivotFields rowFields rowItems colFields colItems pageFields dataFields pivotTableStyleInfo filters '
             'extLst')
    if not readable(root):
        raise unsupported('with unknown filters')
    distinct_children(root)
    extensions(root, cache=False)
    for name in ('rowGrandTotals', 'colGrandTotals', 'dataOnRows'):
        if root.get(name, '0') not in ('0', '1', 'false', 'true'):
            raise unsupported('with malformed layout flags')
    if integer(root, 'pageWrap') < 0:
        raise unsupported('with invalid page wrapping')
    if root.get('compact', '1') != '1' or root.get('compactData', '1') != '1':
        raise unsupported('for noncompact reports')
    fields = root.find(NS + 'pivotFields')
    if fields is None or len(fields) != field_count:
        raise unsupported('with a mismatched report schema')
    attributes(fields, 'count')
    children(fields, 'pivotField')
    for field in fields:
        distinct_children(field)
        attributes(field, 'axis dataField showAll defaultSubtotal compact outline subtotalTop '
                   'multipleItemSelectionAllowed sortType measureFilter')
        children(field, 'items autoSortScope')
        if field.get('compact', '1') != '1':
            raise unsupported('with noncompact field layouts')
        if field.get('defaultSubtotal', '1') != '1':
            raise unsupported('with custom subtotals')
        if field.get('sortType', 'manual') not in ('manual', 'ascending', 'descending') or (
                field.find(NS + 'autoSortScope') is not None and sort_scope(field) is None):
            raise unsupported('with unknown sorts')
        for items in field.findall(NS + 'items'):
            attributes(items, 'count')
            children(items, 'item')
            for item in items:
                attributes(item, 'x t h m')
                if len(item) or item.get('t', 'data') not in ('data', 'default'):
                    raise unsupported('with grouped or filtered items')
    for name, attr in (('rowFields', 'x'), ('colFields', 'x'), ('pageFields', 'fld')):
        for group in root.findall(NS + name):
            attributes(group, 'count')
            children(group, 'pageField' if name == 'pageFields' else 'field')
            for entry in group:
                attributes(entry, 'fld item hier' if name == 'pageFields' else 'x')
                if name == 'pageFields' and entry.get('hier', '-1') != '-1':
                    raise unsupported('with hierarchical page fields')
                index = integer(entry, attr, -3)
                if not 0 <= index < field_count and not (index == -2 and name != 'pageFields'):
                    raise unsupported('with invalid field indices')
                if len(entry):
                    raise unsupported('with custom axis metadata')
    for data in root.findall(f'{NS}dataFields/{NS}dataField'):
        attributes(data, 'name fld subtotal baseField baseItem numFmtId')
        if not 0 <= integer(data, 'fld', -1) < field_count or data.get('subtotal', 'sum') not in (
                'sum', 'average', 'count', 'countNums', 'min', 'max', 'product', 'stdDev', 'stdDevp', 'var', 'varp'):
            raise unsupported('with custom data calculations')
        if len(data):
            raise unsupported('with custom data metadata')
    for name, allowed, sub in (
            ('location', 'ref firstHeaderRow firstDataRow firstDataCol rowPageCount colPageCount', ''),
            ('rowItems', 'count', 'i'), ('colItems', 'count', 'i'),
            ('dataFields', 'count', 'dataField'),
            ('pivotTableStyleInfo', 'name showRowHeaders showColHeaders showRowStripes showColStripes showLastColumn', '')):
        for group in root.findall(NS + name):
            attributes(group, allowed)
            children(group, sub)
            if name == 'location':
                try:
                    parse_area(group.get('ref', ''))
                except ValueError:
                    raise unsupported('with an invalid report location') from None
                for key in ('firstHeaderRow', 'firstDataRow', 'firstDataCol', 'rowPageCount', 'colPageCount'):
                    if integer(group, key) < 0:
                        raise unsupported('with invalid location metadata')
            elif name in ('rowItems', 'colItems'):
                for item in group:
                    attributes(item, 't i r')
                    children(item, 'x')
                    for value in item:
                        attributes(value, 'v')
                        if len(value):
                            raise unsupported('with custom line metadata')
    axes = sum(len(root.findall(f'{NS}{name}/*')) for name in ('rowFields', 'colFields', 'pageFields'))
    data_count = len(root.findall(f'{NS}dataFields/{NS}dataField'))
    if axes > 8 or record_count * max(1, axes) * max(1, data_count) > 1_000_000:
        raise unsupported('beyond the supported report planning budget')


def admit_cells(sheet: Worksheet, areas: list[Area]) -> None:
    from pyopenvba.apps.excel._conditional import meets
    from pyopenvba.apps.excel._conditional_transfer import guard_extensions
    from pyopenvba.apps.excel._hyperlinks import touches
    from pyopenvba.apps.excel._protection import whole

    if any(area.rows * area.columns > 1_000_000 for area in areas):
        raise unsupported('beyond the supported output budget')
    whole(sheet, areas, 'Refreshing a PivotTable')
    guard_extensions(sheet)
    if sheet.auto_filter is not None or getattr(sheet, 'advanced_filter', None) is not None:
        raise unsupported('on filtered sheets')
    blocks = [*sheet.merged_areas, *sheet.array_formulas.values(), *(spill.area for spill in sheet.spills.values()),
              *(table.area for table in sheet.tables), *(table.area for table in sheet.data_tables.values())]
    blocks.extend(part for rule in [*sheet.conditional_formats, *sheet.validations] for part in rule.areas)
    if (any(meets(area, block) for area in areas for block in blocks)
            or any(any(area.contains(*key) for area in areas) for key in sheet.notes)
            or any(any(area.contains(*key) for area in areas) and cell.dynamic for key, cell in sheet.cells_.items())
            or touches(sheet, areas)):
        raise unsupported('over cells with arrays, rules, tables, annotations or merges')
    if any(area.rows * area.columns > 1_000_000 for area in areas):
        raise unsupported('beyond the supported output budget')


def report_areas(table: PivotTable) -> list[Area]:
    """Read an existing report footprint without changing its metadata."""
    from pyopenvba.apps.excel._pivot_ranges import page_range

    location = table.root.find(NS + 'location')
    if location is None:
        raise unsupported('with a missing report location')
    try:
        areas = [parse_area(location.get('ref', ''))]
        pages = table.root.find(NS + 'pageFields')
        if pages is not None and len(pages):
            areas.extend(page_range(table).areas)
    except (ValueError, VBARuntimeError):
        raise unsupported('with an invalid existing report footprint') from None
    return areas


def admit_footprints(plans: list[RenderPlan]) -> None:
    """Check the complete transaction, including report growth, before publishing."""
    from pyopenvba.apps.excel._conditional import meets

    footprints = {table: report_areas(table) for sheet in {plan.table.sheet for plan in plans} for table in sheet.pivot_tables}
    for index, plan in enumerate(plans):
        sheet = plan.table.sheet
        for other in plans[index + 1:]:
            if sheet is other.table.sheet and meets(plan.extent, other.extent):
                raise unsupported('with overlapping report footprints')
        for table in sheet.pivot_tables:
            if table is not plan.table and any(meets(plan.extent, area) for area in footprints[table]):
                raise unsupported('with overlapping report footprints')
        own = footprints[plan.table]
        if any(plan.extent.contains(*key) and not any(area.contains(*key) for area in own)
               and (cell.value is not EMPTY or cell.formula)
               for key, cell in sheet.cells_.items()):
            raise unsupported('over populated cells outside the previous report footprint')


def refresh(cache: PivotCache) -> object:
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._pivot_aggregate import records
    from pyopenvba.apps.excel._pivot_create import apply_materialized, plan_materialize
    from pyopenvba.apps.excel._pivot_filters import retain_missing
    from pyopenvba.apps.excel._pivot_layout import apply_render, cache_layout, plan_render

    if cache.index == 0:
        raise error(1004)
    if not cache.EnableRefresh():
        raise error(1004, 'Refresh has been disabled by a Visual Basic macro.')
    book = cache.book
    if book.application.udf_callers:
        raise error(1004, 'a worksheet function cannot change cells')
    if cache not in book.pivot_caches:
        raise unsupported('with an unregistered cache identity')
    if book.package is not None and cache.part:
        original = book.package.read('xl/workbook.xml')
        cache_order_bytes(original, [entry.cache_id for entry in book.pivot_caches if entry.part])
    reports = [table for sheet in book.sheets_ for table in sheet.pivot_tables if table.cache is cache]
    if not reports or len({table.sheet for table in reports}) != 1:
        raise unsupported('for reportless or cross-sheet cache groups')
    if len(reports[0].sheet.pivot_tables) > MAX_REPORTS:
        raise unsupported('beyond the supported transaction report budget')
    budget = TextBudget(len(reports))
    budget.metadata(cache.root)
    for table in reports[0].sheet.pivot_tables:
        budget.metadata(table.root)
    admit_cache(cache)
    record_part = existing_record_part(cache) if cache.part else cache.records_part
    previous_records = admit_records(cache, record_part, budget) if cache.part or cache.records is not None else None
    origin = cache.root.find(f'{NS}cacheSource/{NS}worksheetSource')
    assert origin is not None
    owner = next((sheet for sheet in book.sheets_ if sheet.name.casefold() == origin.get('sheet', '').casefold()), None)
    try:
        area = parse_area(origin.get('ref', ''))
    except ValueError:
        raise unsupported('for a nonrectangular local source reference') from None
    if owner is None or area.sheet:
        raise unsupported('without its declared local source worksheet')
    source = Range(owner, [area])
    if len(source.areas) != 1 or source.first.rows * source.first.columns > 1_000_000:
        raise unsupported('beyond the supported source budget')
    admit_work(reports, source.first.rows - 1, source.first.columns)
    for row in range(source.first.top, source.first.bottom + 1):
        for column in range(source.first.left, source.first.right + 1):
            cell = source.sheet.cell(row, column)
            if cell is not None:
                budget.source_value(cell.value)
                budget.charge(cell.formula)
                budget.charge(cell.number_format)
    before = cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    for table in reports:
        admit_report(table, len(before), source.first.rows - 1)
    staged = plan_materialize(cache, source=source, charge_text=budget.source_value)
    staged.records_part = record_part
    if previous_records is not None:
        assert staged.records is not None
        staged.records.attrib.update({name: value for name, value in previous_records.attrib.items() if name != 'count'})
    if staged.rich_records:
        raise unsupported('for rich source values')
    after = staged.root.findall(f'{NS}cacheFields/{NS}cacheField')
    if len(before) != len(after) or any(old.get('name') != new.get('name') for old, new in zip(before, after, strict=True)):
        raise unsupported('with a changed source schema')
    for old, new in zip(before, after, strict=True):
        new.attrib.update(old.attrib)
    normalized = cache_layout(staged, records(staged))
    retain_missing(cache, normalized[0], reports)
    plans: list[RenderPlan] = []
    remaining_cells = MAX_STAGED_CELLS
    for table in reports:
        definition = deepcopy(table.root)
        for name in ('rowGrandTotals', 'colGrandTotals'):
            value = definition.get(name)
            if value in ('true', 'false'):
                definition.set(name, '1' if value == 'true' else '0')
        plan = plan_render(table, definition, cache=staged, normalized=normalized, cell_limit=remaining_cells)
        remaining_cells -= plan.extent.rows * plan.extent.columns
        admit_cells(table.sheet, plan.previous_areas)
        if plan.table.sheet is source.sheet and plan.extent.top <= source.first.bottom \
                and source.first.top <= plan.extent.bottom and plan.extent.left <= source.first.right \
                and source.first.left <= plan.extent.right:
            raise unsupported('with an overlapping source and report')
        plans.append(plan)
    admit_footprints(plans)
    # The complete source snapshot and every linked report are admitted first.
    apply_materialized(cache, staged)
    for plan in plans:
        apply_render(plan)
    book.pivot_caches.remove(cache)
    book.pivot_caches.append(cache)
    for index, entry in enumerate(book.pivot_caches, 1):
        entry.index = index
    book.saved = False
    for plan in plans:
        _events.changed(Range(plan.table.sheet, [plan.extent]))
        if book.calculator.automatic:
            book.calculator.calculate_all(only=plan.table.sheet, background=True)
        _events.pivot_refresh_calculated(plan.table.sheet)
        _events.pivot_updated(plan.table)
    if book.calculator.automatic:
        _events.pivot_refresh_calculated(reports[0].sheet)
    return EMPTY


def cache_order_bytes(data: bytes, identifiers: list[int]) -> bytes:
    """Reorder direct workbook cache spans without changing unrelated bytes."""
    from pyopenvba._opc_xml import opc_part

    group = opc_part(data, NS[1:-1], 'workbook', container_path=('pivotCaches',))
    if group is None:
        if identifiers:
            raise unsupported('without its package cache collection')
        return data
    entries: list[tuple[int, int, int]] = []
    for name, attrs, start, finish in group.children:
        if name == NS[1:] + 'pivotCache':
            try:
                identifier = int(attrs.get('cacheId', '-1'))
            except ValueError:
                raise unsupported('with a malformed package cache identifier') from None
            entries.append((identifier, start, finish))
    by_id = {identifier: data[start:finish] for identifier, start, finish in entries}
    if len(by_id) != len(entries) or len(set(identifiers)) != len(identifiers) or set(by_id) != set(identifiers):
        raise unsupported('with ambiguous package cache identities')
    output: list[bytes] = []
    previous = 0
    for (_, start, finish), identifier in zip(entries, identifiers, strict=True):
        output.extend((data[previous:start], by_id[identifier]))
        previous = finish
    output.append(data[previous:])
    return b''.join(output)


def persist_cache_order(book: Workbook, package: OpcFile) -> None:
    if not book.pivot_caches:
        return
    original = package.read('xl/workbook.xml')
    reordered = cache_order_bytes(original, [cache.cache_id for cache in book.pivot_caches])
    if reordered != original:
        package.write('xl/workbook.xml', reordered)
