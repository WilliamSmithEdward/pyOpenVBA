"""Runtime PivotTable views and worksheet-source cache/report creation.

Relationship traversal follows pyOfficeEditor's pivot reader. Untouched XML stays
in the package unchanged. Worksheet sources can build caches, empty reports and
compact aggregate reports and multi-axis field ranges. Refresh, grouping and arbitrary layout edits remain incomplete.
"""
from __future__ import annotations

import posixpath
import re
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._xstring import decode as decode_text
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBADate, VBAErrorValue, VBAInt, error, string_order, to_integer, to_text, to_bool, utf16_text

if TYPE_CHECKING:
    from pyopenvba.interpreter._objects import MemberSpec
    from pyopenvba.apps.excel._pivot_reports import PivotReport
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.powerquery._opc import OpcFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _pivot_argument(value: object) -> str | float | bool:
    if isinstance(value, Range):
        value = value.vba_value()
    if isinstance(value, str):
        return utf16_text(value)
    if isinstance(value, bool):
        return value
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    raise error(1004)


def related_parts(package: OpcFile, source: str, kind: str) -> list[tuple[str, str]]:
    folder, name = posixpath.split(source)
    path = f'{folder}/_rels/{name}.rels'
    if not package.has(path):
        return []
    found: list[tuple[str, str]] = []
    for entry in xml_tree(package.read(path)):
        if entry.get('TargetMode') == 'External' or entry.get('Type') != REL[1:-1] + '/' + kind:
            continue
        target = entry.get('Target', '')
        part = posixpath.normpath(posixpath.join(folder, target)) if not target.startswith('/') else target[1:]
        if package.has(part):
            found.append((entry.get('Id', ''), part))
    return found


def _numbered(root: XmlElement, book: Workbook) -> XmlElement:
    """A cache's or report's tree with each number format it names by id under the id Excel gave that format on
    opening the file (tests/fixtures/number_format_resave/pivot_id.xlsx)."""
    for element in root.iter():
        identifier = element.get('numFmtId')
        if identifier is not None and identifier.isdecimal():
            element.set('numFmtId', str(book.stylesheet.opening_id(identifier)))
    return root


def named_formats(book: Workbook) -> set[int]:
    """The number formats the workbook's pivot caches and reports name by id, which a save keeps."""
    roots = [*(cache.root for cache in book.pivot_caches),
             *(table.root for sheet in book.sheets_ for table in sheet.pivot_tables)]
    return {int(identifier) for root in roots for element in root.iter()
            if (identifier := element.get('numFmtId')) is not None and identifier.isdecimal()}


def number_parts(book: Workbook, package: OpcFile) -> None:
    """Give the parts of the file's caches and reports the ids Excel numbers their formats with as it opens the
    file, for a save to write; a part the save writes afresh holds them already."""
    parts = {*(cache.part for cache in book.pivot_caches),
             *(table.part for sheet in book.sheets_ for table in sheet.pivot_tables)}
    for part in sorted(part for part in parts if part and package.has(part)):
        text = package.read(part).decode('utf-8')
        numbered = book.stylesheet.numbered_part(text)
        if numbered != text:
            package.write(part, numbered.encode('utf-8'))


def read(book: Workbook) -> None:
    """Read caches in workbook order, then each sheet's report relationships."""
    package = book.package
    by_part: dict[str, PivotCache] = {}
    by_id = dict(related_parts(package, 'xl/workbook.xml', 'pivotCacheDefinition'))
    root = xml_tree(package.read('xl/workbook.xml'))
    for entry in root.findall(f'{NS}pivotCaches/{NS}pivotCache'):
        part = by_id.get(entry.get(REL + 'id', ''), '')
        if not part:
            continue
        cache = PivotCache(book, part, _numbered(xml_tree(package.read(part)), book), len(book.pivot_caches) + 1)
        cache.cache_id = int(entry.get('cacheId', '0'))
        cache.rich_records = bool(related_parts(package, part, 'richPivotRecords'))
        if cache.rich_records:
            from pyopenvba.apps.excel._pivot_rich import datasets

            info = cache.root.find(f'{NS}extLst/{NS}ext/'
                                   '{http://schemas.microsoft.com/office/spreadsheetml/2022/pivotRichData}richInfo')
            cache.rich_guid = info.get('pivotCacheGuid', '') if info is not None else ''
            cache.rich_columns = datasets(package).get(cache.rich_guid)
        book.pivot_caches.append(cache)
        by_part[part] = cache
    for sheet in book.sheets_:
        reports = related_parts(package, sheet.part_name, 'pivotTable')
        # Keep import order deterministic. Native enumeration can rotate across
        # repeated opens of an identical package, so its exact order is not a
        # reproducible contract. Foreign ids retain their source order.
        if all(re.fullmatch(r'rId[0-9]+', identifier) for identifier, _ in reports):
            reports.sort(key=lambda pair: int(pair[0][3:]))
        for _, part in reports:
            caches = related_parts(package, part, 'pivotCacheDefinition')
            cache = by_part.get(caches[0][1]) if caches else None
            if cache is not None:
                table = PivotTable(sheet, cache, _numbered(xml_tree(package.read(part)), book))
                table.part = part
                sheet.pivot_tables.append(table)


class _Collection(VBACollection, ExcelObject):
    def __init__(self, parent: ExcelObject, items: list[object]) -> None:
        self.parent = parent
        self.items = items

    def vba_items(self) -> list[object]:
        return self.items

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                if isinstance(item, ExcelObject) and string_order(str(item.vba_get('Name')).casefold(), index.casefold()) == 0:
                    return item
        else:
            position = int(to_integer(index, 'Long'))
            if 1 <= position <= len(items):
                return items[position - 1]
        raise error(1004)

    @member
    def Parent(self) -> object:
        return self.parent


class PivotCaches(_Collection):
    vba_type_name = 'PivotCaches'

    @method
    def Create(self, SourceType: object, SourceData: object = MISSING, Version: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_create import create_cache
        from pyopenvba.apps.excel._model import Workbook

        assert isinstance(self.parent, Workbook)
        return create_cache(self.parent, SourceType, SourceData, Version)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        # Cache indices are numeric; numeric strings receive normal Long coercion.
        return super().vba_lookup(to_integer(index, 'Long'), items)


class PivotTables(_Collection):
    vba_type_name = 'PivotTables'


class PivotFields(_Collection):
    vba_type_name = 'PivotFields'

    def __init__(self, parent: ExcelObject, items: list[object], *, data_names: bool = False) -> None:
        super().__init__(parent, items)
        self.data_names = data_names

    def vba_lookup(self, index: object, items: list[object]) -> object:
        try:
            return super().vba_lookup(index, items)
        except VBARuntimeError:
            # A table's PivotFields("Sum of Amount") finds the data field of that caption
            # (tests/fixtures/pivotfilters/ data_by_name).
            table = self.parent
            if not self.data_names or not isinstance(index, str) or not isinstance(table, PivotTable):
                raise
            wanted = index.casefold()
            for report in table.root.findall(f'{NS}dataFields/{NS}dataField'):
                if string_order(decode_text(report.get('name', '')).casefold(), wanted) == 0:
                    return PivotField(table, int(report.get('fld', '0')), report, True)
            raise


def refuse_1904_dates(cache: PivotCache) -> None:
    """Refuse a cache holding dates in a workbook on the 1904 calendar.

    A cache keeps a date as the calendar date it is, and the model reads
    and writes those as serials on the 1900 calendar only.
    """
    if cache.book.epoch_1904 and any(items.get('containsDate') in ('1', 'true')
                                     for items in cache.root.iter(NS + 'sharedItems')):
        raise VBAUnsupportedError('A PivotCache holding dates in a workbook on the 1904 calendar is not implemented')


class PivotCache(ExcelObject):
    vba_type_name = 'PivotCache'

    def __init__(self, book: Workbook, part: str, root: XmlElement, index: int) -> None:
        self.book, self.part, self.root, self.index = book, part, root, index
        self.pending_source: object = None
        self.source_text: str | None = None
        self.cache_id = 0
        self.records: XmlElement | None = None
        self.records_part = ''
        self.rich_records = False
        self.rich_guid = ''
        self.rich_columns: dict[str, list[str | None]] | None = None
        self.changed = False

    @method
    def CreatePivotTable(self, TableDestination: object, TableName: object = MISSING,
                         ReadData: object = MISSING, DefaultVersion: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_create import create_table

        return create_table(self, TableDestination, TableName, ReadData, DefaultVersion)

    @method
    def Refresh(self) -> object:
        from pyopenvba.apps.excel._pivot_refresh import refresh

        return refresh(self)

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Index(self) -> object:
        return VBAInt(self.index, 'Long')

    @member
    def RecordCount(self) -> object:
        return VBAInt(int(self.root.get('recordCount', '0')), 'Long')

    @member
    def SourceType(self) -> object:
        source = self.root.find(NS + 'cacheSource')
        return VBAInt({'worksheet': 1, 'external': 2, 'consolidation': 3, 'scenario': -4148}.get(
            source.get('type', '') if source is not None else '', 2), 'Long')

    @member
    def SourceData(self) -> object:
        if self.source_text is not None:
            return self.source_text
        source = self.root.find(f'{NS}cacheSource/{NS}worksheetSource')
        if source is None:
            raise error(1004)
        if source.get('name'):
            return source.get('name')
        area = parse_area(source.get('ref', 'A1'))
        sheet = source.get('sheet', '')
        if not sheet.replace('_', '').isalnum() or ' ' in sheet:
            sheet = "'" + sheet.replace("'", "''") + "'"
        return f'{sheet}!R{area.top}C{area.left}:R{area.bottom}C{area.right}'

    @member
    def OLAP(self) -> object:
        source = self.root.find(NS + 'cacheSource')
        return source is not None and source.get('type') == 'external' and self.root.find(NS + 'cacheHierarchies') is not None

    @member
    def EnableRefresh(self) -> object:
        return self.root.get('enableRefresh', '1') not in ('0', 'false')

    @setter('EnableRefresh')
    def set_enable_refresh(self, value: object) -> None:
        if self.book.application.udf_callers:
            raise VBAUnsupportedError('Changing PivotCache refresh flags from worksheet functions is not implemented')
        self.root.set('enableRefresh', str(int(to_bool(value))))
        self.changed = True
        self.book.saved = False

    @member
    def RefreshOnFileOpen(self) -> object:
        return self.root.get('refreshOnLoad', '0') in ('1', 'true')

    @member
    def Version(self) -> object:
        return VBAInt(int(self.root.get('createdVersion', '0')), 'Long')


class PivotTable(ExcelObject):
    vba_type_name = 'PivotTable'

    def __init__(self, sheet: Worksheet, cache: PivotCache, root: XmlElement,
                 report_owner: PivotTable | None = None) -> None:
        self.sheet, self.cache, self.root = sheet, cache, root
        self.entry = report_owner if report_owner is not None else self
        self.part = ''
        self.changed = False
        # Supported layout edits invalidate this parsed view. GETPIVOTDATA reads
        # each result cell afresh rather than caching its value here.
        self.report_metadata: PivotReport | None = None
        self.report_loaded = False
        #: The field a macro's AutoSort xlManual named, by the field it sorted, which no file keeps.
        self.sort_names: dict[int, str] = {}

    @method
    def RefreshTable(self) -> object:
        from pyopenvba.apps.excel._pivot_refresh import refresh

        refresh(self.cache)
        return True

    def _range(self, kind: str) -> object:
        location = self.root.find(NS + 'location')
        if location is None:
            raise error(1004)
        area = parse_area(location.get('ref', 'A1'))
        dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
        if kind == 'all' and self.root.findall(f'{NS}pageFields/{NS}pageField') and not any(
                self.root.findall(f'{NS}{name}/*') for name in ('rowFields', 'colFields', 'dataFields')):
            raise error(1004)
        if kind == 'data':
            area = Area(min(area.top + dr, area.bottom), min(area.left + dc, area.right), area.bottom, area.right)
        elif kind == 'row':
            if not self.root.findall(f'{NS}rowFields/{NS}field') or dc == 0:
                raise error(1004)
            area = Area(area.top + max(0, dr - 1), area.left,
                        area.bottom, area.left + dc - 1)
        elif kind == 'column':
            if not self.root.findall(f'{NS}colFields/{NS}field') or dr == 0:
                raise error(1004)
            area = Area(area.top, area.left + dc, area.top + dr - 1, area.right)
        return Range(self.sheet, [Area(area.top, area.left, area.bottom, area.right, self.sheet.name)])

    @member
    def Name(self) -> object:
        return decode_text(self.root.get('name', ''))

    @member
    def Version(self) -> object:
        return VBAInt(int(self.root.get('createdVersion', '0')), 'Long')

    @member
    def Parent(self) -> object:
        return self.sheet

    @method
    def PivotCache(self) -> object:
        return self.cache

    @method
    def AddDataField(self, Field: object, Caption: object = MISSING, Function: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_aggregate import add_data_field

        return add_data_field(self, Field, Caption, Function)

    @method
    def ClearAllFilters(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import clear_table

        clear_table(self)
        return EMPTY

    @member
    def AllowMultipleFilters(self) -> object:
        return self.entry.root.get('multipleFieldFilters', '1') not in ('0', 'false')

    @setter('AllowMultipleFilters')
    def set_allow_multiple_filters(self, value: object) -> None:
        table = self.entry
        if to_bool(value):
            table.root.attrib.pop('multipleFieldFilters', None)
        else:
            table.root.set('multipleFieldFilters', '0')
        table.changed = True
        table.sheet.book.saved = False

    def report(self) -> PivotReport | None:
        from pyopenvba.apps.excel._pivot_reports import read_report

        if self.entry is not self:
            return self.entry.report()
        if not self.report_loaded:
            formats = self.sheet.book.stylesheet.number_formats
            self.report_metadata = read_report(self, lambda identifier: formats.get(identifier, 'General'))
            self.report_loaded = True
        return self.report_metadata

    @method
    def GetPivotData(self, DataField: object = MISSING,
                     Field1: object = MISSING, Item1: object = MISSING,
                     Field2: object = MISSING, Item2: object = MISSING,
                     Field3: object = MISSING, Item3: object = MISSING,
                     Field4: object = MISSING, Item4: object = MISSING,
                     Field5: object = MISSING, Item5: object = MISSING,
                     Field6: object = MISSING, Item6: object = MISSING,
                     Field7: object = MISSING, Item7: object = MISSING,
                     Field8: object = MISSING, Item8: object = MISSING,
                     Field9: object = MISSING, Item9: object = MISSING,
                     Field10: object = MISSING, Item10: object = MISSING,
                     Field11: object = MISSING, Item11: object = MISSING,
                     Field12: object = MISSING, Item12: object = MISSING,
                     Field13: object = MISSING, Item13: object = MISSING,
                     Field14: object = MISSING, Item14: object = MISSING) -> object:
        report = self.report()
        if report is None:
            raise error(1004)
        if not report.readable:
            raise VBAUnsupportedError('GetPivotData on a cube or unsupported report layout')
        pairs = [Field1, Item1, Field2, Item2, Field3, Item3, Field4, Item4, Field5, Item5,
                 Field6, Item6, Field7, Item7, Field8, Item8, Field9, Item9, Field10, Item10,
                 Field11, Item11, Field12, Item12, Field13, Item13, Field14, Item14]
        while pairs and pairs[-1] is MISSING:
            pairs.pop()
        if len(pairs) % 2:
            raise error(1004)
        if DataField is MISSING or DataField is NULL:
            if any(argument is not NULL for argument in pairs) or len(report.values) != 1:
                raise error(1004)
            value = 0
        else:
            value = report.value_index(to_text(_pivot_argument(DataField)))
            if value is None:
                raise error(1004)
        chosen: dict[int, int] = {}
        for position in range(0, len(pairs), 2):
            field_arg, item_arg = pairs[position:position + 2]
            if field_arg is NULL and item_arg is NULL:
                if any(argument is not NULL for argument in pairs[position:]):
                    raise error(1004)
                break
            field = report.field_index(to_text(_pivot_argument(field_arg)))
            wanted = _pivot_argument(item_arg)
            if field is None:
                raise error(1004)
            item = report.item_index(field, wanted)
            if item is None or (field in chosen and chosen[field] != item):
                raise error(1004)
            chosen[field] = item
        cell = report.cell(value, chosen)
        if cell is None:
            raise error(1004)
        return Range(self.sheet, [Area(cell.row, cell.column, cell.row, cell.column, self.sheet.name)])

    @member
    def CacheIndex(self) -> object:
        return self.cache.Index()

    @member
    def SourceData(self) -> object:
        return self.cache.SourceData()

    @member
    def TableRange1(self) -> object:
        return self._range('all')

    @member
    def TableRange2(self) -> object:
        if self.root.findall(f'{NS}pageFields/{NS}pageField'):
            from pyopenvba.apps.excel._pivot_ranges import page_range

            location = self.root.find(NS + 'location')
            if location is None:
                raise error(1004)
            area = parse_area(location.get('ref', 'A1'))
            pages = page_range(self).first
            return Range(self.sheet, [Area(pages.top, area.left, area.bottom,
                                           max(area.right, pages.right), self.sheet.name)])
        return self._range('all')

    @member
    def PageRange(self) -> object:
        from pyopenvba.apps.excel._pivot_ranges import page_range

        return page_range(self)

    @member
    def DataBodyRange(self) -> object:
        return self._range('data')

    @member
    def RowRange(self) -> object:
        return self._range('row')

    @member
    def ColumnRange(self) -> object:
        return self._range('column')

    @member
    def RowGrand(self) -> object:
        return self.root.get('rowGrandTotals', '1') not in ('0', 'false')

    @member
    def ColumnGrand(self) -> object:
        return self.root.get('colGrandTotals', '1') not in ('0', 'false')

    def _fields(self, index: object, orientation: int | None = None) -> object:
        items: list[object] = []
        for i, field in enumerate(self.root.findall(f'{NS}pivotFields/{NS}pivotField')):
            view = PivotField(self, i, field)
            if orientation == 0 and field.get('dataField') == '1':
                continue
            if orientation is None or view.Orientation() == orientation:
                items.append(view)
        for axis, kind in (('rowFields', 1), ('colFields', 2)):
            if any(field.get('x') == '-2' for field in self.root.findall(f'{NS}{axis}/{NS}field')):
                if orientation is None or orientation == kind:
                    items.append(self._data_pivot_field(axis))
        fields = PivotFields(self, items, data_names=orientation is None)
        return fields if index is MISSING else fields.Item(index)

    def _data_pivot_field(self, axis: str | None) -> PivotField:
        field = XmlElement(NS + 'pivotField', {'name': self.root.get('dataCaption', 'Values')})
        if axis is not None:
            field.set('axis', 'axisRow' if axis == 'rowFields' else 'axisCol')
        return PivotField(self, -2, field)

    @member
    def DataPivotField(self) -> object:
        for axis in ('rowFields', 'colFields'):
            if any(field.get('x') == '-2' for field in self.root.findall(f'{NS}{axis}/{NS}field')):
                return self._data_pivot_field(axis)
        if self.root.find(NS + 'dataFields') is not None:
            return self._data_pivot_field(None)
        raise VBAUnsupportedError('DataPivotField without a visible Values axis is not implemented')

    @member
    def PivotFields(self, Index: object = MISSING) -> object:
        return self._fields(Index)

    @member
    def RowFields(self, Index: object = MISSING) -> object:
        return self._fields(Index, 1)

    @member
    def ColumnFields(self, Index: object = MISSING) -> object:
        return self._fields(Index, 2)

    @member
    def PageFields(self, Index: object = MISSING) -> object:
        return self._fields(Index, 3)

    @member
    def HiddenFields(self, Index: object = MISSING) -> object:
        return self._fields(Index, 0)

    @member
    def DataFields(self, Index: object = MISSING) -> object:
        fields = PivotFields(self, [PivotField(self, int(field.get('fld', '0')), field, True)
                                  for field in self.root.findall(f'{NS}dataFields/{NS}dataField')])
        return fields if Index is MISSING else fields.Item(Index)


class PivotField(ExcelObject):
    vba_type_name = 'PivotField'

    def __init__(self, table: PivotTable, index: int, root: XmlElement, data: bool = False) -> None:
        self.table, self.index, self.root, self.data = table, index, root, data

    def check_alive(self) -> None:
        if self.data and self.root not in self.table.root.findall(f'{NS}dataFields/{NS}dataField'):
            raise error(424)

    def vba_member(self, name: str) -> MemberSpec | None:
        self.check_alive()
        return super().vba_member(name)

    @member
    def SourceName(self) -> object:
        if self.index == -2:
            return VBAErrorValue(2042)
        fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
        return decode_text(fields[self.index].get('name', '')) if self.index < len(fields) else ''

    @member
    def Name(self) -> object:
        if self.index == -2:
            return decode_text(self.table.root.get('dataCaption', 'Values'))
        name = self.root.get('name')
        return decode_text(name) if name is not None else self.SourceName()

    @member
    def Caption(self) -> object:
        return self.Name()

    @member
    def Orientation(self) -> object:
        if self.index == -2:
            for name, orientation in (('rowFields', 1), ('colFields', 2)):
                if any(entry.get('x') == '-2' for entry in self.table.root.findall(f'{NS}{name}/{NS}field')):
                    return VBAInt(orientation, 'Long')
            return VBAInt(0, 'Long')
        return VBAInt(4 if self.data else
                      {'axisRow': 1, 'axisCol': 2, 'axisPage': 3}.get(self.root.get('axis', ''), 0), 'Long')

    @setter('Orientation')
    def set_orientation(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_layout import orient

        orient(self, value)

    @member
    def Parent(self) -> object:
        return self.table

    @member
    def Position(self) -> object:
        axis = {1: 'rowFields', 2: 'colFields', 3: 'pageFields', 4: 'dataFields'}.get(int(self.Orientation()))
        if axis is not None:
            for position, field in enumerate(self.table.root.findall(f'{NS}{axis}/*'), 1):
                if (self.data and field is self.root) or (not self.data and int(field.get('x', field.get('fld', '-1'))) == self.index):
                    count = len(self.table.root.findall(f'{NS}{axis}/*'))
                    return VBAInt(count + 1 - position if axis == 'pageFields' else position, 'Long')
        if self.index == -2:
            raise error(1004)
        return VBAErrorValue(2042)

    @member
    def Function(self) -> object:
        if not self.data:
            raise error(1004)
        # Excel accepts xlCountNums (-4113) when adding the field but returns 11
        # from its Function getter, including after saving and reopening.
        functions = {'sum': -4157, 'count': -4112, 'average': -4106, 'max': -4136, 'min': -4139,
                     'product': -4149, 'countNums': 11, 'stdDev': -4155, 'stdDevp': -4156,
                     'var': -4164, 'varp': -4165}
        return VBAInt(functions[self.root.get('subtotal', 'sum')], 'Long')

    @member
    def NumberFormat(self) -> object:
        identifier = int(self.root.get('numFmtId', '0'))
        return self.table.sheet.book.stylesheet.number_formats.get(identifier, 'General')

    def _range(self, label: bool) -> object:
        from pyopenvba.apps.excel._pivot_ranges import field_range

        return field_range(self, label)

    @member
    def DataRange(self) -> object:
        return self._range(False)

    @member
    def LabelRange(self) -> object:
        return self._range(True)

    @member
    def PivotItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_filters import first_for_names, list_on_reading

        list_on_reading(self)
        collection = PivotItems(self, list(self.listed_items()), first_for_names=first_for_names(self))
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def VisibleItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_filters import items_showing, list_on_reading

        list_on_reading(self)
        collection = PivotItems(self, items_showing(self, True))
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def HiddenItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_filters import items_showing, list_on_reading

        list_on_reading(self)
        collection = PivotItems(self, items_showing(self, False))
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def CurrentPage(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import current_page

        return current_page(self)

    @setter('CurrentPage')
    def set_current_page(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_filters import set_current_page

        set_current_page(self, value)

    @member
    def EnableMultiplePageItems(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import multiple_allowed

        return multiple_allowed(self)

    @setter('EnableMultiplePageItems')
    def set_enable_multiple_page_items(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_filters import set_multiple_allowed

        set_multiple_allowed(self, value)

    @method
    def ClearAllFilters(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import clear_field

        clear_field(self)
        return EMPTY

    @member
    def PivotFilters(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_criteria import field_filters

        filters = field_filters(self)
        return filters if Index is MISSING else filters.Item(Index)

    @method
    def ClearLabelFilters(self) -> object:
        from pyopenvba.apps.excel._pivot_criteria import clear_kind

        clear_kind(self, labels=True, values=False)
        return EMPTY

    @method
    def ClearValueFilters(self) -> object:
        from pyopenvba.apps.excel._pivot_criteria import clear_kind

        clear_kind(self, labels=False, values=True)
        return EMPTY

    @method
    def ClearManualFilter(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import clear_field

        clear_field(self, manual_only=True)
        return EMPTY

    @method
    def AutoSort(self, Order: object, Field: object, PivotLine: object = MISSING,
                 CustomSubtotal: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_sorts import auto_sort

        auto_sort(self, Order, Field, PivotLine, CustomSubtotal)
        return EMPTY

    @member
    def AutoSortOrder(self) -> object:
        from pyopenvba.apps.excel._pivot_sorts import sort_order

        return VBAInt(sort_order(self), 'Long')

    @member
    def AutoSortField(self) -> object:
        from pyopenvba.apps.excel._pivot_sorts import sort_field

        return sort_field(self)

    def listed_items(self) -> list[PivotItem]:
        """The field's items in the order the table lists them, or sorted for a field it has never listed."""
        if self.data:
            return []
        if self.index == -2:
            return [PivotItem(self, XmlElement(NS + 's', {'v': entry.get('name', '')}), position)
                    for position, entry in enumerate(self.table.root.findall(f'{NS}dataFields/{NS}dataField'), 1)]
        fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
        shared = fields[self.index].find(NS + 'sharedItems')
        entries = list(shared) if shared is not None else []
        if not entries:
            # Numeric fields can omit sharedItems and store values in each record.
            package = self.table.sheet.book.package
            rows = self.table.cache.records
            if rows is None and package is not None:
                records = related_parts(package, self.table.cache.part, 'pivotCacheRecords')
                if not records:
                    records = related_parts(package, self.table.cache.part, 'richPivotRecords')
                if records:
                    rows = xml_tree(package.read(records[0][1]))
            if rows is not None:
                seen: set[tuple[str, tuple[tuple[str, str], ...]]] = set()
                for row in rows:
                    if self.index < len(row):
                        value = row[self.index]
                        key = (value.tag, tuple(sorted(value.attrib.items())))
                        if key not in seen:
                            seen.add(key)
                            entries.append(value)
        items: list[PivotItem] = []
        ordered = self.root.findall(f'{NS}items/{NS}item')
        if ordered:
            for item in ordered:
                if item.get('t', 'data') != 'data':
                    continue
                index = int(item.get('x', '0'))
                if index < len(entries):
                    items.append(PivotItem(self, entries[index], len(items) + 1, index))
        else:
            from pyopenvba.apps.excel._pivot_filters import sorted_places

            items = [PivotItem(self, entries[place], position, place)
                     for position, place in enumerate(sorted_places(entries), 1)]
        return items


class PivotItems(_Collection):
    vba_type_name = 'PivotItems'

    def __init__(self, parent: ExcelObject, items: list[object], *, first_for_names: bool = False) -> None:
        super().__init__(parent, items)
        self.first_for_names = first_for_names

    def vba_lookup(self, index: object, items: list[object]) -> object:
        found = super().vba_lookup(index, items)
        return items[0] if self.first_for_names and isinstance(index, str) else found


class PivotItem(ExcelObject):
    """An item of a pivot field: its value in the cache, or None for the (All) a page field showing every item
    returns as its CurrentPage, its place in the field's list from 1, and the value's place in the cache."""

    vba_type_name = 'PivotItem'

    def __init__(self, field: PivotField, root: XmlElement | None, position: int, x: int = -1) -> None:
        self.field, self.root, self.position, self.x = field, root, position, x

    @member
    def Name(self) -> object:
        if self.root is None:
            return '(All)'
        if self.root.tag == NS + 'm':
            return '(blank)'
        value = self.root.get('v', '')
        if self.root.tag == NS + 'b':
            return 'TRUE' if value == '1' else 'FALSE'
        if self.root.tag == NS + 'n':
            return number_text(float(value))
        if self.root.tag == NS + 'd':
            raise VBAUnsupportedError('Date PivotItem captions are not implemented')
        return decode_text(value)

    @member
    def Caption(self) -> object:
        return self.Name()

    @member
    def Value(self) -> object:
        return self.Name()

    @member
    def SourceName(self) -> object:
        return self.Name()

    @member
    def Position(self) -> object:
        from pyopenvba.apps.excel._pivot_sorts import position

        found = position(self)
        return VBAInt(found, 'Long') if isinstance(found, int) else found

    @setter('Position')
    def set_position(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_sorts import set_position

        set_position(self, value)

    @member
    def Visible(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import visible

        return visible(self)

    @setter('Visible')
    def set_visible(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_filters import set_visible

        set_visible(self, value)

    @member
    def Parent(self) -> object:
        return self.field
