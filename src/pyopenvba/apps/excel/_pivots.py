"""Runtime PivotTable views and worksheet-source cache/report creation.

Relationship traversal follows pyOfficeEditor's pivot reader. Untouched XML stays
in the package unchanged. Worksheet sources can build caches, empty reports and
compact aggregate reports and multi-axis field ranges. Refresh, grouping and arbitrary layout edits remain incomplete.
"""
from __future__ import annotations

import posixpath
import re
from copy import deepcopy
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._xstring import decode as decode_text
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBADate, VBAErrorValue, VBAInt, error, string_order, to_integer, to_text, to_bool, utf16_text

if TYPE_CHECKING:
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


class PivotItemList(_Collection):
    """The items a pivot cell's row or column stands at, PivotCell.RowItems and ColumnItems (see _pivot_cells)."""

    vba_type_name = 'PivotItemList'


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

    @member
    def Application(self) -> object:
        return self.book.application

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
        #: The calculated fields, by their cache names, whose formulas a macro gave in this session; Excel reads a
        #: file's again as it opens it (see _pivot_calculated).
        self.given_formulas: set[str] = set()
        #: The formulas a macro gave calculated items in this session, as they read back, by each item's field and
        #: place; Excel reads a file's without their spaces (see _pivot_items).
        self.given_items: dict[tuple[int, int], str] = {}

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
        #: ManualUpdate, which no file keeps.
        self.manual = False

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member(default=True, name="_Default")
    def _Default(self) -> object:
        """Its name, which Value reads too (tests/fixtures/object_members.json)."""
        return self.vba_get("Name")

    @member
    def Value(self) -> object:
        return self.vba_get("Name")

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

    @method
    def ShowPages(self, PageField: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_pages import show_pages

        return show_pages(self, PageField)

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

    @setter('RowGrand')
    def set_row_grand(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_table_flag

        set_table_flag(self, 'rowGrandTotals', value)

    @setter('ColumnGrand')
    def set_column_grand(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_table_flag

        set_table_flag(self, 'colGrandTotals', value)

    @method
    def RowAxisLayout(self, RowLayout: object) -> object:
        from pyopenvba.apps.excel._pivot_forms import row_axis_layout

        row_axis_layout(self, RowLayout)
        return EMPTY

    @method
    def RepeatAllLabels(self, Repeat: object) -> object:
        from pyopenvba.apps.excel._pivot_forms import repeat_all_labels

        repeat_all_labels(self, Repeat)
        return EMPTY

    @member
    def LayoutRowDefault(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import layout_row_default

        return VBAInt(layout_row_default(self), 'Long')

    @setter('LayoutRowDefault')
    def set_layout_row_default(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_layout_row_default

        set_layout_row_default(self, value)

    @member
    def CompactLayoutRowHeader(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import caption

        return caption(self, 'rowHeaderCaption', 'Row Labels')

    @setter('CompactLayoutRowHeader')
    def set_compact_layout_row_header(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_caption

        set_caption(self, 'rowHeaderCaption', value)

    @member
    def CompactLayoutColumnHeader(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import caption

        return caption(self, 'colHeaderCaption', 'Column Labels')

    @setter('CompactLayoutColumnHeader')
    def set_compact_layout_column_header(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_caption

        set_caption(self, 'colHeaderCaption', value)

    @member
    def GrandTotalName(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import caption

        return caption(self, 'grandTotalCaption', 'Grand Total')

    @setter('GrandTotalName')
    def set_grand_total_name(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_caption

        set_caption(self, 'grandTotalCaption', value)

    @member
    def DisplayFieldCaptions(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import display_field_captions

        return display_field_captions(self)

    @setter('DisplayFieldCaptions')
    def set_display_field_captions(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_display_field_captions

        set_display_field_captions(self, value)

    @member
    def CompactRowIndent(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import compact_indent

        return VBAInt(compact_indent(self.entry.root), 'Long')

    @setter('CompactRowIndent')
    def set_compact_row_indent(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_compact_indent

        set_compact_indent(self, value)

    @member
    def MergeLabels(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import merge_labels

        return merge_labels(self)

    @setter('MergeLabels')
    def set_merge_labels(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_merge_labels

        set_merge_labels(self, value)

    @member
    def DataLabelRange(self) -> object:
        from pyopenvba.apps.excel._pivot_ranges import data_label_range

        return data_label_range(self)

    @member
    def TableStyle2(self) -> object:
        from pyopenvba.apps.excel._pivot_options import table_style

        return table_style(self)

    @setter('TableStyle2')
    def set_table_style2(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_table_style

        set_table_style(self, value)

    @member
    def NullString(self) -> object:
        from pyopenvba.apps.excel._pivot_options import caption

        return caption(self, 'missingCaption')

    @setter('NullString')
    def set_null_string(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_caption

        set_caption(self, 'missingCaption', value)

    @member
    def DisplayNullString(self) -> object:
        return self.entry.root.get('showMissing', '1') not in ('0', 'false')

    @setter('DisplayNullString')
    def set_display_null_string(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_display

        set_display(self, 'showMissing', value, True)

    @member
    def ErrorString(self) -> object:
        from pyopenvba.apps.excel._pivot_options import caption

        return caption(self, 'errorCaption')

    @setter('ErrorString')
    def set_error_string(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_caption

        set_caption(self, 'errorCaption', value)

    @member
    def DisplayErrorString(self) -> object:
        return self.entry.root.get('showError', '0') in ('1', 'true')

    @setter('DisplayErrorString')
    def set_display_error_string(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_display

        set_display(self, 'showError', value, False)

    @member
    def ManualUpdate(self) -> object:
        return self.entry.manual

    @setter('ManualUpdate')
    def set_manual_update(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_manual

        set_manual(self, value)

    @member
    def Location(self) -> object:
        # Excel reads no Location back, error 445 (tests/fixtures/pivot_options/ defaults).
        raise error(445)

    @setter('Location')
    def set_location(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_location

        set_location(self, value)

    @member
    def Tag(self) -> object:
        from pyopenvba.apps.excel._pivot_options import tag

        return tag(self)

    @setter('Tag')
    def set_tag(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_tag

        set_tag(self, value)

    @member
    def AlternativeText(self) -> object:
        from pyopenvba.apps.excel._pivot_options import alternative

        return alternative(self, 'altText')

    @setter('AlternativeText')
    def set_alternative_text(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_alternative

        set_alternative(self, 'altText', value)

    @member
    def Summary(self) -> object:
        from pyopenvba.apps.excel._pivot_options import alternative

        return alternative(self, 'altTextSummary')

    @setter('Summary')
    def set_summary(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_alternative

        set_alternative(self, 'altTextSummary', value)

    @member
    def PageFieldOrder(self) -> object:
        from pyopenvba.apps.excel._pivot_options import page_order

        return VBAInt(page_order(self), 'Long')

    @setter('PageFieldOrder')
    def set_page_field_order(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_page_order

        set_page_order(self, value)

    @member
    def PageFieldWrapCount(self) -> object:
        from pyopenvba.apps.excel._pivot_options import page_wrap

        return VBAInt(page_wrap(self), 'Long')

    @setter('PageFieldWrapCount')
    def set_page_field_wrap_count(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_page_wrap

        set_page_wrap(self, value)

    @member
    def SaveData(self) -> object:
        return self.cache.root.get('saveData', '1') not in ('0', 'false')

    @setter('SaveData')
    def set_save_data(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_save_data

        set_save_data(self, value)

    @member
    def ShowValuesRow(self) -> object:
        from pyopenvba.apps.excel._pivot_options import shows_values_row

        return shows_values_row(self)

    @setter('ShowValuesRow')
    def set_show_values_row(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_shows_values_row

        set_shows_values_row(self, value)

    @member
    def SubtotalHiddenPageItems(self) -> object:
        return self.entry.root.get('subtotalHiddenItems', '0') in ('1', 'true')

    @setter('SubtotalHiddenPageItems')
    def set_subtotal_hidden_page_items(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_kept_on

        set_kept_on(self, value, 'subtotalHiddenItems')

    @member
    def DisplayImmediateItems(self) -> object:
        return self.entry.root.get('showItems', '1') not in ('0', 'false')

    @setter('DisplayImmediateItems')
    def set_display_immediate_items(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_kept_on

        set_kept_on(self, value, 'showItems')

    @member
    def DisplayEmptyRow(self) -> object:
        # A table not from a cube has no empty rows to show, error 1004 (tests/fixtures/pivot_options/ flags).
        raise error(1004)

    @setter('DisplayEmptyRow')
    def set_display_empty_row(self, value: object) -> None:
        raise error(1004)

    @member
    def DisplayEmptyColumn(self) -> object:
        raise error(1004)

    @setter('DisplayEmptyColumn')
    def set_display_empty_column(self, value: object) -> None:
        raise error(1004)

    @method
    def AddFields(self, RowFields: object = MISSING, ColumnFields: object = MISSING, PageFields: object = MISSING,
                  AddToTable: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_options import add_fields

        return add_fields(self, RowFields, ColumnFields, PageFields, AddToTable)

    @method
    def ClearTable(self) -> object:
        from pyopenvba.apps.excel._pivot_options import clear_table

        clear_table(self)
        return EMPTY

    @method
    def Update(self) -> object:
        from pyopenvba.apps.excel._pivot_filters import rerender

        rerender(self.entry, deepcopy(self.entry.root))
        return EMPTY

    @member
    def VisibleFields(self, Index: object = MISSING) -> object:
        fields: list[object] = []
        for index, field in enumerate(self.root.findall(f'{NS}pivotFields/{NS}pivotField')):
            view = PivotField(self, index, field)
            if int(view.Orientation()) in (1, 2, 3):
                fields.append(view)
        for axis in ('rowFields', 'colFields'):
            if any(field.get('x') == '-2' for field in self.root.findall(f'{NS}{axis}/{NS}field')):
                fields.append(self._data_pivot_field(axis))
        fields.extend(PivotField(self, int(field.get('fld', '0')), field, True)
                      for field in self.root.findall(f'{NS}dataFields/{NS}dataField'))
        collection = PivotFields(self, fields)
        return collection if Index is MISSING else collection.Item(Index)

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

    def values_field(self) -> PivotField:
        """The Values field, on the axis that shows it or on none, which a table names even without data fields
        (tests/fixtures/pivot_cells/)."""
        for axis in ('rowFields', 'colFields'):
            if any(field.get('x') == '-2' for field in self.root.findall(f'{NS}{axis}/{NS}field')):
                return self._data_pivot_field(axis)
        return self._data_pivot_field(None)

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

    @method
    def CalculatedFields(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_calculated import CalculatedFields

        fields = CalculatedFields(self)
        return fields if Index is MISSING else fields.Item(Index)


class PivotField(ExcelObject):
    vba_type_name = 'PivotField'

    def __init__(self, table: PivotTable, index: int, root: XmlElement, data: bool = False) -> None:
        self.table, self.index, self.root, self.data = table, index, root, data

    @member
    def Application(self) -> object:
        return self.table.sheet.book.application

    @member(default=True, name="_Default")
    def _Default(self) -> object:
        """Its name, which Value reads too (tests/fixtures/object_members.json)."""
        return self.vba_get("Name")

    @member
    def Value(self) -> object:
        return self.vba_get("Name")

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

    @setter('Name')
    def set_name(self, value: object) -> None:
        self._rename(value, caption=False)

    @member
    def Caption(self) -> object:
        return self.Name()

    @setter('Caption')
    def set_caption(self, value: object) -> None:
        self._rename(value, caption=True)

    def _rename(self, value: object, *, caption: bool) -> None:
        from pyopenvba.apps.excel._pivot_names import rename_data, rename_field, rename_values

        if self.index == -2:
            rename_values(self, value)
        elif self.data:
            rename_data(self, value)
        else:
            rename_field(self, value, caption=caption)

    @member
    def Formula(self) -> object:
        from pyopenvba.apps.excel._pivot_calculated import formula

        return formula(self)

    @setter('Formula')
    def set_formula(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_calculated import set_formula

        set_formula(self, value)

    @member
    def StandardFormula(self) -> object:
        return self.Formula()

    @setter('StandardFormula')
    def set_standard_formula(self, value: object) -> None:
        self.set_formula(value)

    @member
    def IsCalculated(self) -> object:
        from pyopenvba.apps.excel._pivot_calculated import is_calculated_field

        return is_calculated_field(self)

    @member
    def DataType(self) -> object:
        from pyopenvba.apps.excel._pivot_calculated import data_type

        return data_type(self)

    @method
    def Delete(self) -> object:
        from pyopenvba.apps.excel._pivot_calculated import delete

        delete(self)
        return EMPTY

    @member
    def ShowDetail(self) -> object:
        from pyopenvba.apps.excel._pivot_details import field_detail

        return field_detail(self)

    @setter('ShowDetail')
    def set_show_detail(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_details import set_field_detail

        set_field_detail(self, value)

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

    @setter('Function')
    def set_function(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_aggregate import set_function

        set_function(self, value)

    @member
    def NumberFormat(self) -> object:
        # The Values field and a field of text have none, error 1004 (tests/fixtures/pivot_date_filters/).
        if self.index == -2:
            raise error(1004)
        if not self.data:
            fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
            shared = fields[self.index].find(NS + 'sharedItems') if self.index < len(fields) else None
            kinds: set[str] = {entry.tag for entry in shared} if shared is not None else set()
            if kinds and kinds <= {NS + 's', NS + 'm'}:
                raise error(1004)
        identifier = int(self.root.get('numFmtId', '0'))
        return self.table.sheet.book.stylesheet.number_formats.get(identifier, 'General')

    @setter('NumberFormat')
    def set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_calculations import set_number_format

        set_number_format(self, value)

    @member
    def Calculation(self) -> object:
        from pyopenvba.apps.excel._pivot_calculations import calculation

        return calculation(self)

    @setter('Calculation')
    def set_calculation(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_calculations import set_calculation

        set_calculation(self, value)

    @member
    def BaseField(self) -> object:
        from pyopenvba.apps.excel._pivot_calculations import base_field

        return base_field(self)

    @setter('BaseField')
    def set_base_field(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_calculations import set_base_field

        set_base_field(self, value)

    @member
    def BaseItem(self) -> object:
        from pyopenvba.apps.excel._pivot_calculations import base_item

        return base_item(self)

    @setter('BaseItem')
    def set_base_item(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_calculations import set_base_item

        set_base_item(self, value)

    @member
    def LayoutForm(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import layout_form

        return VBAInt(layout_form(self), 'Long')

    @setter('LayoutForm')
    def set_layout_form(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_layout_form

        set_layout_form(self, value)

    @member
    def LayoutCompactRow(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import compact_row

        return compact_row(self)

    @setter('LayoutCompactRow')
    def set_layout_compact_row(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_compact_row

        set_compact_row(self, value)

    @member
    def LayoutSubtotalLocation(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import subtotal_location

        return VBAInt(subtotal_location(self), 'Long')

    @setter('LayoutSubtotalLocation')
    def set_layout_subtotal_location(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_subtotal_location

        set_subtotal_location(self, value)

    @member
    def LayoutBlankLine(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import flag_of

        return flag_of(self, 'insertBlankRow')

    @setter('LayoutBlankLine')
    def set_layout_blank_line(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_blank_line

        set_blank_line(self, value)

    @member
    def LayoutPageBreak(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import flag_of

        return flag_of(self, 'insertPageBreak')

    @setter('LayoutPageBreak')
    def set_layout_page_break(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_page_break

        set_page_break(self, value)

    @member
    def RepeatLabels(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import repeat_labels

        return repeat_labels(self)

    @setter('RepeatLabels')
    def set_repeat_labels(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_repeat_labels

        set_repeat_labels(self, value)

    @member
    def Subtotals(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_forms import subtotals

        return subtotals(self, Index)

    @setter('Subtotals')
    def set_subtotals(self, Index: object, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_subtotals

        set_subtotals(self, Index, value)

    @member
    def SubtotalName(self) -> object:
        from pyopenvba.apps.excel._pivot_forms import subtotal_name

        return subtotal_name(self)

    @setter('SubtotalName')
    def set_subtotal_name(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_forms import set_subtotal_name

        set_subtotal_name(self, value)

    def _range(self, label: bool) -> object:
        from pyopenvba.apps.excel._pivot_ranges import field_range

        return field_range(self, label)

    @member
    def DataRange(self) -> object:
        return self._range(False)

    @member
    def LabelRange(self) -> object:
        return self._range(True)

    @method
    def CalculatedItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_items import CalculatedItems

        items = CalculatedItems(self)
        return items if Index is MISSING else items.Item(Index)

    @member
    def ParentField(self) -> object:
        from pyopenvba.apps.excel._pivot_groups import parent_field

        return parent_field(self)

    @member
    def ChildField(self) -> object:
        from pyopenvba.apps.excel._pivot_groups import child_field

        return child_field(self)

    @member
    def GroupLevel(self) -> object:
        from pyopenvba.apps.excel._pivot_groups import group_level

        return group_level(self)

    @member
    def TotalLevels(self) -> object:
        from pyopenvba.apps.excel._pivot_groups import total_levels

        return total_levels(self)

    @member
    def ParentItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_groups import parent_items

        collection = parent_items(self)
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def PivotItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_calculated import calculated_view
        from pyopenvba.apps.excel._pivot_filters import first_for_names, list_on_reading

        if calculated_view(self):
            # A calculated field has no items, error 1004 (tests/fixtures/pivot_calculated_fields/
            # items_of_calculated).
            raise error(1004)
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
        from pyopenvba.apps.excel._pivot_groups import item_entries, natural_places

        fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
        entries = item_entries(fields[self.index])
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

            places = natural_places(fields, self.index) if len(entries) == len(item_entries(fields[self.index])) \
                else sorted_places(entries)
            items = [PivotItem(self, entries[place], position, place) for position, place in enumerate(places, 1)]
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
    def Application(self) -> object:
        return self.field.table.sheet.book.application

    @member(default=True, name="_Default")
    def _Default(self) -> object:
        """Its name (tests/fixtures/object_members.json)."""
        return self.vba_get("Name")

    @member
    def Name(self) -> object:
        """The item's caption where it was given one, else its value's name (tests/fixtures/pivot_names/)."""
        if self.x >= 0 and self.field.index >= 0 and not self.field.data:
            from pyopenvba.apps.excel._pivot_names import caption_of

            given = caption_of(self.field, self.x)
            if given is not None:
                return given
        return self.SourceName()

    @setter('Name')
    def set_name(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_names import rename_item

        rename_item(self, value)

    @setter('Caption')
    def set_caption(self, value: object) -> None:
        self.set_name(value)

    @setter('Value')
    def set_value(self, value: object) -> None:
        self.set_name(value)

    @member
    def SourceName(self) -> object:
        """The item's value as its field shows it, whatever caption it was given."""
        if self.root is None:
            return '(All)'
        if self.root.tag == NS + 'm':
            return '(blank)'
        if self.x >= 0 and self.field.index >= 0 and not self.field.data:
            from pyopenvba.apps.excel._pivot_groups import item_name

            fields = self.field.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
            named = item_name(fields, self.field.index, self.x) if self.field.index < len(fields) else None
            if named is not None:
                return named
        value = self.root.get('v', '')
        if self.root.tag == NS + 'b':
            return 'TRUE' if value == '1' else 'FALSE'
        if self.root.tag in (NS + 'n', NS + 'd') and self.x >= 0:
            # A number or a date shows in its field's format (tests/fixtures/pivot_date_filters/).
            from pyopenvba.apps.excel._pivot_layout import field_format, item_caption

            table = self.field.table
            fields = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
            cache_field = fields[self.field.index] if 0 <= self.field.index < len(fields) else None
            return item_caption(self.root, field_format(table.sheet.book.stylesheet.number_formats,
                                                        self.field.root, cache_field))
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
    def Position(self) -> object:
        from pyopenvba.apps.excel._pivot_sorts import position

        found = position(self)
        return VBAInt(found, 'Long') if isinstance(found, int) else found

    @setter('Position')
    def set_position(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_sorts import set_position

        set_position(self, value)

    @member
    def ShowDetail(self) -> object:
        from pyopenvba.apps.excel._pivot_details import item_detail

        return item_detail(self)

    @setter('ShowDetail')
    def set_show_detail(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_details import set_item_detail

        set_item_detail(self, value)

    @member
    def LabelRange(self) -> object:
        from pyopenvba.apps.excel._pivot_ranges import item_range

        return item_range(self, True)

    @member
    def DataRange(self) -> object:
        from pyopenvba.apps.excel._pivot_ranges import item_range

        return item_range(self, False)

    @member
    def ParentItem(self) -> object:
        from pyopenvba.apps.excel._pivot_groups import parent_item

        return parent_item(self)

    @member
    def ChildItems(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._pivot_groups import child_items

        return child_items(self, Index)

    @member
    def IsCalculated(self) -> object:
        from pyopenvba.apps.excel._pivot_items import is_calculated

        return is_calculated(self)

    @member
    def Formula(self) -> object:
        from pyopenvba.apps.excel._pivot_items import formula_of

        return formula_of(self)

    @setter('Formula')
    def set_formula(self, value: object) -> None:
        from pyopenvba.apps.excel._pivot_items import set_formula

        set_formula(self, value)

    @member
    def StandardFormula(self) -> object:
        return self.Formula()

    @setter('StandardFormula')
    def set_standard_formula(self, value: object) -> None:
        self.set_formula(value)

    @method
    def Delete(self) -> object:
        from pyopenvba.apps.excel._pivot_items import delete

        delete(self)
        return EMPTY

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


#: Options a pivot table keeps as attributes and no cell shows: member -> (attribute, its default, whether the
#: member reads its opposite) (tests/fixtures/pivot_options/ flags).
TABLE_FLAGS: dict[str, tuple[str, bool, bool]] = {
    'DisplayContextTooltips': ('showDataTips', True, False),
    'EnableDataValueEditing': ('editData', False, False),
    'EnableDrilldown': ('enableDrill', True, False),
    'EnableFieldDialog': ('enableFieldProperties', True, False),
    'EnableFieldList': ('disableFieldList', False, True),
    'EnableWizard': ('enableWizard', True, False),
    'FieldListSortAscending': ('fieldListSortAscending', False, False),
    'HasAutoFormat': ('useAutoFormatting', False, False),
    'InGridDropZones': ('gridDropZones', False, False),
    'PreserveFormatting': ('preserveFormatting', True, False),
    'PrintTitles': ('fieldPrintTitles', False, False),
    'RepeatItemsOnEachPrintedPage': ('itemPrintTitles', False, False),
    'ShowDrillIndicators': ('showDrill', True, False),
    'SortUsingCustomLists': ('customListSort', True, False),
}
#: The parts of its style a table draws: member -> pivotTableStyleInfo's attribute.
STYLE_FLAGS = {'ShowTableStyleRowHeaders': 'showRowHeaders', 'ShowTableStyleColumnHeaders': 'showColHeaders',
               'ShowTableStyleRowStripes': 'showRowStripes', 'ShowTableStyleColumnStripes': 'showColStripes',
               'ShowTableStyleLastColumn': 'showLastColumn'}


def _option_member(name: str, style: bool) -> MemberSpec:
    """One of the options above as a property of PivotTable."""

    def get(table: PivotTable) -> object:
        from pyopenvba.apps.excel._pivot_options import flag, style_flag

        return style_flag(table, name) if style else flag(table, name)

    def put(table: PivotTable, value: object) -> None:
        from pyopenvba.apps.excel._pivot_options import set_flag_member, set_style_flag

        if style:
            set_style_flag(table, name, value)
        else:
            set_flag_member(table, name, value)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _name in TABLE_FLAGS:
    PivotTable.vba_add_member(_option_member(_name, False))
for _name in STYLE_FLAGS:
    PivotTable.vba_add_member(_option_member(_name, True))
