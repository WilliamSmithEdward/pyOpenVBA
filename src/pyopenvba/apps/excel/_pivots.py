"""Runtime views of PivotTables already stored in an OOXML workbook.

Relationship traversal follows pyOfficeEditor's pivot reader. Original XML stays
in the package unchanged. Creation, refresh, grouping and layout edits remain
explicitly unsupported until the engine can build and persist their caches.
"""
from __future__ import annotations

import posixpath
import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import MISSING, VBAErrorValue, VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivot_reports import PivotReport
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.powerquery._opc import OpcFile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def decode_text(value: str) -> str:
    return re.sub(r'_x([0-9A-Fa-f]{4})_', lambda match: chr(int(match.group(1), 16)), value)


def _parts(package: OpcFile, source: str, kind: str) -> list[tuple[str, str]]:
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


def read(book: Workbook) -> None:
    """Read caches in workbook order, then each sheet's report relationships."""
    package = book.package
    by_part: dict[str, PivotCache] = {}
    by_id = dict(_parts(package, 'xl/workbook.xml', 'pivotCacheDefinition'))
    root = xml_tree(package.read('xl/workbook.xml'))
    for entry in root.findall(f'{NS}pivotCaches/{NS}pivotCache'):
        part = by_id.get(entry.get(REL + 'id', ''), '')
        if not part:
            continue
        cache = PivotCache(book, part, xml_tree(package.read(part)), len(book.pivot_caches) + 1)
        book.pivot_caches.append(cache)
        by_part[part] = cache
    for sheet in book.sheets_:
        reports = _parts(package, sheet.part_name, 'pivotTable')
        # Native reopen controls: reversing XML or ZIP order has no effect;
        # swapping relationship ids swaps the collection order. Foreign ids
        # retain their source order until their ordering is measured.
        if all(re.fullmatch(r'rId[0-9]+', identifier) for identifier, _ in reports):
            reports.sort(key=lambda pair: int(pair[0][3:]))
        for _, part in reports:
            caches = _parts(package, part, 'pivotCacheDefinition')
            cache = by_part.get(caches[0][1]) if caches else None
            if cache is not None:
                sheet.pivot_tables.append(PivotTable(sheet, cache, xml_tree(package.read(part))))


class _Collection(VBACollection, ExcelObject):
    def __init__(self, parent: ExcelObject, items: list[object]) -> None:
        self.parent = parent
        self.items = items

    def vba_items(self) -> list[object]:
        return self.items

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                if isinstance(item, ExcelObject) and str(item.vba_get('Name')).casefold() == index.casefold():
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

    def vba_lookup(self, index: object, items: list[object]) -> object:
        # Cache indices are numeric; numeric strings receive normal Long coercion.
        return super().vba_lookup(to_integer(index, 'Long'), items)


class PivotTables(_Collection):
    vba_type_name = 'PivotTables'


class PivotFields(_Collection):
    vba_type_name = 'PivotFields'


class PivotCache(ExcelObject):
    vba_type_name = 'PivotCache'

    def __init__(self, book: Workbook, part: str, root: XmlElement, index: int) -> None:
        self.book, self.part, self.root, self.index = book, part, root, index

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

    @member
    def RefreshOnFileOpen(self) -> object:
        return self.root.get('refreshOnLoad', '0') in ('1', 'true')

    @member
    def Version(self) -> object:
        return VBAInt(int(self.root.get('createdVersion', '0')), 'Long')


class PivotTable(ExcelObject):
    vba_type_name = 'PivotTable'

    def __init__(self, sheet: Worksheet, cache: PivotCache, root: XmlElement) -> None:
        self.sheet, self.cache, self.root = sheet, cache, root
        # Report definitions are immutable until layout/refresh writers exist.
        # Cache metadata only: GETPIVOTDATA reads each result cell afresh.
        self.report_metadata: PivotReport | None = None
        self.report_loaded = False

    def _range(self, kind: str) -> object:
        location = self.root.find(NS + 'location')
        if location is None:
            raise error(1004)
        area = parse_area(location.get('ref', 'A1'))
        dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
        if kind == 'data':
            area = Area(area.top + dr, area.left + dc, area.bottom, area.right)
        elif kind == 'row':
            if not self.root.findall(f'{NS}rowFields/{NS}field') or dc == 0:
                raise error(1004)
            area = Area(area.top, area.left, area.bottom, area.left + dc - 1)
        elif kind == 'column':
            if not self.root.findall(f'{NS}colFields/{NS}field') or dr == 0:
                raise error(1004)
            area = Area(area.top, area.left + dc, area.top + dr - 1, area.right)
        return Range(self.sheet, [Area(area.top, area.left, area.bottom, area.right, self.sheet.name)])

    @member
    def Name(self) -> object:
        return decode_text(self.root.get('name', ''))

    @member
    def Parent(self) -> object:
        return self.sheet

    @method
    def PivotCache(self) -> object:
        return self.cache

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
            raise VBAUnsupportedError('PivotTable.TableRange2 with page fields is not implemented')
        return self._range('all')

    @member
    def PageRange(self) -> object:
        if self.root.findall(f'{NS}pageFields/{NS}pageField'):
            raise VBAUnsupportedError('PivotTable.PageRange with page fields is not implemented')
        raise error(1004)

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
        fields = PivotFields(self, items)
        return fields if index is MISSING else fields.Item(index)

    def _data_pivot_field(self, axis: str) -> PivotField:
        field = XmlElement(NS + 'pivotField', {'name': self.root.get('dataCaption', 'Values'),
                                              'axis': 'axisRow' if axis == 'rowFields' else 'axisCol'})
        return PivotField(self, -2, field)

    @member
    def DataPivotField(self) -> object:
        for axis in ('rowFields', 'colFields'):
            if any(field.get('x') == '-2' for field in self.root.findall(f'{NS}{axis}/{NS}field')):
                return self._data_pivot_field(axis)
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

    @member
    def SourceName(self) -> object:
        if self.index == -2:
            return VBAErrorValue(2042)
        fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
        return decode_text(fields[self.index].get('name', '')) if self.index < len(fields) else ''

    @member
    def Name(self) -> object:
        name = self.root.get('name')
        return decode_text(name) if name is not None else self.SourceName()

    @member
    def Caption(self) -> object:
        return self.Name()

    @member
    def Orientation(self) -> object:
        return VBAInt(4 if self.data else
                      {'axisRow': 1, 'axisCol': 2, 'axisPage': 3}.get(self.root.get('axis', ''), 0), 'Long')

    @member
    def Parent(self) -> object:
        return self.table

    @member
    def Position(self) -> object:
        axis = {1: 'rowFields', 2: 'colFields', 3: 'pageFields', 4: 'dataFields'}.get(int(self.Orientation()))
        if axis is not None:
            for position, field in enumerate(self.table.root.findall(f'{NS}{axis}/*'), 1):
                if (self.data and field is self.root) or (not self.data and int(field.get('x', field.get('fld', '-1'))) == self.index):
                    return VBAInt(position, 'Long')
        return VBAErrorValue(2042)

    @member
    def Function(self) -> object:
        if not self.data:
            raise error(1004)
        functions = {'sum': -4157, 'count': -4112, 'average': -4106, 'max': -4136, 'min': -4139,
                     'product': -4149, 'countNums': -4113, 'stdDev': -4155, 'stdDevp': -4156,
                     'var': -4164, 'varp': -4165}
        return VBAInt(functions[self.root.get('subtotal', 'sum')], 'Long')

    @member
    def NumberFormat(self) -> object:
        identifier = int(self.root.get('numFmtId', '0'))
        return self.table.sheet.book.stylesheet.number_formats.get(identifier, 'General')

    def _range(self, label: bool) -> object:
        root = self.table.root
        row_fields = root.findall(f'{NS}rowFields/{NS}field')
        if len(row_fields) != 1 or root.find(NS + 'colFields') is not None or root.find(NS + 'pageFields') is not None:
            raise VBAUnsupportedError('PivotField ranges for multi-axis reports are not implemented')
        location = root.find(NS + 'location')
        if location is None:
            raise error(1004)
        area = parse_area(location.get('ref', 'A1'))
        dr, dc = int(location.get('firstDataRow', '0')), int(location.get('firstDataCol', '0'))
        if self.data:
            column = area.left + dc + int(self.Position()) - 1
        elif int(self.Orientation()) == 1:
            column = area.left
        else:
            raise error(1004)
        if label:
            top = bottom = area.top + dr - 1
        else:
            top = area.top + dr
            rows = root.findall(f'{NS}rowItems/{NS}i')
            bottom = area.bottom - int(bool(rows and rows[-1].get('t') == 'grand'))
        if bottom < top:
            raise error(1004)
        return Range(self.table.sheet, [Area(top, column, bottom, column, self.table.sheet.name)])

    @member
    def DataRange(self) -> object:
        return self._range(False)

    @member
    def LabelRange(self) -> object:
        return self._range(True)

    @member
    def PivotItems(self, Index: object = MISSING) -> object:
        if self.index == -2:
            raise VBAUnsupportedError('Values-axis PivotItems are not implemented')
        fields = self.table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
        shared = fields[self.index].find(NS + 'sharedItems')
        entries = list(shared) if shared is not None else []
        if not entries:
            # Numeric fields can omit sharedItems and store values in each record.
            package = self.table.sheet.book.package
            records = _parts(package, self.table.cache.part, 'pivotCacheRecords')
            if records:
                seen: set[tuple[str, tuple[tuple[str, str], ...]]] = set()
                for row in xml_tree(package.read(records[0][1])):
                    if self.index < len(row):
                        value = row[self.index]
                        key = (value.tag, tuple(sorted(value.attrib.items())))
                        if key not in seen:
                            seen.add(key)
                            entries.append(value)
        items: list[object] = []
        ordered = self.root.findall(f'{NS}items/{NS}item')
        if ordered:
            for item in ordered:
                if item.get('t', 'data') != 'data':
                    continue
                index = int(item.get('x', '0'))
                if index < len(entries):
                    items.append(PivotItem(self, entries[index], len(items) + 1, item.get('h') != '1'))
        else:
            items = [PivotItem(self, item, index, True) for index, item in enumerate(entries, 1)]
        collection = PivotItems(self, items)
        return collection if Index is MISSING else collection.Item(Index)


class PivotItems(_Collection):
    vba_type_name = 'PivotItems'


class PivotItem(ExcelObject):
    vba_type_name = 'PivotItem'

    def __init__(self, field: PivotField, root: XmlElement, position: int, visible: bool) -> None:
        self.field, self.root, self.position, self.visible = field, root, position, visible

    @member
    def Name(self) -> object:
        if self.root.tag == NS + 'm':
            return '(blank)'
        value = self.root.get('v', '')
        if self.root.tag == NS + 'b':
            return 'TRUE' if value == '1' else 'FALSE'
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
        return VBAInt(self.position, 'Long')

    @member
    def Visible(self) -> object:
        return self.visible

    @member
    def Parent(self) -> object:
        return self.field
