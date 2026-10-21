"""Measured, persisted worksheet PageSetup configuration; no print rendering.

Attribute mappings and point/inch conversion follow William's pyOfficeEditor
_pagesetup. Only owned attributes are patched; printer relationships, namespace
bindings, unknown children and header/footer markup remain untouched.
"""
from __future__ import annotations

import math
import re
from copy import deepcopy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference, quote_sheet
from pyopenvba._opc_xml import opc_part
from pyopenvba._xml import escape
from pyopenvba.apps.excel._advanced_filter import attribute_values, patch_attribute
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import EMPTY, NULL, VBAInt, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import NameEntry, Worksheet

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
MARGINS = dict(zip(('LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin', 'HeaderMargin', 'FooterMargin'),
                   ('left', 'right', 'top', 'bottom', 'header', 'footer'), strict=True))
DEFAULT_MARGINS = dict(zip(MARGINS.values(), (0.7, 0.7, 0.75, 0.75, 0.3, 0.3), strict=True))
OPTIONS = {'CenterHorizontally': 'horizontalCentered', 'CenterVertically': 'verticalCentered',
           'PrintGridlines': 'gridLines', 'PrintHeadings': 'headings'}
FLAGS = {'BlackAndWhite': 'blackAndWhite', 'Draft': 'draft'}
ENUMS = {'Orientation': ('orientation', {1: 'portrait', 2: 'landscape'}, 1),
         'Order': ('pageOrder', {1: 'downThenOver', 2: 'overThenDown'}, 1),
         'PrintComments': ('cellComments', {-4142: 'none', 1: 'atEnd', 16: 'asDisplayed'}, -4142),
         'PrintErrors': ('errors', {0: 'displayed', 1: 'blank', 2: 'dash', 3: 'NA'}, 0)}
TAIL = ('headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags',
        'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems',
        'tableParts', 'extLst')
FOLLOWERS = {'printOptions': ('pageMargins', 'pageSetup', *TAIL), 'pageMargins': ('pageSetup', *TAIL),
             'pageSetup': TAIL}


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('PageSetup ' + detail + ' is not implemented')


def _opening_end(data: bytes) -> int:
    quote = 0
    for index, char in enumerate(data):
        if quote:
            if char == quote:
                quote = 0
        elif char in (34, 39):
            quote = char
        elif char == 62:
            return index + 1
    raise unsupported('with an unterminated XML opening tag')


def _patch(xml: str, root_name: str, path: tuple[str, ...], values: dict[str, str | None]) -> str:
    assert path
    local = path[-1]
    part = opc_part(xml.encode('utf-8'), NS, root_name, container_path=path[:-1])
    if part is None:
        if all(value is None for value in values.values()):
            return xml
        xml = _patch(xml, root_name, path[:-1], {})
        part = opc_part(xml.encode('utf-8'), NS, root_name, container_path=path[:-1])
        assert part is not None
    matches = [(start, end) for name, _attrs, start, end in part.children if name == NS + '}' + local]
    if len(matches) > 1:
        raise unsupported('with duplicate configuration elements')
    if not matches:
        if values and all(value is None for value in values.values()):
            return xml
        # Excel's absent pageSetup means portrait, but a present element without
        # orientation reopens with Orientation=0. Preserve the blank-sheet state
        # when another setting first materializes the element.
        if local == 'pageSetup':
            values = {'orientation': 'portrait', **values}
        attrs = ''.join(f' {key}="{escape(value)}"' for key, value in values.items() if value is not None)
        child = part.child(local, attrs.strip())
        if local == 'sheetPr':
            at = part.opening_end
        else:
            at = next((start for name, _attrs, start, _end in part.children
                       if name.rsplit('}', 1)[-1] in FOLLOWERS.get(local, ())), part.closing_start)
        return part.insert(at, child).decode('utf-8')
    start, _end = matches[0]
    stop = start + _opening_end(part.data[start:])
    tag = part.data[start:stop].decode('utf-8')
    first = re.match(r'<[^\s/>]+', tag)
    assert first is not None
    tail = len(tag) - (2 if tag.endswith('/>') else 1)
    fragment = tag[first.end():tail]
    for name, value in values.items():
        existing = attribute_values(fragment)
        if (name in existing and existing[name] == value) or (name not in existing and value is None):
            continue
        fragment = patch_attribute(fragment, name, value)
    updated = (tag[:first.end()] + fragment + tag[tail:]).encode('utf-8')
    return (part.data[:start] + updated + part.data[stop:]).decode('utf-8')


def configuration_tag_end(data: bytes) -> int:
    """End of an opening tag already validated in its full XML namespace context."""
    return _opening_end(data)


def patch_configuration(xml: str, root_name: str, path: tuple[str, ...],
                        values: dict[str, str | None]) -> str:
    """Patch owned configuration attributes in their complete namespace context."""
    return _patch(xml, root_name, path, values)


@dataclass
class PageState:
    attributes: dict[tuple[str, ...], dict[str, str]] = field(default_factory=lambda: {})
    changes: dict[tuple[str, ...], dict[str, str | None]] = field(default_factory=lambda: {})


def state(sheet: Worksheet) -> PageState:
    if sheet.page_setup_state is None:
        result = PageState()
        package = sheet.book.package
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            data = package.read(sheet.part_name)
            part = opc_part(data, NS, 'worksheet')
            for name, attrs, _start, _end in part.children:
                if not name.startswith(NS + '}'):
                    continue
                local = name.removeprefix(NS + '}')
                if local in ('pageMargins', 'pageSetup', 'printOptions'):
                    if (local,) in result.attributes:
                        raise unsupported('with duplicate configuration elements')
                    result.attributes[(local,)] = dict(attrs)
            nested = opc_part(data, NS, 'worksheet', container_path=('sheetPr',))
            if nested is not None:
                for name, attrs, _start, _end in nested.children:
                    if name == NS + '}pageSetUpPr':
                        if ('sheetPr', 'pageSetUpPr') in result.attributes:
                            raise unsupported('with duplicate configuration elements')
                        result.attributes[('sheetPr', 'pageSetUpPr')] = dict(attrs)
        sheet.page_setup_state = result
    return sheet.page_setup_state


def with_setup(sheet: Worksheet, xml: str) -> str:
    current = sheet.page_setup_state
    if current is not None:
        for path, values in current.changes.items():
            xml = _patch(xml, 'worksheet', path, values)
    return xml


def _change(sheet: Worksheet, updates: dict[tuple[str, ...], dict[str, str | None]]) -> None:
    current = state(sheet)
    prepared = deepcopy(current)
    for path, values in updates.items():
        if path == ('pageSetup',) and path not in prepared.attributes:
            if all(value is None for value in values.values()):
                continue
            values = {'orientation': 'portrait', **values}
        attrs = prepared.attributes.setdefault(path, {})
        for name, value in values.items():
            if value is None:
                attrs.pop(name, None)
            else:
                attrs[name] = value
        prepared.changes.setdefault(path, {}).update(values)
    package = sheet.book.package
    xml = package.read(sheet.part_name).decode('utf-8') if package is not None and sheet.part_name else f'<worksheet xmlns="{NS}"/>'
    for path, values in updates.items():
        xml = _patch(xml, 'worksheet', path, values)
    current.attributes, current.changes = prepared.attributes, prepared.changes
    sheet.book.saved = False


def _entry(sheet: Worksheet, bare: str) -> NameEntry | None:
    local = (quote_sheet(sheet.name) + '!' + bare).casefold()
    return next((entry for entry in sheet.book.names_.entries if entry.name.casefold() == local), None)


def _print_ranges(sheet: Worksheet, bare: str) -> list[Area]:
    entry = _entry(sheet, bare)
    if entry is None:
        return []
    try:
        areas = parse_reference(entry.refers_to.removeprefix('='))
    except ValueError:
        raise unsupported('with nonrectangular print names') from None
    if any(area.sheet and area.sheet.casefold() != sheet.name.casefold() for area in areas):
        raise unsupported('with print names on another worksheet')
    return areas


def _name(sheet: Worksheet, bare: str, areas: list[Area]) -> None:
    from pyopenvba.apps.excel._model import NameEntry

    if bare == 'Print_Titles':
        areas = sorted(areas, key=lambda area: area.whole_rows)
    entry = _entry(sheet, bare)
    refers = '=' + ','.join(quote_sheet(sheet.name) + '!' + area.address() for area in areas)
    if entry is None and not areas:
        return
    original_attrs = entry.attributes if entry is not None else ''
    parsed = attribute_values(original_attrs)
    if entry is not None and areas and entry.refers_to == refers and entry.visible and parsed.get('name') == '_xlnm.' + bare:
        sheet.page_print_names_owned.add(bare)
        return
    attrs = patch_attribute(original_attrs, 'name', '_xlnm.' + bare)
    sheet.page_print_names_owned.add(bare)
    if not areas:
        if entry is not None:
            sheet.book.names_.entries.remove(entry)
    elif entry is None:
        sheet.book.names_.entries.append(NameEntry(quote_sheet(sheet.name) + '!' + bare, refers, sheet.book,
                                                   attributes=attrs))
    elif entry.refers_to == refers and entry.visible and entry.attributes == attrs:
        return
    else:
        entry.refers_to, entry.visible, entry.attributes = refers, True, attrs
    sheet.book.names_.changed = True
    sheet.book.saved = False


def _finite(value: object) -> float:
    if value is EMPTY:
        return 0.0
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise unsupported('with unmeasured numeric coercions')
    result = float(value)
    if not math.isfinite(result):
        raise unsupported('with nonfinite values')
    return result


class PageSetup(ExcelObject):
    vba_type_name = 'PageSetup'

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def _check_alive(self) -> None:
        super()._check_alive()
        if self.sheet not in self.sheet.book.sheets_:
            raise unsupported('on a removed worksheet')

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, 'Long')

    def _get(self, name: str) -> object:
        self._check_alive()
        current = state(self.sheet)
        setup = current.attributes.get(('pageSetup',), {})
        if name in MARGINS:
            key = MARGINS[name]
            return float(current.attributes.get(('pageMargins',), {}).get(key, DEFAULT_MARGINS[key])) * 72
        if name in OPTIONS:
            return current.attributes.get(('printOptions',), {}).get(OPTIONS[name], '0') in ('1', 'true')
        if name in FLAGS:
            return setup.get(FLAGS[name], '0') in ('1', 'true')
        if name in ENUMS:
            attribute, mapping, default = ENUMS[name]
            raw = setup.get(attribute)
            if name == 'Orientation' and raw is None and ('pageSetup',) in current.attributes:
                return 0.0
            if raw is not None and raw not in mapping.values():
                raise unsupported('with unrecognized stored enumeration values')
            value = next((number for number, text in mapping.items() if text == raw), default)
            return float(value) if name in ('Orientation', 'Order') else VBAInt(value, 'Long')
        if name == 'PaperSize':
            return float(setup.get('paperSize', '1'))
        if name == 'FirstPageNumber':
            return float(setup.get('firstPageNumber', '1')) if setup.get('useFirstPageNumber') in ('1', 'true') else VBAInt(-4105, 'Long')
        if name == 'Zoom':
            if current.attributes.get(('sheetPr', 'pageSetUpPr'), {}).get('fitToPage') in ('1', 'true'):
                return False
            return VBAInt(int(setup.get('scale', '100')), 'Long')
        if name in ('FitToPagesWide', 'FitToPagesTall'):
            value = int(setup.get('fitToWidth' if name.endswith('Wide') else 'fitToHeight', '1'))
            return VBAInt(value, 'Long') if value else False
        if name == 'PrintArea':
            return ','.join(area.address() for area in _print_ranges(self.sheet, 'Print_Area'))
        areas = _print_ranges(self.sheet, 'Print_Titles')
        return ','.join(area.address() for area in areas if area.whole_rows == (name == 'PrintTitleRows'))

    def _set(self, name: str, value: object) -> None:
        self._check_alive()
        udf = bool(self.sheet.book.application.udf_callers)
        if udf and name not in ('LeftMargin', 'Orientation', 'PrintArea'):
            raise unsupported('setting this property from a worksheet function')
        updates: dict[tuple[str, ...], dict[str, str | None]] = {}
        if name.startswith('Print') and name in ('PrintArea', 'PrintTitleRows', 'PrintTitleColumns'):
            if not isinstance(value, str) or len(value) > 8192:
                raise unsupported('with unmeasured print-reference arguments')
            try:
                areas = parse_reference(value) if value else []
            except ValueError:
                raise unsupported('with nonrectangular print references') from None
            if any(area.sheet and area.sheet.casefold() != self.sheet.name.casefold() for area in areas):
                raise unsupported('with print references on another worksheet')
            if name == 'PrintArea':
                if any(area.whole_rows or area.whole_columns for area in areas):
                    raise unsupported('with whole-axis print areas')
                _name(self.sheet, 'Print_Area', areas)
            else:
                rows = name == 'PrintTitleRows'
                if len(areas) > 1 or any(not (area.whole_rows if rows else area.whole_columns)
                                         or (area.whole_rows and area.whole_columns) for area in areas):
                    raise unsupported('with non-axis print titles')
                retained = [area for area in _print_ranges(self.sheet, 'Print_Titles')
                            if area.whole_rows != rows]
                _name(self.sheet, 'Print_Titles', [*retained, *areas])
            return
        if name in MARGINS:
            number = _finite(value)
            if number < 0:
                raise unsupported('with negative margins (native files can be unreadable)')
            updates[('pageMargins',)] = {MARGINS[name]: repr(number / 72)}
        elif name in OPTIONS or name in FLAGS:
            if not isinstance(value, bool):
                raise unsupported('with unmeasured Boolean coercions')
            key = OPTIONS[name] if name in OPTIONS else FLAGS[name]
            updates[('printOptions',) if name in OPTIONS else ('pageSetup',)] = {key: str(int(value))}
        elif name in ENUMS:
            number = _finite(value)
            attribute, mapping, _default = ENUMS[name]
            if number not in mapping:
                raise error(1004)
            updates[('pageSetup',)] = {attribute: mapping[int(number)]}
        elif name == 'PaperSize':
            number = _finite(value)
            if number == 0:
                raise error(1004)
            if number not in (1, 9):
                raise unsupported('with unmeasured paper sizes')
            updates[('pageSetup',)] = {'paperSize': str(int(number))}
        elif name == 'FirstPageNumber':
            number = _finite(value)
            if number == -4105:
                updates[('pageSetup',)] = {'useFirstPageNumber': None}
            elif number.is_integer() and 0 <= number <= 32767:
                updates[('pageSetup',)] = {'firstPageNumber': str(int(number)), 'useFirstPageNumber': '1'}
            else:
                raise unsupported('with unmeasured page numbers')
        elif name == 'Zoom':
            if value is False:
                updates[('sheetPr', 'pageSetUpPr')] = {'fitToPage': '1'}
            else:
                if value is NULL or isinstance(value, (str, bool)):
                    raise error(1004)
                number = _finite(value)
                if not number.is_integer():
                    raise unsupported('with fractional scale')
                if not 10 <= number <= 400:
                    raise error(1004)
                updates[('pageSetup',)] = {'scale': str(int(number))}
                updates[('sheetPr', 'pageSetUpPr')] = {'fitToPage': None}
        else:
            if value is True or value is EMPTY or isinstance(value, str):
                raise error(1004)
            number = 0.0 if value is False else _finite(value)
            if number < 0:
                raise error(1004)
            if not number.is_integer() or number > 32767:
                raise unsupported('with unmeasured fit counts')
            updates[('pageSetup',)] = {'fitToWidth' if name.endswith('Wide') else 'fitToHeight': str(int(number))}
        if not udf:
            _change(self.sheet, updates)

    @member
    def LeftMargin(self) -> object:
        return self._get('LeftMargin')

    @setter('LeftMargin')
    def set_leftmargin(self, value: object) -> None:
        self._set('LeftMargin', value)

    @member
    def RightMargin(self) -> object:
        return self._get('RightMargin')

    @setter('RightMargin')
    def set_rightmargin(self, value: object) -> None:
        self._set('RightMargin', value)

    @member
    def TopMargin(self) -> object:
        return self._get('TopMargin')

    @setter('TopMargin')
    def set_topmargin(self, value: object) -> None:
        self._set('TopMargin', value)

    @member
    def BottomMargin(self) -> object:
        return self._get('BottomMargin')

    @setter('BottomMargin')
    def set_bottommargin(self, value: object) -> None:
        self._set('BottomMargin', value)

    @member
    def HeaderMargin(self) -> object:
        return self._get('HeaderMargin')

    @setter('HeaderMargin')
    def set_headermargin(self, value: object) -> None:
        self._set('HeaderMargin', value)

    @member
    def FooterMargin(self) -> object:
        return self._get('FooterMargin')

    @setter('FooterMargin')
    def set_footermargin(self, value: object) -> None:
        self._set('FooterMargin', value)

    @member
    def Orientation(self) -> object:
        return self._get('Orientation')

    @setter('Orientation')
    def set_orientation(self, value: object) -> None:
        self._set('Orientation', value)

    @member
    def PaperSize(self) -> object:
        return self._get('PaperSize')

    @setter('PaperSize')
    def set_papersize(self, value: object) -> None:
        self._set('PaperSize', value)

    @member
    def Order(self) -> object:
        return self._get('Order')

    @setter('Order')
    def set_order(self, value: object) -> None:
        self._set('Order', value)

    @member
    def FirstPageNumber(self) -> object:
        return self._get('FirstPageNumber')

    @setter('FirstPageNumber')
    def set_firstpagenumber(self, value: object) -> None:
        self._set('FirstPageNumber', value)

    @member
    def BlackAndWhite(self) -> object:
        return self._get('BlackAndWhite')

    @setter('BlackAndWhite')
    def set_blackandwhite(self, value: object) -> None:
        self._set('BlackAndWhite', value)

    @member
    def Draft(self) -> object:
        return self._get('Draft')

    @setter('Draft')
    def set_draft(self, value: object) -> None:
        self._set('Draft', value)

    @member
    def PrintComments(self) -> object:
        return self._get('PrintComments')

    @setter('PrintComments')
    def set_printcomments(self, value: object) -> None:
        self._set('PrintComments', value)

    @member
    def PrintErrors(self) -> object:
        return self._get('PrintErrors')

    @setter('PrintErrors')
    def set_printerrors(self, value: object) -> None:
        self._set('PrintErrors', value)

    @member
    def Zoom(self) -> object:
        return self._get('Zoom')

    @setter('Zoom')
    def set_zoom(self, value: object) -> None:
        self._set('Zoom', value)

    @member
    def FitToPagesWide(self) -> object:
        return self._get('FitToPagesWide')

    @setter('FitToPagesWide')
    def set_fittopageswide(self, value: object) -> None:
        self._set('FitToPagesWide', value)

    @member
    def FitToPagesTall(self) -> object:
        return self._get('FitToPagesTall')

    @setter('FitToPagesTall')
    def set_fittopagestall(self, value: object) -> None:
        self._set('FitToPagesTall', value)

    @member
    def CenterHorizontally(self) -> object:
        return self._get('CenterHorizontally')

    @setter('CenterHorizontally')
    def set_centerhorizontally(self, value: object) -> None:
        self._set('CenterHorizontally', value)

    @member
    def CenterVertically(self) -> object:
        return self._get('CenterVertically')

    @setter('CenterVertically')
    def set_centervertically(self, value: object) -> None:
        self._set('CenterVertically', value)

    @member
    def PrintGridlines(self) -> object:
        return self._get('PrintGridlines')

    @setter('PrintGridlines')
    def set_printgridlines(self, value: object) -> None:
        self._set('PrintGridlines', value)

    @member
    def PrintHeadings(self) -> object:
        return self._get('PrintHeadings')

    @setter('PrintHeadings')
    def set_printheadings(self, value: object) -> None:
        self._set('PrintHeadings', value)

    @member
    def PrintArea(self) -> object:
        return self._get('PrintArea')

    @setter('PrintArea')
    def set_printarea(self, value: object) -> None:
        self._set('PrintArea', value)

    @member
    def PrintTitleRows(self) -> object:
        return self._get('PrintTitleRows')

    @setter('PrintTitleRows')
    def set_printtitlerows(self, value: object) -> None:
        self._set('PrintTitleRows', value)

    @member
    def PrintTitleColumns(self) -> object:
        return self._get('PrintTitleColumns')

    @setter('PrintTitleColumns')
    def set_printtitlecolumns(self, value: object) -> None:
        self._set('PrintTitleColumns', value)


def guard_copy(sheet: Worksheet) -> None:
    """Refuse configured copies until printer/name-copy persistence is measured."""
    if _entry(sheet, 'Print_Area') is not None or _entry(sheet, 'Print_Titles') is not None:
        raise unsupported('copying a worksheet with print names')
    page = PageSetup(sheet)
    defaults: dict[str, object] = {name: value * 72 for name, value in
                                   ((name, DEFAULT_MARGINS[key]) for name, key in MARGINS.items())}
    defaults.update(dict.fromkeys((*OPTIONS, *FLAGS), False))
    defaults.update({'Orientation': 1.0, 'PaperSize': 1.0, 'Order': 1.0, 'FirstPageNumber': -4105,
                     'PrintComments': -4142, 'PrintErrors': 0, 'Zoom': 100,
                     'FitToPagesWide': 1, 'FitToPagesTall': 1})
    for name, wanted in defaults.items():
        value = page.vba_get(name)
        if isinstance(value, float) and isinstance(wanted, (float, int)):
            same = math.isfinite(value) and math.isclose(value, wanted, rel_tol=0, abs_tol=1e-9)
        else:
            same = value == wanted
        if not same:
            raise unsupported('copying a worksheet with nondefault configuration')
