"""Range.SparklineGroups: the sparkline groups of a sheet, which a save keeps in the worksheet's Excel 2010 extension.

Measured in live Excel (scripts/measure_sparklines.py, tests/fixtures/sparklines.json):

- Add(Type, SourceData) puts a group of line, column or win/loss sparklines in the range, one for each of its
  cells, which run down a column or across a row; with one cell its source is one row or column, with more each
  takes a row of the source where the source has that many rows, else a column. A type Excel does not know, a
  block of cells, a source that does not fit and one that is no reference are error 1004. A cell holding a
  sparkline already takes the new one and loses the old.
- A source reads as the reference given, in capitals, a cell as a range of one (A1 reads A1:A1), another sheet
  named before it; a group's source is the block its sparklines' sources make. The groups a range reaches read
  newest first.
- A cell's value and a sparkline in it live apart. Range.Clear takes the range's sparklines away, a group left
  with none going too; ClearContents leaves them.
- A new group's line is 0.75pt, in Excel's series colour and red for the rest, an empty cell shown as a gap. A
  save writes each group with its attributes in the schema's order, the colours, and each sparkline's source with
  its sheet's name and its cell.

The model moves no sparkline: a row or column put in or taken away, a copy, a cut, a sort or a fill on a sheet a
sparkline sits on or reads, and that sheet copied, renamed or deleted, each report themselves unsupported.
"""
from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba._opc_xml import opc_part
from pyopenvba._xml import escape, escape_text
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook, Worksheet

URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
X14_URI = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM_URI = 'http://schemas.microsoft.com/office/excel/2006/main'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
#: XlSparkType by the type attribute a file gives it; a line has none.
TYPES = {1: None, 2: 'column', 3: 'stacked'}
#: The group's attributes in the schema's order, as Excel writes them, its revision id last.
ORDER = ('manualMax', 'manualMin', 'lineWeight', 'type', 'dateAxis', 'displayEmptyCellsAs', 'markers', 'high',
         'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'minAxisType', 'maxAxisType',
         'rightToLeft', 'xr2:uid')
#: A new group's colours, by element, as Excel writes them.
COLOURS = (('colorSeries', 'FF376092'), ('colorNegative', 'FFD00000'), ('colorAxis', 'FF000000'),
           ('colorMarkers', 'FFD00000'), ('colorFirst', 'FFD00000'), ('colorLast', 'FFD00000'),
           ('colorHigh', 'FFD00000'), ('colorLow', 'FFD00000'))
#: The exts an extLst puts ahead of sparklines: conditional formats and data validations.
_BEFORE = ('{78C0D931-6437-407d-A8EE-F0AAD7539E65}', '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}')
_REFERENCE = re.compile(r"(?:(?P<sheet>'(?:[^']|'')+'|[^!'\s]+)!)?(?P<c1>\$?[A-Z]{1,3})(?P<r1>\$?\d+)"
                        r"(?::(?P<c2>\$?[A-Z]{1,3})(?P<r2>\$?\d+))?", re.IGNORECASE)


@dataclass
class Source:
    """A sparkline's source: its sheet's name, None for the group's own, and its corners as written, $ and all."""

    sheet: str | None
    first: tuple[str, str]
    last: tuple[str, str]

    def area(self) -> Area:
        from pyopenvba._a1 import parse_area

        return parse_area(f'{"".join(self.first)}:{"".join(self.last)}'.replace('$', ''))

    def text(self, own: str) -> str:
        """As SourceData reads it, the sheet named where it is not the group's own."""
        reference = f'{"".join(self.first)}:{"".join(self.last)}'
        return reference if self.sheet is None or self.sheet == own else f'{_quoted(self.sheet)}!{reference}'

    def formula(self, own: str) -> str:
        """As a file keeps it, always with its sheet's name."""
        reference = f'{"".join(self.first)}:{"".join(self.last)}'
        return f'{_quoted(self.sheet or own)}!{reference}'


def _quoted(name: str) -> str:
    return name if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*', name) else "'" + name.replace("'", "''") + "'"


@dataclass
class Sparkline:
    source: Source
    #: Its cell, row and column.
    at: tuple[int, int]


@dataclass
class Group:
    #: Its attributes in the schema's order, as a file keeps them.
    attributes: dict[str, str]
    #: Its colours and anything else before its sparklines, as markup in its order.
    children: list[str]
    sparklines: list[Sparkline] = field(default_factory=list[Sparkline])


@dataclass
class State:
    groups: list[Group]
    changed: bool = False


def _book_sheet(book: Workbook, name: str) -> Worksheet | None:
    return next((sheet for sheet in book.sheets_ if sheet.name.casefold() == name.casefold()), None)


def state(sheet: Worksheet) -> State:
    """The sheet's groups, read from its part the first time they are reached."""
    if sheet.sparklines is None:
        sheet.sparklines = State(_loaded(sheet))
    return sheet.sparklines


def _loaded(sheet: Worksheet) -> list[Group]:
    from pyopenvba.apps.excel._charts_file import xml_tree

    package = sheet.book.package
    if package is None or not sheet.part_name or not package.has(sheet.part_name):
        return []
    data = package.read(sheet.part_name)
    if URI.encode() not in data:
        return []
    root = xml_tree(data)
    x14, xm = '{' + X14_URI + '}', '{' + XM_URI + '}'
    groups: list[Group] = []
    for ext in root.iter('{' + MAIN + '}ext'):
        if ext.get('uri') != URI:
            continue
        for found in ext.iter(x14 + 'sparklineGroup'):
            children: list[str] = []
            group = Group({_prefixed(key): value for key, value in found.attrib.items()}, children)
            for child in found:
                if child.tag != x14 + 'sparklines':
                    children.append(_markup(child))
                    continue
                for line in child:
                    formula, where = line.find(xm + 'f'), line.find(xm + 'sqref')
                    if formula is None or where is None or not (formula.text or '') or ' ' in (where.text or ''):
                        raise VBAUnsupportedError('a sparkline with no source or several cells is not implemented')
                    source = _source(formula.text or '', sheet.name)
                    if source is None:
                        raise VBAUnsupportedError('a sparkline whose source is no range is not implemented')
                    from pyopenvba._a1 import parse_area

                    cell = parse_area(where.text or '')
                    group.sparklines.append(Sparkline(source, (cell.top, cell.left)))
            groups.append(group)
    # A file holds the newest group first; the model keeps them oldest first.
    return groups[::-1]


def _prefixed(key: str) -> str:
    return 'xr2:' + key.split('}', 1)[1] if key.startswith('{http://schemas.microsoft.com/office/spreadsheetml/2015') else key


def _markup(element: XmlElement) -> str:
    local = element.tag.split('}', 1)[1]
    attributes = ''.join(f' {key}="{escape(value)}"' for key, value in element.attrib.items())
    if len(element) or element.text:
        raise VBAUnsupportedError(f'a sparkline group with x14:{local} holding more than attributes is not implemented')
    return f'<x14:{local}{attributes}/>'


def _source(text: str, own: str) -> Source | None:
    found = _REFERENCE.fullmatch(text.strip())
    if found is None:
        return None
    sheet = found['sheet']
    if sheet is not None and sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    first = (found['c1'].upper(), found['r1'])
    last = (found['c2'].upper(), found['r2']) if found['c2'] else first
    return Source(None if sheet is None or sheet == own else sheet, first, last)


def touches(sheet: Worksheet) -> bool:
    """Whether a sparkline sits on the sheet or reads from it."""
    for other in sheet.book.sheets_:
        groups = state(other).groups
        if other is sheet and groups:
            return True
        if any(line.source.sheet is not None and line.source.sheet.casefold() == sheet.name.casefold()
               for group in groups for line in group.sparklines):
            return True
    return False


def refuse(sheet: Worksheet, what: str) -> None:
    """The model moves no sparkline (see the module's notes)."""
    if touches(sheet):
        raise VBAUnsupportedError(f'{what} a sheet a sparkline sits on or reads is not implemented')


def cleared(target: Range) -> None:
    """Range.Clear: the range's sparklines go, and a group left with none."""
    found = state(target.sheet)
    if not found.groups:
        return
    kept: list[Group] = []
    for group in found.groups:
        left = [line for line in group.sparklines if not _inside(target, line.at)]
        if len(left) != len(group.sparklines):
            found.changed = True
        group.sparklines = left
        if left:
            kept.append(group)
    found.groups = kept


def _inside(target: Range, at: tuple[int, int]) -> bool:
    return any(area.contains(*at) for area in target.areas)


def _pairs(sheet: Worksheet, cells: list[tuple[int, int]], given: Source) -> list[Sparkline]:
    """The sparklines a source makes for these cells, one for each, a row of the source each where it has as
    many rows, else a column; error 1004 where it fits neither."""
    area = given.area()
    rows = [(row, row, area.left, area.right) for row in range(area.top, area.bottom + 1)]
    columns = [(area.top, area.bottom, column, column) for column in range(area.left, area.right + 1)]
    if len(cells) == 1:
        if area.rows != 1 and area.columns != 1:
            raise error(1004)
        return [Sparkline(given, cells[0])]
    chosen = rows if len(rows) == len(cells) else columns if len(columns) == len(cells) else None
    if chosen is None:
        raise error(1004)
    made: list[Sparkline] = []
    for (top, bottom, left, right), cell in zip(chosen, cells, strict=True):
        made.append(Sparkline(Source(given.sheet, _corner(given.first, top, left), _corner(given.last, bottom, right)),
                              cell))
    return made


def _corner(written: tuple[str, str], row: int, column: int) -> tuple[str, str]:
    from pyopenvba._a1 import column_letter

    if written[0].startswith('$') or written[1].startswith('$'):
        raise VBAUnsupportedError('splitting an absolute source among several sparklines is not implemented')
    return column_letter(column), str(row)


def _given_source(sheet: Worksheet, value: object) -> Source:
    text = to_text(value)
    found = _source(text, sheet.name)
    if found is None:
        raise error(1004)
    if found.sheet is not None and _book_sheet(sheet.book, found.sheet) is None:
        raise error(1004)
    return found


def _cells(target: Range) -> list[tuple[int, int]]:
    if len(target.areas) != 1:
        raise VBAUnsupportedError('sparklines in several areas are not implemented')
    area = target.first
    if area.rows != 1 and area.columns != 1:
        raise error(1004)
    return [(row, column) for row in range(area.top, area.bottom + 1) for column in range(area.left, area.right + 1)]


def _new_group(kind: int, lines: list[Sparkline]) -> Group:
    attributes: dict[str, str] = {}
    if TYPES[kind] is not None:
        attributes['type'] = str(TYPES[kind])
    attributes['displayEmptyCellsAs'] = 'gap'
    if kind == 3:
        attributes['negative'] = '1'
    attributes['xr2:uid'] = '{' + str(uuid.uuid4()).upper() + '}'
    return Group(attributes, [f'<x14:{name} rgb="{rgb}"/>' for name, rgb in COLOURS], lines)


def _kind(value: object) -> int:
    found = int(to_integer(value, 'Long'))
    if found not in TYPES:
        raise error(1004)
    return found


def _ordered(attributes: dict[str, str]) -> dict[str, str]:
    rank = {name: place for place, name in enumerate(ORDER)}
    return dict(sorted(attributes.items(), key=lambda item: rank.get(item[0], len(ORDER))))


class SparklineGroups(VBACollection, ExcelObject):
    """The sparkline groups a range reaches, newest first."""

    vba_type_name = 'SparklineGroups'

    def __init__(self, target: Range) -> None:
        self.target = target
        self.sheet = target.sheet

    def _reached(self) -> list[Group]:
        return [group for group in reversed(state(self.sheet).groups)
                if any(_inside(self.target, line.at) for line in group.sparklines)]

    def vba_items(self) -> list[object]:
        return [SparklineGroup(self.sheet, group, self) for group in self._reached()]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(1004)
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @method
    def Add(self, Type: object = MISSING, SourceData: object = MISSING) -> object:
        if Type is MISSING or SourceData is MISSING:
            raise error(449)
        kind = _kind(Type)
        cells = _cells(self.target)
        lines = _pairs(self.sheet, cells, _given_source(self.sheet, SourceData))
        found = state(self.sheet)
        # A cell holding a sparkline already takes the new one.
        taken = {line.at for line in lines}
        for group in found.groups:
            group.sparklines = [line for line in group.sparklines if line.at not in taken]
        found.groups = [group for group in found.groups if group.sparklines]
        group = _new_group(kind, lines)
        found.groups.append(group)
        found.changed = True
        self.sheet.book.saved = False
        return SparklineGroup(self.sheet, group, self)

    @method
    def Clear(self) -> object:
        cleared(self.target)
        self.sheet.book.saved = False
        return EMPTY


class SparklineGroup(ExcelObject):
    vba_type_name = 'SparklineGroup'

    def __init__(self, sheet: Worksheet, group: Group, parent: SparklineGroups) -> None:
        self.sheet, self.group, self.parent = sheet, group, parent

    def _check_alive(self) -> None:
        super()._check_alive()
        if self.group not in state(self.sheet).groups:
            raise VBAUnsupportedError('a sparkline group since deleted is not implemented')

    def edited(self) -> None:
        state(self.sheet).changed = True
        self.sheet.book.saved = False

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Count(self) -> object:
        return VBAInt(len(self.group.sparklines), 'Long')

    @member(default=True)
    def Item(self, Index: object = MISSING) -> object:
        place = int(to_integer(Index, 'Long'))
        if not 1 <= place <= len(self.group.sparklines):
            raise error(1004)
        return SparklineObject(self, self.group.sparklines[place - 1])

    @member
    def Type(self) -> object:
        written = self.group.attributes.get('type')
        return VBAInt(next(kind for kind, name in TYPES.items() if name == written), 'Long')

    @setter('Type')
    def _set_type(self, value: object) -> None:
        kind = _kind(value)
        if kind == 3 or 'negative' in self.group.attributes:
            raise VBAUnsupportedError('changing a sparkline group to or from win/loss is not implemented')
        if TYPES[kind] is None:
            self.group.attributes.pop('type', None)
        else:
            self.group.attributes['type'] = str(TYPES[kind])
        self.group.attributes = _ordered(self.group.attributes)
        self.edited()

    @member
    def Location(self) -> object:
        cells = [line.at for line in self.group.sparklines]
        from pyopenvba.apps.excel._model import coalesced

        return Range(self.sheet, coalesced([Area(row, column, row, column, self.sheet.name) for row, column in cells]))

    @setter('Location')
    def _set_location(self, value: object) -> None:
        """Passed over, as Excel passes it over: ModifyLocation moves a group."""

    @member
    def SourceData(self) -> object:
        sources = [line.source for line in self.group.sparklines]
        sheets = {source.sheet for source in sources}
        if len(sheets) != 1 or any('$' in ''.join(source.first + source.last) for source in sources[1:]):
            raise VBAUnsupportedError('the source of a group whose sparklines read apart is not implemented')
        if len(sources) == 1:
            return sources[0].text(self.sheet.name)
        areas = [source.area() for source in sources]
        top, bottom = min(area.top for area in areas), max(area.bottom for area in areas)
        left, right = min(area.left for area in areas), max(area.right for area in areas)
        if sum(area.rows * area.columns for area in areas) != (bottom - top + 1) * (right - left + 1):
            raise VBAUnsupportedError('the source of a group whose sparklines read apart is not implemented')
        from pyopenvba._a1 import column_letter

        block = Source(sources[0].sheet, (column_letter(left), str(top)), (column_letter(right), str(bottom)))
        return block.text(self.sheet.name)

    @setter('SourceData')
    def _set_source_data(self, value: object) -> None:
        self.ModifySourceData(value)

    @method
    def ModifySourceData(self, SourceData: object = MISSING) -> object:
        cells = [line.at for line in self.group.sparklines]
        self.group.sparklines = _pairs(self.sheet, cells, _given_source(self.sheet, SourceData))
        self.edited()
        return EMPTY

    @method
    def ModifyLocation(self, Location: object = MISSING) -> object:
        if not isinstance(Location, Range) or Location.sheet is not self.sheet:
            raise VBAUnsupportedError('moving a sparkline group elsewhere than its sheet is not implemented')
        cells = _cells(Location)
        if len(cells) != len(self.group.sparklines):
            raise error(1004)
        self.refuse_taken(set(cells))
        for line, cell in zip(self.group.sparklines, cells, strict=True):
            line.at = cell
        self.edited()
        return EMPTY

    def refuse_taken(self, cells: set[tuple[int, int]]) -> None:
        found = state(self.sheet)
        if any(line.at in cells for group in found.groups if group is not self.group for line in group.sparklines):
            raise VBAUnsupportedError('moving sparklines onto cells holding others is not implemented')

    @method
    def Modify(self, Location: object = MISSING, SourceData: object = MISSING) -> object:
        if not isinstance(Location, Range) or Location.sheet is not self.sheet:
            raise VBAUnsupportedError('moving a sparkline group elsewhere than its sheet is not implemented')
        cells = _cells(Location)
        lines = _pairs(self.sheet, cells, _given_source(self.sheet, SourceData))
        self.refuse_taken({line.at for line in lines})
        self.group.sparklines = lines
        self.edited()
        return EMPTY

    @method
    def Delete(self) -> object:
        found = state(self.sheet)
        found.groups = [group for group in found.groups if group is not self.group]
        found.changed = True
        self.sheet.book.saved = False
        return EMPTY

    @member
    def LineWeight(self) -> object:
        return float(self.group.attributes.get('lineWeight', '0.75'))

    @setter('LineWeight')
    def _set_line_weight(self, value: object) -> None:
        weight = float(to_number(value))
        if weight < 0 or weight > 1584:
            raise error(1004)
        if weight == 0.75:
            self.group.attributes.pop('lineWeight', None)
        else:
            self.group.attributes['lineWeight'] = f'{weight:g}'
        self.group.attributes = _ordered(self.group.attributes)
        self.edited()

    @member
    def DisplayBlanksAs(self) -> object:
        """xlNotPlotted, gaps for empty cells."""
        if self.group.attributes.get('displayEmptyCellsAs', 'zero') != 'gap':
            raise VBAUnsupportedError('a sparkline group showing empty cells other than as gaps is not implemented')
        return VBAInt(1, 'Long')

    @member
    def DisplayHidden(self) -> object:
        return self.group.attributes.get('displayHidden', '0') in ('1', 'true')

    @member
    def SeriesColor(self) -> object:
        return SparkColor(self, 'colorSeries')


class SparklineObject(ExcelObject):
    vba_type_name = 'Sparkline'

    def __init__(self, parent: SparklineGroup, line: Sparkline) -> None:
        self.parent, self.line = parent, line

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Location(self) -> object:
        row, column = self.line.at
        sheet = self.parent.sheet
        return Range(sheet, [Area(row, column, row, column, sheet.name)])

    @member
    def SourceData(self) -> object:
        return self.line.source.text(self.parent.sheet.name)

    @setter('SourceData')
    def _set_source_data(self, value: object) -> None:
        self.ModifySourceData(value)

    @method
    def ModifySourceData(self, Formula: object = MISSING) -> object:
        sheet = self.parent.sheet
        self.line.source = _pairs(sheet, [self.line.at], _given_source(sheet, Formula))[0].source
        self.parent.edited()
        return EMPTY

    @method
    def ModifyLocation(self, Range: object = MISSING) -> object:
        target = Range
        from pyopenvba.apps.excel._model import Range as RangeType

        if not isinstance(target, RangeType) or target.sheet is not self.parent.sheet:
            raise VBAUnsupportedError('moving a sparkline elsewhere than its sheet is not implemented')
        cells = _cells(target)
        if len(cells) != 1:
            raise error(1004)
        self.parent.refuse_taken(set(cells))
        self.line.at = cells[0]
        self.parent.edited()
        return EMPTY


class SparkColor(ExcelObject):
    """A group's colour: its RGB, or a theme colour's."""

    vba_type_name = 'SparkColor'

    def __init__(self, parent: SparklineGroup, element: str) -> None:
        self.parent, self.element = parent, element

    def _attributes(self) -> dict[str, str]:
        for child in self.parent.group.children:
            if child.startswith(f'<x14:{self.element} ') or child == f'<x14:{self.element}/>':
                return dict(re.findall(r'([\w:]+)="([^"]*)"', child))
        raise VBAUnsupportedError(f'a sparkline group with no {self.element} is not implemented')

    def _write(self, attributes: dict[str, str]) -> None:
        markup = f'<x14:{self.element}' + ''.join(f' {key}="{escape(value)}"' for key, value in attributes.items()) + '/>'
        children = self.parent.group.children
        for place, child in enumerate(children):
            if child.startswith(f'<x14:{self.element} ') or child == f'<x14:{self.element}/>':
                children[place] = markup
        self.parent.edited()

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Color(self) -> object:
        from pyopenvba.apps.excel._styles import Color, bgr

        attributes = self._attributes()
        if 'rgb' in attributes:
            return float(bgr(attributes['rgb'][-6:]))
        if 'theme' in attributes and 'tint' not in attributes:
            colors = self.parent.sheet.book.stylesheet.colors
            return float(bgr(colors.rrggbb(Color('theme', int(attributes['theme'])), '000000')))
        raise VBAUnsupportedError('a sparkline colour of this form is not implemented')

    @setter('Color')
    def _set_color(self, value: object) -> None:
        from pyopenvba.apps.excel._styles import Color

        self._write({'rgb': str(Color.from_bgr(int(to_integer(value, 'Long'))).value)})

    @member
    def ThemeColor(self) -> object:
        attributes = self._attributes()
        if 'theme' in attributes:
            return VBAInt(int(attributes['theme']) + 1, 'Long')
        return VBAInt(0, 'Long')

    @member
    def TintAndShade(self) -> object:
        return float(self._attributes().get('tint', '0'))


def written(sheet: Worksheet, xml: str) -> str:
    """The sheet's part with its sparkline groups as Excel writes them, where a macro changed them."""
    if sheet.sparklines is None or not sheet.sparklines.changed:
        return xml
    groups = sheet.sparklines.groups
    markup = ''
    if groups:
        if groups and 'xmlns:xr2=' not in xml[:xml.find('>', xml.find('<worksheet'))]:
            for group in groups:
                group.attributes.pop('xr2:uid', None)
        # Newest first, as Excel writes them (tests/fixtures/sparklines.json two).
        body = ''.join(_group_markup(sheet, group) for group in reversed(groups))
        markup = (f'<ext uri="{URI}" xmlns:x14="{X14_URI}"><x14:sparklineGroups xmlns:xm="{XM_URI}">{body}'
                  '</x14:sparklineGroups></ext>')
    sheet.sparklines.changed = False
    part = opc_part(xml.encode('utf-8'), MAIN, 'worksheet')
    data = part.data
    lists = [(start, end) for tag, _attrs, start, end in part.children if tag == MAIN + '}extLst']
    if not lists:
        if not markup:
            return xml
        return part.insert(part.closing_start, f'<extLst>{markup}</extLst>'.encode()).decode('utf-8')
    start, end = lists[0]
    inner = opc_part(data, MAIN, 'worksheet', container_path=('extLst',))
    assert inner is not None
    exts = [(attributes.get('uri', ''), data[first:last]) for _tag, attributes, first, last in inner.children]
    kept = [text for uri, text in exts if uri != URI]
    place = next((index for index, (uri, _) in enumerate([pair for pair in exts if pair[0] != URI])
                  if uri not in _BEFORE), len(kept))
    if markup:
        kept.insert(place, markup.encode('utf-8'))
    replacement = b'<extLst>' + b''.join(kept) + b'</extLst>' if kept else b''
    return (data[:start] + replacement + data[end:]).decode('utf-8')


def _group_markup(sheet: Worksheet, group: Group) -> str:
    attributes = ''.join(f' {key}="{escape(value)}"' for key, value in _ordered(group.attributes).items())
    lines = ''.join(f'<x14:sparkline><xm:f>{escape_text(line.source.formula(sheet.name))}</xm:f>'
                    f'<xm:sqref>{_cell_name(line.at)}</xm:sqref></x14:sparkline>' for line in group.sparklines)
    return f'<x14:sparklineGroup{attributes}>{"".join(group.children)}<x14:sparklines>{lines}</x14:sparklines>' \
           '</x14:sparklineGroup>'


def _cell_name(at: tuple[int, int]) -> str:
    from pyopenvba._a1 import column_letter

    return f'{column_letter(at[1])}{at[0]}'
