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

Inserts and deletes of cells, rows and columns move sparklines as ``shifted`` has it
(scripts/measure_sparkline_shifts.py, tests/fixtures/sparkline_shifts*.json). A group whose sparklines no longer
run along a line reading the rows or columns of a block reads its Location and SourceData cell by cell.

Otherwise the model moves no sparkline: a copy, a cut, a sort or a fill on a sheet a sparkline sits on or reads,
and that sheet copied, renamed or deleted, each report themselves unsupported.
"""
from __future__ import annotations

import re
import uuid
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba._opc_xml import opc_part
from pyopenvba._xml import escape, escape_text
from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (EMPTY, MISSING, NULL, VBAInt, error, to_bool, to_integer, to_number,
                                           to_text)

if TYPE_CHECKING:
    from collections.abc import Callable

    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._editing import CellShift
    from pyopenvba.apps.excel._styles import Color
    from pyopenvba.apps.excel._model import Workbook, Worksheet

URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'
X14_URI = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM_URI = 'http://schemas.microsoft.com/office/excel/2006/main'
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
#: XlSparkType by the type attribute a file gives it; a line has none.
TYPES = {1: None, 2: 'column', 3: 'stacked'}
#: XlDisplayBlanksAs by the displayEmptyCellsAs a file gives it; zero is the schema's default.
BLANKS = {1: 'gap', 2: 'zero', 3: 'span'}
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
    #: Its cells deleted: a sparkline's SourceData reads "", its group's "#REF!", and a file writes no formula.
    broken: bool = False

    def area(self) -> Area:
        from pyopenvba._a1 import parse_area

        if self.broken:
            raise VBAUnsupportedError('a sparkline whose source was deleted is not implemented here')
        return parse_area(f'{"".join(self.first)}:{"".join(self.last)}'.replace('$', ''))

    def text(self, own: str) -> str:
        """As SourceData reads it, the sheet named where it is not the group's own."""
        if self.broken:
            return ''
        reference = f'{"".join(self.first)}:{"".join(self.last)}'
        return reference if self.sheet is None or self.sheet == own else f'{_quoted(self.sheet)}!{reference}'

    def formula(self, own: str) -> str:
        """As a file keeps it, always with its sheet's name; "" once its cells are deleted."""
        if self.broken:
            return ''
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
    #: Its date range: the sheet named, None for its own, and the reference as given, in capitals.
    dates: tuple[str | None, str] | None = None


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
                if child.tag == xm + 'f':
                    dates = _dates(child.text or '', sheet.name)
                    if dates is None:
                        raise VBAUnsupportedError('a sparkline date range that is no range is not implemented')
                    group.dates = dates
                    continue
                if child.tag != x14 + 'sparklines':
                    children.append(_markup(child))
                    continue
                for line in child:
                    formula, where = line.find(xm + 'f'), line.find(xm + 'sqref')
                    if where is None or ' ' in (where.text or '') or formula is not None and not formula.text:
                        raise VBAUnsupportedError('a sparkline with an empty source or several cells is not implemented')
                    # One whose source was deleted keeps no formula (tests/fixtures/sparkline_shifts_c.json).
                    source = (Source(None, ('', ''), ('', ''), broken=True) if formula is None
                              else _source(formula.text or '', sheet.name))
                    if source is None:
                        raise VBAUnsupportedError('a sparkline whose source is no range is not implemented')
                    from pyopenvba._a1 import parse_area

                    cell = parse_area(where.text or '')
                    group.sparklines.append(Sparkline(source, (cell.top, cell.left)))
            groups.append(group)
    # Loaded groups read last in the file first (tests/fixtures/sparkline_shifts_b.json two): the model keeps them
    # in the file's order, so a save writes them the other way round, as Excel does.
    return groups


def _prefixed(key: str) -> str:
    revision = '{http://schemas.microsoft.com/office/spreadsheetml/2015'
    return 'xr2:' + key.split('}', 1)[1] if key.startswith(revision) else key


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


def _dates(text: str, own: str) -> tuple[str | None, str] | None:
    """A date range as given: its sheet, None for the group's own, and its reference in capitals, a cell kept a
    cell (tests/fixtures/sparkline_dates.json)."""
    found = _REFERENCE.fullmatch(text.strip())
    if found is None:
        return None
    sheet = found['sheet']
    if sheet is not None and sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    reference = found['c1'].upper() + found['r1'] + (':' + found['c2'].upper() + found['r2'] if found['c2'] else '')
    return (None if sheet is None or sheet == own else sheet), reference


def touches(sheet: Worksheet) -> bool:
    """Whether a sparkline sits on the sheet or reads from it."""
    for other in sheet.book.sheets_:
        groups = state(other).groups
        if other is sheet and groups:
            return True
        if any(line.source.sheet is not None and line.source.sheet.casefold() == sheet.name.casefold()
               for group in groups for line in group.sparklines):
            return True
        if any(group.dates is not None and group.dates[0] is not None
               and group.dates[0].casefold() == sheet.name.casefold() for group in groups):
            return True
    return False


def refuse(sheet: Worksheet, what: str) -> None:
    """The model moves no sparkline (see the module's notes)."""
    if touches(sheet):
        raise VBAUnsupportedError(f'{what} a sheet a sparkline sits on or reads is not implemented')


def _regular(lines: list[Sparkline]) -> bool:
    """Whether a group reads as one location and one source: its cells run one by one along a row or a column, in
    its order, and its sources, on one sheet, are the rows or the columns of a block one by one in the same order
    (tests/fixtures/sparkline_shifts_c.json: F1, F2, G3, F4 reading A1:D4 by rows reads as lists)."""
    if len(lines) == 1:
        return not lines[0].source.broken
    if any(line.source.broken for line in lines) or len({line.source.sheet for line in lines}) != 1:
        return False
    cells = [line.at for line in lines]
    steps = {(second[0] - first[0], second[1] - first[1]) for first, second in zip(cells, cells[1:])}
    if steps not in ({(1, 0)}, {(0, 1)}):
        return False
    areas = [line.source.area() for line in lines]
    by_rows = all(area.rows == 1 and (area.left, area.right) == (areas[0].left, areas[0].right) for area in areas)
    by_columns = all(area.columns == 1 and (area.top, area.bottom) == (areas[0].top, areas[0].bottom)
                     for area in areas)
    return (by_rows and all(second.top == first.top + 1 for first, second in zip(areas, areas[1:]))
            or by_columns and all(second.left == first.left + 1 for first, second in zip(areas, areas[1:])))


def shifted(sheet: Worksheet, shift: CellShift, rewrite: Callable[[str, str], str], *, below: bool = False
            ) -> Callable[[], None]:
    """Cells, rows or columns of ``sheet`` deleted or inserted, as ``shift`` has it, ``rewrite`` taking a formula
    and the sheet it is on to the formula that follows them. Checks first and returns what carries it out
    (tests/fixtures/sparkline_shifts*.json):

    - a sparkline on ``sheet`` moves with its cell and goes with it, a group left with none going too;
    - sources and date ranges follow as formulas' references do, $ or not, one whose cells all go reading #REF!;
    - each cell put in takes the sparkline of the cell before it -- above for cells going down, to the left for
      cells going right -- with that one's source moved as far, $ and all, coming after it in its group.
    """
    found: list[tuple[Worksheet, State, list[tuple[Group, list[Sparkline], tuple[str | None, str] | None]]]] = []
    for owner in sheet.book.sheets_:
        own = state(owner)
        if not own.groups:
            continue
        groups: list[tuple[Group, list[Sparkline], tuple[str | None, str] | None]] = []
        for group in own.groups:
            lines: list[Sparkline] = []
            for line in group.sparklines:
                at = line.at
                if owner is sheet:
                    moved = _moved_cell(at, shift)
                    if moved is None:
                        continue
                    at = moved
                kept = Sparkline(_followed(line.source, owner.name, rewrite), at)
                lines.append(kept)
                if owner is sheet and not shift.delete:
                    lines.extend(_copies(line, kept, shift, below))
            dates = group.dates
            if dates is not None:
                text = rewrite('=' + _quoted(dates[0] or owner.name) + '!' + dates[1], owner.name)[1:]
                dates = None if '#REF!' in text else _dates(text, owner.name)
            groups.append((group, lines, dates))
        found.append((owner, own, groups))

    def carry_out() -> None:
        for owner, own, groups in found:
            if all(lines == group.sparklines and dates == group.dates for group, lines, dates in groups):
                continue
            for group, lines, dates in groups:
                group.sparklines, group.dates = lines, dates
            own.groups = [group for group, lines, _dates in groups if lines]
            own.changed = True
            owner.book.saved = False

    return carry_out


def _moved_cell(at: tuple[int, int], shift: CellShift) -> tuple[int, int] | None:
    from pyopenvba.apps.excel._editing import shifted_box

    box = shifted_box((at[0], at[1], at[0], at[1]), shift)
    return None if box is None else (box[0], box[1])


def _followed(source: Source, own: str, rewrite: Callable[[str, str], str]) -> Source:
    if source.broken:
        return source
    text = rewrite('=' + source.formula(own), own)[1:]
    if '#REF!' in text:
        return Source(None, ('', ''), ('', ''), broken=True)
    followed = _source(text, own)
    if followed is None:
        raise VBAUnsupportedError('a sparkline source that an edit turns into no range is not implemented')
    return followed


def _copies(line: Sparkline, kept: Sparkline, shift: CellShift, below: bool) -> list[Sparkline]:
    """The sparklines the cells put in after ``line``'s cell take from it, ``kept`` being where it went."""
    row, column = line.at
    along, across = (row, column) if shift.vertical else (column, row)
    if not shift.band[0] <= across <= shift.band[1]:
        return []
    if along != (shift.start + shift.count if below else shift.start - 1):
        return []
    if below or kept.source.broken:
        raise VBAUnsupportedError('cells put in beside a sparkline taking it from below, or one whose source was '
                                  'deleted, are not implemented')
    made: list[Sparkline] = []
    for step in range(1, shift.count + 1):
        at = (row + step, column) if shift.vertical else (row, column + step)
        made.append(Sparkline(_offset(kept.source, step if shift.vertical else 0, 0 if shift.vertical else step), at))
    return made


def _offset(source: Source, rows: int, columns: int) -> Source:
    from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_letter, column_number

    def moved(corner: tuple[str, str]) -> tuple[str, str]:
        letters, digits = corner
        column = column_number(letters.lstrip('$')) + columns
        row = int(digits.lstrip('$')) + rows
        if not (1 <= column <= MAX_COLUMNS and 1 <= row <= MAX_ROWS):
            raise VBAUnsupportedError('a sparkline copied past the edge of the sheet is not implemented')
        return (('$' if letters.startswith('$') else '') + column_letter(column),
                ('$' if digits.startswith('$') else '') + str(row))

    return Source(source.sheet, moved(source.first), moved(source.last))


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


class SparklineGroups(VBACollection, ItemDefault, ExcelObject):
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

    def _edited(self) -> None:
        state(self.sheet).changed = True
        self.sheet.book.saved = False

    @method
    def ClearGroups(self) -> object:
        """Every group the range reaches goes, its sparklines outside the range too
        (tests/fixtures/sparkline_grouping.json)."""
        reached = self._reached()
        if reached:
            found = state(self.sheet)
            found.groups = [group for group in found.groups if group not in reached]
            self._edited()
        return EMPTY

    @method
    def Group(self, Location: object = MISSING) -> object:
        """The range's sparklines in one group, with the settings of the group Location's cell holds; a Location
        outside the range or holding none is error 1004 (tests/fixtures/sparkline_grouping.json)."""
        if not isinstance(Location, Range) or Location.sheet is not self.sheet or not Location.single:
            raise error(1004)
        at = (Location.first.top, Location.first.left)
        if not _inside(self.target, at):
            raise error(1004)
        found = state(self.sheet)
        model = next((group for group in found.groups if any(line.at == at for line in group.sparklines)), None)
        if model is None:
            raise error(1004)
        reached = [group for group in found.groups if any(_inside(self.target, line.at) for line in group.sparklines)]
        if any(not _inside(self.target, line.at) for group in reached for line in group.sparklines):
            raise VBAUnsupportedError('grouping part of a sparkline group is not implemented')
        # Oldest group first, each in its own order, in the place of Location's group.
        model.sparklines = [line for group in reached for line in group.sparklines]
        found.groups = [group for group in found.groups if group is model or group not in reached]
        self._edited()
        return EMPTY

    @method
    def Ungroup(self) -> object:
        """Each of the range's sparklines in a group of its own with its group's settings, the newest last in the
        range's order, its old group keeping the rest (tests/fixtures/sparkline_grouping.json)."""
        found = state(self.sheet)
        made: list[Group] = []
        for group in list(found.groups):
            taken = [line for line in group.sparklines if _inside(self.target, line.at)]
            if not taken or len(group.sparklines) == 1:
                if taken:
                    raise VBAUnsupportedError('ungrouping a group of one sparkline is not implemented')
                continue
            group.sparklines = [line for line in group.sparklines if line not in taken]
            for line in taken:
                attributes = dict(group.attributes)
                if 'xr2:uid' in attributes:
                    attributes['xr2:uid'] = '{' + str(uuid.uuid4()).upper() + '}'
                made.append(Group(attributes, list(group.children), [line]))
        if made:
            found.groups = [group for group in found.groups if group.sparklines]
            made.sort(key=lambda group: group.sparklines[0].at)
            found.groups.extend(made)
            self._edited()
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
        if kind == 3 or self.group.attributes.get('type') == 'stacked':
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
        areas = [Area(row, column, row, column, self.sheet.name) for row, column in cells]
        if not _regular(self.group.sparklines):
            # Its cells one by one, in the group's order (tests/fixtures/sparkline_shifts_c.json).
            return Range(self.sheet, areas)
        from pyopenvba.apps.excel._model import coalesced

        return Range(self.sheet, coalesced(areas))

    @setter('Location')
    def _set_location(self, value: object) -> None:
        """Passed over, as Excel passes it over: ModifyLocation moves a group."""

    @member
    def SourceData(self) -> object:
        """The block its sparklines' sources make, from the first one's first corner to the last one's last; where
        they make none, as _regular has it, each source in turn, "#REF!" for one deleted
        (tests/fixtures/sparkline_shifts_c.json)."""
        sources = [line.source for line in self.group.sparklines]
        own = self.sheet.name
        if len(sources) == 1:
            return '#REF!' if sources[0].broken else sources[0].text(own)
        if not _regular(self.group.sparklines):
            return ','.join('#REF!' if source.broken else source.text(own) for source in sources)
        return Source(sources[0].sheet, sources[0].first, sources[-1].last).text(own)

    @setter('SourceData')
    def _set_source_data(self, value: object) -> None:
        self.ModifySourceData(value)

    @member
    def DateRange(self) -> object:
        """The date range as given, the group's own sheet left out; none is error 1004
        (tests/fixtures/sparkline_dates.json). One whose cells were deleted reads "": the group keeps its date
        axis and the file no range (tests/fixtures/sparkline_shifts_c.json)."""
        if self.group.dates is None:
            if self.group.attributes.get('dateAxis') == '1':
                return ''
            raise error(1004)
        sheet, reference = self.group.dates
        return reference if sheet is None else f'{_quoted(sheet)}!{reference}'

    @setter('DateRange')
    def _set_date_range(self, value: object) -> None:
        self.ModifyDateRange(value)

    @method
    def ModifyDateRange(self, DateRange: object = MISSING) -> object:
        """A new date range, which makes the horizontal axis a date axis; text that is no reference is error 1004,
        and a sheet the workbook does not have is left out, as Excel leaves it out."""
        found = _dates(to_text(DateRange), self.sheet.name)
        if found is None:
            raise error(1004)
        if found[0] is not None and _book_sheet(self.sheet.book, found[0]) is None:
            found = (None, found[1])
        self.group.dates = found
        self.group.attributes['dateAxis'] = '1'
        self.group.attributes = _ordered(self.group.attributes)
        self.edited()
        return EMPTY

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
        """xlNotPlotted for gaps, xlZero, which a file leaves out, or xlInterpolated for a span."""
        written = self.group.attributes.get('displayEmptyCellsAs', 'zero')
        found = next((kind for kind, name in BLANKS.items() if name == written), None)
        if found is None:
            raise VBAUnsupportedError('a sparkline group showing empty cells another way is not implemented')
        return VBAInt(found, 'Long')

    @setter('DisplayBlanksAs')
    def _set_display_blanks_as(self, value: object) -> None:
        kind = int(to_integer(value, 'Long'))
        if kind not in BLANKS:
            raise error(1004)
        if BLANKS[kind] == 'zero':
            self.group.attributes.pop('displayEmptyCellsAs', None)
        else:
            self.group.attributes['displayEmptyCellsAs'] = BLANKS[kind]
        self.group.attributes = _ordered(self.group.attributes)
        self.edited()

    @member
    def DisplayHidden(self) -> object:
        return self.group.attributes.get('displayHidden', '0') in ('1', 'true')

    @setter('DisplayHidden')
    def _set_display_hidden(self, value: object) -> None:
        if _flag(value):
            self.group.attributes['displayHidden'] = '1'
        else:
            self.group.attributes.pop('displayHidden', None)
        self.group.attributes = _ordered(self.group.attributes)
        self.edited()

    @member
    def SeriesColor(self) -> object:
        return SparkFormatColor(self, 'colorSeries')

    @member
    def Points(self) -> object:
        return SparkPoints(self)

    @member
    def PlotBy(self) -> object:
        """xlSparklineNonSquare for a source that is no square; a set is error 1004 there."""
        if self._square():
            raise VBAUnsupportedError('PlotBy of a sparkline group with a square source is not implemented')
        return VBAInt(0, 'Long')

    @setter('PlotBy')
    def _set_plot_by(self, value: object) -> None:
        if self._square():
            raise VBAUnsupportedError('PlotBy of a sparkline group with a square source is not implemented')
        raise error(1004)

    def _square(self) -> bool:
        areas = [line.source.area() for line in self.group.sparklines]
        rows = max(area.bottom for area in areas) - min(area.top for area in areas) + 1
        columns = max(area.right for area in areas) - min(area.left for area in areas) + 1
        return rows == columns and rows > 1

    @member
    def Axes(self) -> object:
        return SparkAxes(self)


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


def _flag(value: object) -> bool:
    if value is NULL:
        raise error(13)
    return to_bool(value)


class SparkFormatColor(ExcelObject):
    """A group's colour as SeriesColor and each SparkColor's Color give it: an RGB or a theme colour, tinted or not
    (tests/fixtures/sparkline_points.json)."""

    vba_type_name = 'FormatColor'

    def __init__(self, parent: SparklineGroup, element: str) -> None:
        self.parent, self.element = parent, element

    def _found(self) -> int:
        for place, child in enumerate(self.parent.group.children):
            if child.startswith(f'<x14:{self.element} ') or child == f'<x14:{self.element}/>':
                return place
        raise VBAUnsupportedError(f'a sparkline group with no {self.element} is not implemented')

    def _attributes(self) -> dict[str, str]:
        return dict(re.findall(r'([\w:]+)="([^"]*)"', self.parent.group.children[self._found()]))

    def _write(self, attributes: dict[str, str]) -> None:
        written = ''.join(f' {key}="{escape(value)}"' for key, value in attributes.items())
        markup = f'<x14:{self.element}{written}/>'
        self.parent.group.children[self._found()] = markup
        self.parent.edited()

    def _color(self) -> Color:
        from pyopenvba.apps.excel._styles import Color, tint_n

        attributes = self._attributes()
        tint = tint_n(attributes['tint']) if 'tint' in attributes else 0
        if 'rgb' in attributes:
            return Color('rgb', attributes['rgb'], tint)
        if 'theme' in attributes:
            return Color('theme', int(attributes['theme']), tint)
        raise VBAUnsupportedError('a sparkline colour of this form is not implemented')

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Color(self) -> object:
        from pyopenvba.apps.excel._styles import bgr

        return float(bgr(self.parent.sheet.book.stylesheet.colors.rrggbb(self._color(), '000000')))

    @setter('Color')
    def _set_color(self, value: object) -> None:
        """An RGB, the theme colour and tint gone."""
        from pyopenvba.apps.excel._styles import Color

        self._write({'rgb': str(Color.from_bgr(int(to_integer(value, 'Long'))).value)})

    @member
    def ThemeColor(self) -> object:
        attributes = self._attributes()
        return VBAInt(int(attributes['theme']) + 1 if 'theme' in attributes else 0, 'Long')

    @setter('ThemeColor')
    def _set_theme_color(self, value: object) -> None:
        """One of the theme's 12 colours, error 9 else; the RGB and tint go."""
        index = int(to_integer(value, 'Long'))
        if not 1 <= index <= 12:
            raise error(9, 'Subscript out of range')
        self._write({'theme': str(index - 1)})

    @member
    def TintAndShade(self) -> object:
        from pyopenvba.apps.excel._styles import TINT_SCALE

        return self._color().tint / TINT_SCALE

    @setter('TintAndShade')
    def _set_tint_and_shade(self, value: object) -> None:
        """-1 to 1, error 5 else, kept as Excel keeps it, n / 32767 truncated toward zero."""
        from pyopenvba.apps.excel._styles import TINT_SCALE, excel_double, tint_of

        tint = float(to_number(value))
        if not -1 <= tint <= 1:
            raise error(5, 'Invalid procedure call or argument')
        attributes = {key: text for key, text in self._attributes().items() if key != 'tint'}
        n = tint_of(tint)
        if n:
            attributes['tint'] = excel_double(n / TINT_SCALE)
        self._write(attributes)


class SparkColor(ExcelObject):
    """A point's marking or the horizontal axis: whether it shows, and its colour."""

    vba_type_name = 'SparkColor'

    def __init__(self, parent: SparklineGroup, attribute: str, element: str, holder: ExcelObject) -> None:
        self.parent, self.attribute, self.element, self.holder = parent, attribute, element, holder

    @member
    def Parent(self) -> object:
        """The SparkPoints or the axis it came from (tests/fixtures/object_trio.json)."""
        return self.holder

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Visible(self) -> object:
        return self.parent.group.attributes.get(self.attribute, '0') in ('1', 'true')

    @setter('Visible')
    def _set_visible(self, value: object) -> None:
        if _flag(value):
            self.parent.group.attributes[self.attribute] = '1'
        else:
            self.parent.group.attributes.pop(self.attribute, None)
        self.parent.group.attributes = _ordered(self.parent.group.attributes)
        self.parent.edited()

    @member
    def Color(self) -> object:
        return SparkFormatColor(self.parent, self.element)


class SparkPoints(ExcelObject):
    vba_type_name = 'SparkPoints'

    def __init__(self, parent: SparklineGroup) -> None:
        self.parent = parent

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Markers(self) -> object:
        return SparkColor(self.parent, 'markers', 'colorMarkers', self)

    @member
    def Highpoint(self) -> object:
        return SparkColor(self.parent, 'high', 'colorHigh', self)

    @member
    def Lowpoint(self) -> object:
        return SparkColor(self.parent, 'low', 'colorLow', self)

    @member
    def Firstpoint(self) -> object:
        return SparkColor(self.parent, 'first', 'colorFirst', self)

    @member
    def Lastpoint(self) -> object:
        return SparkColor(self.parent, 'last', 'colorLast', self)

    @member
    def Negative(self) -> object:
        return SparkColor(self.parent, 'negative', 'colorNegative', self)


class SparkAxes(ExcelObject):
    vba_type_name = 'SparkAxes'

    def __init__(self, parent: SparklineGroup) -> None:
        self.parent = parent

    @member
    def Parent(self) -> object:
        return self.parent

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Horizontal(self) -> object:
        return SparkHorizontalAxis(self.parent, self)

    @member
    def Vertical(self) -> object:
        return SparkVerticalAxis(self.parent, self)


class SparkHorizontalAxis(ExcelObject):
    vba_type_name = 'SparkHorizontalAxis'

    def __init__(self, parent: SparklineGroup, axes: SparkAxes) -> None:
        self.parent, self.axes = parent, axes

    @member
    def Parent(self) -> object:
        return self.axes

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def Axis(self) -> object:
        """The axis line: displayXAxis, and its colour colorAxis."""
        return SparkColor(self.parent, 'displayXAxis', 'colorAxis', self)

    @member
    def IsDateAxis(self) -> object:
        return self.parent.group.attributes.get('dateAxis', '0') in ('1', 'true')

    @member
    def RightToLeftPlotOrder(self) -> object:
        return self.parent.group.attributes.get('rightToLeft', '0') in ('1', 'true')

    @setter('RightToLeftPlotOrder')
    def _set_right_to_left_plot_order(self, value: object) -> None:
        """rightToLeft, left out at False (tests/fixtures/sparkline_dates.json)."""
        if _flag(value):
            self.parent.group.attributes['rightToLeft'] = '1'
        else:
            self.parent.group.attributes.pop('rightToLeft', None)
        self.parent.group.attributes = _ordered(self.parent.group.attributes)
        self.parent.edited()

    @member
    def RightToLeft(self) -> object:
        """Error 438 in Excel, read or set (tests/fixtures/sparkline_points.json)."""
        raise error(438, "Object doesn't support this property or method")

    @setter('RightToLeft')
    def _set_right_to_left(self, value: object) -> None:
        raise error(438, "Object doesn't support this property or method")


#: XlSparkScale by the minAxisType or maxAxisType a file gives it; xlSparkScaleSingle has none.
SCALES = {1: 'group', 2: None, 3: 'custom'}


class SparkVerticalAxis(ExcelObject):
    vba_type_name = 'SparkVerticalAxis'

    def __init__(self, parent: SparklineGroup, axes: SparkAxes) -> None:
        self.parent, self.axes = parent, axes

    def _scale(self, attribute: str) -> object:
        written = self.parent.group.attributes.get(attribute, 'individual')
        found = next((kind for kind, name in SCALES.items() if (name or 'individual') == written), None)
        if found is None:
            raise VBAUnsupportedError('a sparkline axis scaled another way is not implemented')
        return VBAInt(found, 'Long')

    def _set_scale(self, attribute: str, value: object) -> None:
        kind = int(to_integer(value, 'Long'))
        if kind not in SCALES:
            raise error(1004)
        name = SCALES[kind]
        if name is None:
            self.parent.group.attributes.pop(attribute, None)
        else:
            self.parent.group.attributes[attribute] = name
        self.parent.group.attributes = _ordered(self.parent.group.attributes)
        self.parent.edited()

    def _value(self, attribute: str) -> object:
        return float(self.parent.group.attributes.get(attribute, '0'))

    def _set_value(self, attribute: str, value: object) -> None:
        number = float(to_number(value))
        if number == 0:
            self.parent.group.attributes.pop(attribute, None)
        else:
            # Seventeen digits, as short as they go: 7, -2.5, 1.0000000000000001E+300.
            self.parent.group.attributes[attribute] = f'{number:.17G}'
        self.parent.group.attributes = _ordered(self.parent.group.attributes)
        self.parent.edited()

    @member
    def Parent(self) -> object:
        return self.axes

    @member
    def Application(self) -> object:
        return self.parent.sheet.book.application

    @member
    def MinScaleType(self) -> object:
        return self._scale('minAxisType')

    @setter('MinScaleType')
    def _set_min_scale_type(self, value: object) -> None:
        self._set_scale('minAxisType', value)

    @member
    def MaxScaleType(self) -> object:
        return self._scale('maxAxisType')

    @setter('MaxScaleType')
    def _set_max_scale_type(self, value: object) -> None:
        self._set_scale('maxAxisType', value)

    @member
    def CustomMinScaleValue(self) -> object:
        return self._value('manualMin')

    @setter('CustomMinScaleValue')
    def _set_custom_min(self, value: object) -> None:
        self._set_value('manualMin', value)

    @member
    def CustomMaxScaleValue(self) -> object:
        return self._value('manualMax')

    @setter('CustomMaxScaleValue')
    def _set_custom_max(self, value: object) -> None:
        self._set_value('manualMax', value)


def rewritten(sheet: Worksheet) -> bool:
    """Whether a save writes the sheet's groups: where a macro changed them, or where it holds several, which
    Excel writes the other way round from a reopened sheet, changed or not (tests/fixtures/sparkline_shifts_e.json)."""
    if sheet.sparklines is None:
        package = sheet.book.package
        if package is None or not sheet.part_name or not package.has(sheet.part_name) \
                or URI.encode() not in package.read(sheet.part_name):
            return False
        state(sheet)
    assert sheet.sparklines is not None
    return sheet.sparklines.changed or len(sheet.sparklines.groups) > 1


def written(sheet: Worksheet, xml: str) -> str:
    """The sheet's part with its sparkline groups as Excel writes them, where a macro changed them, or where it
    holds several: Excel writes a reopened sheet's groups the other way round, changed or not
    (tests/fixtures/sparkline_shifts_e.json)."""
    if not rewritten(sheet):
        return xml
    assert sheet.sparklines is not None
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
    dates = ''
    if group.dates is not None:
        dates = f'<xm:f>{escape_text(_quoted(group.dates[0] or sheet.name) + "!" + group.dates[1])}</xm:f>'
    lines = ''.join('<x14:sparkline>' + (f'<xm:f>{escape_text(line.source.formula(sheet.name))}</xm:f>'
                                         if not line.source.broken else '') +
                    f'<xm:sqref>{_cell_name(line.at)}</xm:sqref></x14:sparkline>' for line in group.sparklines)
    return f'<x14:sparklineGroup{attributes}>{"".join(group.children)}{dates}<x14:sparklines>{lines}</x14:sparklines>' \
           '</x14:sparklineGroup>'


def _cell_name(at: tuple[int, int]) -> str:
    from pyopenvba._a1 import column_letter

    return f'{column_letter(at[1])}{at[0]}'
