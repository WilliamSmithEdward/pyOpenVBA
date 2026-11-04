"""Worksheet.QueryTables over local text files, as live Excel answers them (scripts/measure_query_tables.py,
tests/fixtures/query_tables.json).

A query table is its settings until it is refreshed. A refresh reads the file with the parser TextToColumns and OpenText
use, places the rows by RefreshStyle, and names the result on the sheet: that name is ResultRange, so it follows edits to
the sheet as any name does. Only a table refreshed at least once reaches the file, as a queryTable part and a connection.

Refused rather than guessed: any source but a text file, a file with bytes past ASCII, setting PostText or a URL
Connection (each stops Excel outright), TextFilePromptOnRefresh, and a table's width changing between refreshes. Column
widths are left alone: Excel's AdjustColumnWidth follows a rule the 144-DPI measurements did not settle. A workbook loaded
with query tables is still refused by the sheet's QueryTables (see _empty_collections).
"""
from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, cast

from pyopenvba._a1 import Area, column_letter, parse_area, quote_sheet, split_sheet
from pyopenvba._xml import escape
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, member, method, setter
from pyopenvba.interpreter._values import (EMPTY, MISSING, VBAArray, VBAErrorValue, VBAInt, error, to_bool, to_integer,
                                           to_text)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

_APP_ERROR = (1004, 'Application-defined or object-defined error')
_INVALID = (5, 'Invalid procedure call or argument')
_MEMORY = (7, 'Out of memory')
#: QueryTable.RefreshStyle and what the file calls each; xlInsertDeleteCells is the default and is not written.
_INSERT_DELETE, _OVERWRITE, _INSERT_ROWS = 1, 0, 2
_GROWTH = {_OVERWRITE: 'overwriteClear', _INSERT_ROWS: 'insertClear'}
_PLATFORMS = {1: 'mac', 2: None, 3: 'dos'}
_FIELD_TYPES = {1: None, 2: 'text', 3: 'MDY', 4: 'DMY', 5: 'YMD', 6: 'MYD', 7: 'DYM', 8: 'YDM', 9: 'skip'}
_QUALIFIERS = {1: None, 2: 'singleQuote', -4142: 'none'}
#: What RobustConnect's xlAlways and xlNever write; xlAsRequired is the default and is not written.
_RECONNECTIONS = {1: 2, 2: 3}
_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
CONNECTIONS_PART = 'xl/connections.xml'
CONNECTIONS_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
CONNECTIONS_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
QUERY_TABLE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'
QUERY_TABLE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable'


def _source(connection: object) -> str:
    """The file a connection string names. Text that names no kind of source is what Excel refuses; another kind is what
    the model does."""
    text = to_text(connection)
    kind, separator, rest = text.partition(';')
    if not separator:
        raise error(*_APP_ERROR)
    if kind.upper() != 'TEXT':
        raise VBAUnsupportedError(f'A {kind} query table is not implemented: only a local text file is')
    return rest


class QueryTables(VBACollection, ItemDefault, Created):
    vba_type_name = 'QueryTables'

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def vba_items(self) -> list[object]:
        return list(self.sheet.query_tables)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for table in self.sheet.query_tables:
                if table.name.lower() == index.lower():
                    return table
            raise error(9, 'Subscript out of range')
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(9, 'Subscript out of range')
        return items[position - 1]

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @method
    def Add(self, Connection: object = MISSING, Destination: object = MISSING, Sql: object = MISSING) -> object:  # noqa: N803 - native named arguments
        from pyopenvba.apps.excel._model import Range

        source = _source(Connection)
        if not isinstance(Destination, Range):
            raise error(424, 'Object required')
        if Destination.sheet is not self.sheet or Sql is not MISSING:
            raise error(*_INVALID)
        self.sheet.query_table_count += 1
        table = QueryTable(self.sheet, source, Destination, f'ExternalData_{self.sheet.query_table_count}')
        self.sheet.query_tables.insert(0, table)
        self.sheet.book.saved = False
        return table


class QueryTable(Created):
    vba_type_name = 'QueryTable'

    def __init__(self, sheet: Worksheet, source: str, destination: Range, name: str) -> None:
        self.sheet, self.source, self.destination, self.name = sheet, source, destination, name
        self.settings: dict[str, object] = {}
        self.command_text: str | None = None
        self.refreshed = False
        #: The rows of the last refresh and the columns they filled; None before one, or after an empty file.
        self.size: tuple[int, int] | None = None
        self.emptied = False
        #: The name of the file the connection was first made for, which names the connection.
        self.stem = Path(source.replace('\\', '/')).stem

    def setting(self, name: str, default: object) -> object:
        return self.settings.get(name, default)

    # -- where it is

    def _entry(self) -> object:
        wanted = _name_of(self.name).lower()
        for entry in self.sheet.book.names_.entries:
            scope, bare = split_sheet(entry.name)
            if scope.lower() == self.sheet.name.lower() and bare.lower() == wanted:
                return entry
        return None

    def result_area(self) -> Area | None:
        """The block the table's name says, None where it has none."""
        entry = self._entry()
        if not self.refreshed or entry is None:
            return None
        return parse_area(str(getattr(entry, 'refers_to')).rpartition('!')[2].replace('$', ''), sheet=self.sheet.name)

    def _rename(self, wanted: str) -> None:
        entry = self._entry()
        self.name = wanted
        if entry is not None:
            setattr(entry, 'name', f'{quote_sheet(self.sheet.name)}!{_name_of(wanted)}')
            self.sheet.book.names_.changed = True

    def _named(self, area: Area) -> None:
        from pyopenvba.apps.excel._model import NameEntry

        names = self.sheet.book.names_
        text = f'={quote_sheet(self.sheet.name)}!{_address(area)}'
        entry = self._entry()
        if entry is None:
            names.entries.append(NameEntry(f'{quote_sheet(self.sheet.name)}!{_name_of(self.name)}', text, self.sheet.book))
        else:
            setattr(entry, 'refers_to', text)
        names.changed = True

    # -- refreshing

    def _lines(self) -> list[str]:
        from pyopenvba.apps.excel._text_files import _lines  # pyright: ignore[reportPrivateUsage]

        self.sheet.book.application.outside('Refreshing a query table')
        path = Path(self.source)
        if not path.is_file():
            raise error(*_APP_ERROR)
        data = path.read_bytes()
        if any(byte > 127 for byte in data):
            raise VBAUnsupportedError('A query table over a text file with bytes past ASCII is not implemented')
        lines = _lines(data.decode('ascii'))
        first = int(to_integer(self.setting('TextFileStartRow', 1), 'Long'))
        return lines[first - 1:]

    def _rows(self) -> list[list[tuple[str, int]]]:
        """Each line's fields, with the kind of column each is in, those skipped gone."""
        from pyopenvba.apps.excel import _text_columns as columns

        kinds = [int(kind) for kind in self._list('TextFileColumnDataTypes')]
        fixed = self.setting('TextFileParseType', 1) == 2
        widths = [int(width) for width in self._list('TextFileFixedColumnWidths')]
        marks = {mark for name, mark in (('TextFileTabDelimiter', '\t'), ('TextFileSemicolonDelimiter', ';'), ('TextFileCommaDelimiter', ','), ('TextFileSpaceDelimiter', ' '))
                 if self.setting(name, name == 'TextFileTabDelimiter')}
        other = to_text(self.setting('TextFileOtherDelimiter', ''))
        if other:
            marks.add(other)
        quote = columns.QUALIFIERS.get(int(to_integer(self.setting('TextFileTextQualifier', 1), 'Long')), '"')
        rows: list[list[tuple[str, int]]] = []
        for line in self._lines():
            if fixed:
                fields: list[str] = []
                at = 0
                for width in widths:
                    fields.append(line[at:at + width])
                    at += width
                fields.append(line[at:])
            else:
                fields = columns.split_fields(line, marks, quote, bool(self.setting('TextFileConsecutiveDelimiter', False)))
            kept = [(field, kinds[index] if index < len(kinds) else columns.GENERAL) for index, field in enumerate(fields)]
            rows.append([one for one in kept if one[1] != columns.SKIP])
        return rows

    def _list(self, name: str) -> list[int]:
        found = self.setting(name, None)
        return cast(list[int], found) if isinstance(found, list) else []

    def refresh(self) -> object:
        from pyopenvba.apps.excel import _text_columns as columns
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            raise error(*_APP_ERROR)
        if (self.refreshed and not self.setting('EnableRefresh', True)) or self.command_text is not None:
            raise error(*_APP_ERROR)
        rows = self._rows()
        origin = self.destination.areas[0]
        if not rows and self.setting('TextFileStartRow', 1) == 1:
            self.refreshed, self.emptied = True, True
            return True
        height, width = max(1, len(rows)), max([len(row) for row in rows] or [1])
        old = self.result_area()
        if old is not None and old.right - old.left + 1 != width:
            raise VBAUnsupportedError('A query table whose width changes between refreshes is not implemented')
        self._place(origin, old, height, width)
        self.refreshed, self.emptied, self.size = True, False, (height, width)
        separators = (to_text(self.setting('TextFileDecimalSeparator', '.'))[:1] or '.', to_text(self.setting('TextFileThousandsSeparator', ','))[:1] or ',')
        minus = bool(self.setting('TextFileTrailingMinusNumbers', False))
        start = old if old is not None else origin
        target = self._range(Area(start.top, start.left, start.top, start.left, self.sheet.name))
        for row, fields in enumerate(rows, start=start.top):
            for column, (field, kind) in enumerate(fields, start=start.left):
                if field:
                    columns.write_field(target, row, column, field, kind, separators, minus)
        if old is not None and height > old.bottom - old.top + 1 and self.setting('FillAdjacentFormulas', False):
            self._fill_down(old, start.top + height - 1)
        self._named(Area(start.top, start.left, start.top + height - 1, start.left + width - 1, self.sheet.name))
        self.sheet.book.saved = False
        return True

    def _range(self, area: Area) -> Range:
        from pyopenvba.apps.excel._model import Range

        return Range(self.sheet, [area])

    def _occupied(self, area: Area) -> list[tuple[int, int]]:
        return [(row, column) for row, column in self.sheet.cells_ if area.top <= row <= area.bottom and area.left <= column <= area.right
                and self.sheet.cells_[(row, column)].value not in (None, EMPTY, '')]

    def _place(self, origin: Area, old: Area | None, height: int, width: int) -> None:
        """Room for the rows, by RefreshStyle."""
        from pyopenvba.apps.excel._editing import shift_cells

        style = self.setting('RefreshStyle', _INSERT_DELETE)
        name = self.sheet.name
        if old is None:
            block = Area(origin.top, origin.left, origin.top + height - 1, origin.left + width - 1, name)
            occupied = self._occupied(block)
            if occupied:
                if style != _INSERT_DELETE:
                    raise VBAUnsupportedError('A first refresh onto cells in use, but for xlInsertDeleteCells, is not implemented')
                # Whatever is in its way moves across, in the block's rows, by the columns from the first in use.
                band = Area(block.top, min(column for _, column in occupied), block.bottom, block.right, name)
                shift_cells(self._range(band), band, delete=False, vertical=False)
            return
        self._range(old).vba_call('ClearContents', [])
        before = old.bottom - old.top + 1
        last = old.top + height - 1
        if style == _INSERT_DELETE and height != before:
            band = Area(min(old.bottom, last) + 1, old.left, max(old.bottom, last), old.right, name)
            shift_cells(self._range(band), band, delete=height < before, vertical=True)
        elif style == _INSERT_ROWS and height > before:
            in_way = self._occupied(Area(old.bottom + 1, old.left, last, old.right, name))
            if len(in_way) > 1:
                raise VBAUnsupportedError('xlInsertEntireRows growing over more than one cell in use is not implemented')
            for row, column in in_way:
                below = self._range(Area(last + 1, column, last + 1, column, name))
                if self._occupied(below.areas[0]):
                    raise VBAUnsupportedError('xlInsertEntireRows moving a cell onto another is not implemented')
                held = self._range(Area(row, column, row, column, name))
                below.vba_set('Value', held.vba_get('Value', []))
                held.vba_call('ClearContents', [])

    def _fill_down(self, old: Area, last: int) -> None:
        """Formulas beside the table carried down its new rows."""
        column = old.right + 1
        while True:
            cell = self.sheet.cells_.get((old.bottom, column))
            if cell is None or not getattr(cell, 'formula', None):
                return
            area = Area(old.bottom, column, last, column, self.sheet.name)
            self._range(area).vba_call('FillDown', [])
            column += 1

    # -- members with a rule of their own

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Name(self) -> object:
        return self.name

    @setter('Name')
    def _name_set(self, value: object) -> None:
        self._rename(to_text(value))

    @member
    def Connection(self) -> object:
        return f'TEXT;{self.source}'

    @setter('Connection')
    def _connection(self, value: object) -> None:
        text = to_text(value)
        if not text.upper().startswith('TEXT;') and ';' in text:
            raise VBAUnsupportedError('Setting a query table\'s Connection to another kind of source stops Excel, and is refused')
        self.source = _source(value)

    @member
    def Destination(self) -> object:
        return self.destination

    @member
    def ResultRange(self) -> object:
        if self.emptied:
            raise error(*_MEMORY)
        area = self.result_area()
        if area is None:
            raise error(*_APP_ERROR)
        return self._range(area)

    @member
    def CommandText(self) -> object:
        if self.command_text is None:
            raise error(*_MEMORY)
        return self.command_text

    @setter('CommandText')
    def _command_text(self, value: object) -> None:
        self.command_text = to_text(value)

    @member
    def CommandType(self) -> object:
        raise error(*_APP_ERROR)

    @setter('CommandType')
    def _command_type(self, value: object) -> None:
        del value
        raise error(*_INVALID)

    @member
    def PostText(self) -> object:
        raise error(*_APP_ERROR)

    @setter('PostText')
    def _post_text(self, value: object) -> None:
        del value
        raise VBAUnsupportedError('Setting QueryTable.PostText stops Excel, and is refused')

    @member
    def ListObject(self) -> object:
        raise error(*_APP_ERROR)

    @member
    def Recordset(self) -> object:
        raise error(*_APP_ERROR)

    @member
    def TextFileColumnDataTypes(self, Index: object = MISSING) -> object:  # noqa: N803 - what an index to the array arrives as
        if Index is not MISSING:
            raise error(451, 'Property let procedure not defined and property get procedure did not return an object')
        return _array(self._list('TextFileColumnDataTypes'))

    @setter('TextFileColumnDataTypes')
    def _column_types(self, value: object) -> None:
        self.settings['TextFileColumnDataTypes'] = _whole_numbers(value)

    @member
    def TextFileFixedColumnWidths(self, Index: object = MISSING) -> object:  # noqa: N803 - what an index to the array arrives as
        if Index is not MISSING:
            raise error(451, 'Property let procedure not defined and property get procedure did not return an object')
        found = self._list('TextFileFixedColumnWidths')
        return _array(found) if found else VBAErrorValue(2042)

    @setter('TextFileFixedColumnWidths')
    def _fixed_widths(self, value: object) -> None:
        self.settings['TextFileFixedColumnWidths'] = _whole_numbers(value)

    @method
    def Refresh(self, BackgroundQuery: object = MISSING) -> object:  # noqa: N803 - native named argument
        del BackgroundQuery
        return self.refresh()

    @method
    def CancelRefresh(self) -> object:
        raise error(*_APP_ERROR)

    @method
    def ResetTimer(self) -> object:
        return EMPTY

    @method
    def SaveAsODC(self, ODCFileName: object = MISSING, Description: object = MISSING, Keywords: object = MISSING) -> object:  # noqa: N803 - native named arguments
        del ODCFileName, Description, Keywords
        raise error(*_INVALID)

    @method
    def Delete(self) -> object:
        """The table gone, its cells and its name left."""
        self.sheet.query_tables.remove(self)
        self.sheet.book.saved = False
        return EMPTY


def _name_of(table: str) -> str:
    """The defined name a table's name gives: what a name cannot hold becomes an underscore."""
    return re.sub(r'[^\w.]', '_', table)


def _address(area: Area) -> str:
    first = f'${column_letter(area.left)}${area.top}'
    if (area.top, area.left) == (area.bottom, area.right):
        return first
    return f'{first}:${column_letter(area.right)}${area.bottom}'


def _array(values: list[int]) -> VBAArray:
    found = VBAArray([(1, max(1, len(values)))])
    for place, value in enumerate(values, start=1):
        found.set([place], VBAInt(value, 'Long'))
    return found


def _whole_numbers(value: object) -> list[int]:
    if not isinstance(value, VBAArray):
        raise error(*_INVALID)
    return [int(to_integer(item, 'Long')) for item in value.items]


def _kept(name: str, default: object, convert: Callable[[object], object]) -> MemberSpec:
    def get(table: QueryTable) -> object:
        found = table.setting(name, default)
        return VBAInt(found, 'Long') if isinstance(found, int) and not isinstance(found, bool) else found

    def put(table: QueryTable, value: object) -> None:
        table.settings[name] = convert(value)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _fixed(name: str, value: object) -> MemberSpec:
    def get(table: QueryTable) -> object:
        del table
        return value

    return MemberSpec(name=name, kind='property', getter=get)


def _only_for_the_web(name: str) -> MemberSpec:
    def get(table: QueryTable) -> object:
        del table
        raise error(*_APP_ERROR)

    def put(table: QueryTable, value: object) -> None:
        del table, value
        raise error(*_APP_ERROR)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _long(value: object) -> object:
    return int(to_integer(value, 'Long'))


def _first_character(value: object) -> object:
    return to_text(value)[:1]


def _in(allowed: set[int], number: int = 5) -> Callable[[object], object]:
    def convert(value: object) -> object:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in allowed:
            raise error(number, 'Invalid procedure call or argument')
        return wanted

    return convert


def _at_least(lowest: int, number: int) -> Callable[[object], object]:
    def convert(value: object) -> object:
        wanted = int(to_integer(value, 'Long'))
        if wanted < lowest:
            raise error(number, 'Invalid procedure call or argument')
        return wanted

    return convert


def _platform(value: object) -> object:
    wanted = int(to_integer(value, 'Long'))
    if wanted not in _PLATFORMS:
        raise VBAUnsupportedError('A query table over a file in a code page is not implemented')
    return wanted


def _no_prompt(value: object) -> object:
    if to_bool(value):
        raise VBAUnsupportedError('TextFilePromptOnRefresh asks the user for a file, and is refused')
    return False


for _spec in (_kept('AdjustColumnWidth', True, to_bool), _kept('BackgroundQuery', True, to_bool), _kept('EnableEditing', True, to_bool), _kept('EnableRefresh', True, to_bool),
              _kept('FieldNames', True, to_bool), _kept('FillAdjacentFormulas', False, to_bool), _kept('MaintainConnection', False, to_bool), _kept('PreserveColumnInfo', False, to_bool),
              _kept('PreserveFormatting', True, to_bool), _kept('RefreshOnFileOpen', False, to_bool), _kept('RefreshPeriod', 0, _at_least(0, 5)), _kept('RefreshStyle', _INSERT_DELETE, _in({0, 1, 2})),
              _kept('RobustConnect', 0, _in({0, 1, 2})), _kept('RowNumbers', False, to_bool), _kept('SaveData', True, to_bool), _kept('SavePassword', False, to_bool), _kept('SourceConnectionFile', '', to_text),
              _kept('SourceDataFile', '', to_text), _kept('TextFileCommaDelimiter', False, to_bool), _kept('TextFileConsecutiveDelimiter', False, to_bool), _kept('TextFileDecimalSeparator', '.', to_text),
              _kept('TextFileOtherDelimiter', '', _first_character), _kept('TextFileParseType', 1, _long), _kept('TextFilePlatform', 2, _platform), _kept('TextFilePromptOnRefresh', False, _no_prompt),
              _kept('TextFileSemicolonDelimiter', False, to_bool), _kept('TextFileSpaceDelimiter', False, to_bool), _kept('TextFileStartRow', 1, _at_least(1, 1004)), _kept('TextFileTabDelimiter', True, to_bool),
              _kept('TextFileTextQualifier', 1, _long), _kept('TextFileThousandsSeparator', ',', to_text), _kept('TextFileTrailingMinusNumbers', False, to_bool), _kept('TextFileVisualLayout', 2, _long),
              _fixed('FetchedRowOverflow', False), _fixed('QueryType', VBAInt(6, 'Long')), _fixed('Refreshing', False),
              *(_only_for_the_web(_name) for _name in ('WebConsecutiveDelimitersAsOne', 'WebDisableDateRecognition', 'WebDisableRedirections', 'WebFormatting', 'WebPreFormattedTextToColumns', 'WebSelectionType', 'WebSingleBlockTextImport', 'WebTables'))):
    QueryTable.vba_add_member(_spec)


# -- the file

def _connection_markup(table: QueryTable, number: int, name: str) -> str:
    get = table.setting
    method_ = _RECONNECTIONS.get(int(str(get('RobustConnect', 0))))
    attributes = f' id="{number}"'
    if get('SourceConnectionFile', ''):
        attributes += f' odcFile="{escape(to_text(get("SourceConnectionFile", "")))}"'
    if get('SourceDataFile', ''):
        attributes += f' sourceFile="{escape(to_text(get("SourceDataFile", "")))}"'
    if get('RefreshPeriod', 0):
        attributes += f' interval="{get("RefreshPeriod", 0)}"'
    attributes += f' name="{escape(name)}" type="6"'
    if method_ is not None:
        attributes += f' reconnectionMethod="{method_}"'
    attributes += ' refreshedVersion="8"'
    if get('SavePassword', False):
        attributes += ' savePassword="1"'
    if get('BackgroundQuery', True):
        attributes += ' background="1"'
    if get('RefreshOnFileOpen', False):
        attributes += ' refreshOnLoad="1"'
    if get('SaveData', True):
        attributes += ' saveData="1"'
    text = ' prompt="0"'
    platform = _PLATFORMS[int(str(get('TextFilePlatform', 2)))]
    if platform:
        text += f' fileType="{platform}"'
    if get('TextFileStartRow', 1) != 1:
        text += f' firstRow="{get("TextFileStartRow", 1)}"'
    text += f' sourceFile="{escape(table.source)}"'
    if get('TextFileParseType', 1) == 2:
        text += ' delimited="0"'
    if get('TextFileDecimalSeparator', '.') != '.':
        text += f' decimal="{escape(to_text(get("TextFileDecimalSeparator", ".")))}"'
    if get('TextFileThousandsSeparator', ',') != ',':
        text += f' thousands="{escape(to_text(get("TextFileThousandsSeparator", ",")))}"'
    if not get('TextFileTabDelimiter', True):
        text += ' tab="0"'
    for name_, attribute in (('TextFileSpaceDelimiter', 'space'), ('TextFileCommaDelimiter', 'comma'), ('TextFileSemicolonDelimiter', 'semicolon'), ('TextFileConsecutiveDelimiter', 'consecutive')):
        if get(name_, False):
            text += f' {attribute}="1"'
    qualifier = _QUALIFIERS.get(int(str(get('TextFileTextQualifier', 1))))
    if qualifier:
        text += f' qualifier="{qualifier}"'
    if get('TextFileOtherDelimiter', ''):
        text += f' delimiter="{escape(to_text(get("TextFileOtherDelimiter", "")))}"'
    return f'<connection{attributes}><textPr{text}>{_fields_markup(table)}</textPr></connection>'


def _fields_markup(table: QueryTable) -> str:
    kinds = table._list('TextFileColumnDataTypes')  # pyright: ignore[reportPrivateUsage]
    widths = table._list('TextFileFixedColumnWidths')  # pyright: ignore[reportPrivateUsage]
    if widths and table.setting('TextFileParseType', 1) == 2:
        if kinds:
            raise VBAUnsupportedError('Saving a fixed-width query table with column types is not implemented')
        places = [sum(widths[:count]) for count in range(len(widths) + 1)]
        fields = ''.join('<textField/>' if place == 0 else f'<textField position="{place}"/>' for place in places)
        return f'<textFields count="{len(places)}">{fields}</textFields>'
    if len(kinds) < 2 and all(kind == 1 for kind in kinds):
        return '<textFields><textField/></textFields>'
    fields = ''.join('<textField/>' if _FIELD_TYPES.get(kind) is None else f'<textField type="{_FIELD_TYPES[kind]}"/>' for kind in kinds)
    return f'<textFields count="{len(kinds)}">{fields}</textFields>'


def _table_markup(table: QueryTable, number: int) -> str:
    get = table.setting
    attributes = f' name="{escape(table.name)}"'
    if not get('FieldNames', True):
        attributes += ' headers="0"'
    if get('RowNumbers', False):
        attributes += ' rowNumbers="1"'
    if not get('EnableRefresh', True):
        attributes += ' disableRefresh="1"'
    if not get('BackgroundQuery', True):
        attributes += ' backgroundRefresh="0"'
    if get('RefreshOnFileOpen', False):
        attributes += ' refreshOnLoad="1"'
    growth = _GROWTH.get(int(str(get('RefreshStyle', _INSERT_DELETE))))
    if growth:
        attributes += f' growShrinkType="{growth}"'
    if get('FillAdjacentFormulas', False):
        attributes += ' fillFormulas="1"'
    if not get('SaveData', True):
        attributes += ' removeDataOnSave="1"'
    if not get('EnableEditing', True):
        attributes += ' disableEdit="1"'
    if not get('PreserveFormatting', True):
        attributes += ' preserveFormatting="0"'
    if not get('AdjustColumnWidth', True):
        attributes += ' adjustColumnWidth="0"'
    return (f'{_DECLARATION}<queryTable xmlns="{_MAIN}"{attributes} connectionId="{number}" autoFormatId="0" applyNumberFormats="0" '
            'applyBorderFormats="0" applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0"/>')


def write(book: object, package: object, add_relationship: Callable[..., str], add_override: Callable[..., None]) -> None:
    """Every refreshed table's part, and the connections part, in a workbook that came with none."""
    sheets = getattr(book, 'sheets_')
    tables = [table for sheet in sheets for table in reversed(getattr(sheet, 'query_tables', [])) if table.refreshed and not table.emptied]
    if not tables:
        return
    if getattr(package, 'has')(CONNECTIONS_PART) and not getattr(book, 'query_parts_written', False):
        raise VBAUnsupportedError('Saving a query table in a workbook that has connections is not implemented')
    names: list[str] = []
    entries: list[str] = []
    write_part = getattr(package, 'write')
    for number, table in enumerate(tables, start=1):
        name = table.stem
        taken = 0
        while name in names:
            taken += 1
            name = f'{table.stem}{taken}'
        names.append(name)
        entries.append(_connection_markup(table, number, name))
        part = f'xl/queryTables/queryTable{number}.xml'
        if not getattr(package, 'has')(part):
            add_override(package, part, QUERY_TABLE_TYPE)
            add_relationship(package, table.sheet.part_name, f'../queryTables/queryTable{number}.xml', QUERY_TABLE_RELATIONSHIP)
        write_part(part, _table_markup(table, number).encode('utf-8'))
    if not getattr(package, 'has')(CONNECTIONS_PART):
        add_override(package, CONNECTIONS_PART, CONNECTIONS_TYPE)
        rels = 'xl/_rels/workbook.xml.rels'
        text = getattr(package, 'read')(rels).decode('utf-8')
        used = {int(found) for found in re.findall(r'Id="rId(\d+)"', text)}
        relationship = f'<Relationship Id="rId{max(used, default=0) + 1}" Type="{CONNECTIONS_RELATIONSHIP}" Target="connections.xml"/>'
        write_part(rels, text.replace('</Relationships>', relationship + '</Relationships>').encode('utf-8'))
    write_part(CONNECTIONS_PART, (f'{_DECLARATION}<connections xmlns="{_MAIN}">{"".join(entries)}</connections>').encode('utf-8'))
    setattr(book, 'query_parts_written', True)


__all__ = ['QueryTable', 'QueryTables', 'write']
