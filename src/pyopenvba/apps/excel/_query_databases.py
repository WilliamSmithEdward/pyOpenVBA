"""Query tables over a database, refreshed through the function the host sets with set_hook('Refresh', ...).

The host fetches; the model lays the rows out as Excel does, measured against an Access file through Excel's own ACE OLEDB provider
(scripts/measure_query_refresh.py, scripts/measure_query_refresh_layout.py):

- A value goes on the sheet by its Python type, as Excel puts one by its ADO type: text as text, even where it looks like a number, and
  an empty text as one; int, float and Decimal as numbers in General, Currency and Decimal included; bool as a Boolean; a date or
  datetime as a Date shown m/d/yyyy, its time kept in the value; None as an empty cell.
- The field names head the rows in bold, unless FieldNames is False. RowNumbers puts a column first, numbered from 0, under an
  empty bold heading. With no rows a blank row stands under the names.
- A first refresh moves what is in its way right, across its rows, by the table's width, whatever RefreshStyle says. On a later one,
  xlInsertDeleteCells moves what is below down or up by the change in rows, in the table's columns; xlOverwriteCells writes over it;
  xlInsertEntireRows inserts whole rows from the first row in its way to the new last row, and deletes none.
- FillAdjacentFormulas carries the formulas beside the first row down to the last. PreserveFormatting False takes the cells' formats
  off before the rows go in.
- BackgroundQuery is False for a table a macro adds. A refresh in the background leaves the sheet as it was and Refreshing True until
  the macro is over; ExcelApplication.run then refreshes it, and raises what failed after the run's own result.

A connection string's password never shows: Connection drops it once a refresh has run unless SavePassword is True, as Excel does,
and a save writes it only then. What Excel reads back is the provider's own fuller string; the model gives back the one the macro set.
That is a known difference, as is the uid Excel writes on each part.
"""
from __future__ import annotations

import datetime as _dt
import re
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba._xml import escape
from pyopenvba.apps.excel._query_tables import QueryTable, free_name
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBADate, VBAInt, error, to_bool, to_text

if TYPE_CHECKING:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Application, Range, Worksheet

#: The kinds of source a refresh hands to the host, by the prefix of the connection string.
KINDS = {'OLEDB': VBAInt(5, 'Long'), 'ODBC': VBAInt(1, 'Long')}
#: A key=value pair of a connection string whose key names a password, its value quoted or not.
_SECRET = re.compile(r"""(?i)((?:^|;)\s*[^;=]*(?:password|pwd)[^;=]*=\s*)("[^"]*"|'[^']*'|[^;]*)""")
_INSERT_DELETE, _OVERWRITE, _INSERT_ROWS = 1, 0, 2
_DATE_FORMAT = 'm/d/yyyy'


def masked(connection: str) -> str:
    """A connection string as the model shows one, each password ****."""
    return _SECRET.sub(lambda found: found.group(1) + '****', connection)


def without_passwords(connection: str) -> str:
    """A connection string with its passwords taken out, as Excel keeps one once SavePassword is False."""
    return _SECRET.sub(lambda found: '' if found.group(1).startswith(';') or found.start() == 0 else found.group(0), connection).lstrip(';')


def is_database(connection: str) -> bool:
    return connection.partition(';')[0].strip().upper() in KINDS


class DatabaseQueryTable(QueryTable):
    """A query table whose rows the host's Refresh function gives."""

    vba_type_name = 'QueryTable'

    def __init__(self, sheet: Worksheet, connection: str, destination: Range, name: str, command: str | None) -> None:
        kind, _, rest = connection.partition(';')
        self.kind, self.connection = kind.strip().upper(), rest
        super().__init__(sheet, rest, destination, name, free_name(sheet.book, 'Connection'))
        self.link.element.set('type', str(int(KINDS[self.kind])))
        self.command_text = command
        #: The field names of the last refresh, which a save lists.
        self.fields: list[str] = []
        #: Waiting for the run to end, as Excel's background refresh waits for the macro.
        self.pending = False
        self.password_dropped = False

    def setting(self, name: str, default: object) -> object:
        if name == 'BackgroundQuery':
            return self.settings.get(name, False)
        return super().setting(name, default)

    # -- members that differ from a text file's

    @member
    def Connection(self) -> object:
        text = without_passwords(self.connection) if self.password_dropped else self.connection
        return f'{self.kind};{text}'

    @setter('Connection')
    def _set_connection(self, value: object) -> None:
        text = to_text(value)
        if not is_database(text):
            raise VBAUnsupportedError("Setting a database query table's Connection to another kind of source is not implemented")
        kind, _, rest = text.partition(';')
        self.kind, self.connection, self.password_dropped = kind.strip().upper(), rest, False
        self.link.element.set('type', str(int(KINDS[self.kind])))

    @member
    def CommandText(self) -> object:
        return self.command_text or ''

    @setter('CommandText')
    def _set_command_text(self, value: object) -> None:
        self.command_text = to_text(value)

    @member
    def CommandType(self) -> object:
        return VBAInt(2, 'Long')

    @setter('CommandType')
    def _set_command_type(self, value: object) -> None:
        if int(to_text(value) or 0) != 2:
            raise VBAUnsupportedError('A query table command other than SQL text is not implemented')

    @member
    def QueryType(self) -> object:
        return KINDS[self.kind]

    @member
    def Refreshing(self) -> object:
        return self.pending

    @method
    def Refresh(self, BackgroundQuery: object = MISSING) -> object:  # noqa: N803 - native named argument
        from pyopenvba.apps.excel._hooks import hook_of

        application = self.sheet.book.application
        if hook_of(application, 'Refresh') is None:
            raise VBAUnsupportedError(f"Refreshing a query table over {self.kind} takes a function set with set_hook('Refresh', ...)")
        if self.pending:
            raise VBAUnsupportedError('Refreshing a query table while its background refresh waits is not implemented')
        background = to_bool(BackgroundQuery) if BackgroundQuery is not MISSING else bool(self.setting('BackgroundQuery', False))
        if background:
            self.pending = True
            pending(application).append(self)
            return True
        return self.refresh()

    # -- refreshing

    def refresh(self) -> object:
        from pyopenvba.apps.excel._hooks import ExternalSource, call, hook_of
        from pyopenvba.apps.excel._protection import enforced

        if enforced(self.sheet) is not None:
            raise error(1004, 'Application-defined or object-defined error')
        application = self.sheet.book.application
        hook = hook_of(application, 'Refresh')
        if hook is None:
            raise VBAUnsupportedError(f"Refreshing a query table over {self.kind} takes a function set with set_hook('Refresh', ...)")
        source = ExternalSource(self.kind, self.connection, self.command_text or '', 'xlCmdSql', (), bool(self.setting('FieldNames', True)))
        rows = _checked(call(application, hook, 'Refresh', 'QueryTable', self, self.sheet.book, {}, source=source))
        self._lay_out(rows)
        if not self.setting('SavePassword', False):
            self.password_dropped = True
        return True

    def _lay_out(self, rows: list[list[object]]) -> None:
        names, data = [str(one) for one in rows[0]], rows[1:]
        heading = bool(self.setting('FieldNames', True))
        numbered = bool(self.setting('RowNumbers', False))
        if numbered:
            names = ['', *names]
            data = [[float(at), *row] for at, row in enumerate(data)]
        body = data or [[None] * len(names)]
        grid: list[list[object]] = ([list(names)] if heading else []) + [list(row) for row in body]
        height, width = len(grid), len(names)
        origin = self.destination.areas[0]
        old = self.result_area()
        self._make_room(origin, old, height, width)
        start = old if old is not None else origin
        from pyopenvba._a1 import Area

        area = Area(start.top, start.left, start.top + height - 1, start.left + width - 1, self.sheet.name)
        if old is not None and not self.setting('PreserveFormatting', True):
            self._range(old).vba_call('ClearFormats', [])
        for row, values in enumerate(grid, start=area.top):
            for column, value in enumerate(values, start=area.left):
                self._put(row, column, value)
        if heading:
            self._range(Area(area.top, area.left, area.top, area.right, self.sheet.name)).vba_get('Font', []).vba_set('Bold', True)  # type: ignore[attr-defined]
        if old is not None and self.setting('FillAdjacentFormulas', False):
            self._fill_formulas(area, 1 if heading else 0)
        self.fields = names
        self._named(area)
        self.refreshed, self.emptied, self.size = True, False, (height, width)
        self.sheet.book.saved = False

    def _put(self, row: int, column: int, value: object) -> None:
        from pyopenvba._a1 import Area

        cell = self._range(Area(row, column, row, column, self.sheet.name))
        if value is None:
            cell.vba_call('ClearContents', [])
            return
        if isinstance(value, (_dt.datetime, _dt.date)):
            cell.vba_set('NumberFormat', _DATE_FORMAT)
            cell.vba_set('Value', VBADate.from_datetime(value))
            return
        if isinstance(value, Decimal):
            value = float(value)
        elif isinstance(value, int) and not isinstance(value, bool):
            value = float(value)
        # Text goes in as text, as a database's does: "00123" is not a number here.
        cell.vba_set('Value2' if not isinstance(value, str) else 'Formula', value if not isinstance(value, str) else _as_text(value))

    def _make_room(self, origin: Area, old: Area | None, height: int, width: int) -> None:
        from pyopenvba._a1 import Area
        from pyopenvba.apps.excel._editing import shift_cells

        name = self.sheet.name
        if old is None:
            block = Area(origin.top, origin.left, origin.top + height - 1, origin.left + width - 1, name)
            occupied = self._occupied(block)
            if occupied:
                # Whatever is in its way moves across, in its rows, by the table's width, whatever RefreshStyle says.
                band = Area(block.top, min(column for _, column in occupied), block.bottom, block.right, name)
                shift_cells(self._range(band), band, delete=False, vertical=False)
            return
        if width != old.right - old.left + 1:
            raise VBAUnsupportedError('A refresh that changes a query table\'s number of columns is not implemented')
        self._range(old).vba_call('ClearContents', [])
        style = int(to_text(self.setting('RefreshStyle', _INSERT_DELETE)) or 0)
        before, last = old.bottom - old.top + 1, old.top + height - 1
        if style == _INSERT_DELETE and height > before:
            # Room opens at the last row, so what referred to the last row still does.
            band = Area(old.bottom, old.left, old.bottom + height - before - 1, old.right, name)
            shift_cells(self._range(band), band, delete=False, vertical=True)
        elif style == _INSERT_DELETE and height < before:
            band = Area(last + 1, old.left, old.bottom, old.right, name)
            shift_cells(self._range(band), band, delete=True, vertical=True)
        elif style == _INSERT_ROWS and height > before:
            in_way = [row for row, _ in self._occupied(Area(old.bottom + 1, 1, last, 16384, name))]
            if in_way:
                first = min(in_way)
                rows = self.sheet.vba_get('Rows', [f'{first}:{last}'])
                rows.vba_call('Insert', [])  # type: ignore[attr-defined]

    def _fill_formulas(self, area: Area, header: int) -> None:
        """Formulas beside the first row carried down to the last."""
        from pyopenvba._a1 import Area

        first = area.top + header
        column = area.right + 1
        while True:
            cell = self.sheet.cells_.get((first, column))
            if cell is None or not getattr(cell, 'formula', None):
                return
            self._range(Area(first, column, area.bottom, column, self.sheet.name)).vba_call('FillDown', [])
            column += 1

    # -- the file

    def connection_markup(self, number: int, name: str) -> str:
        get = self.setting
        keep = bool(get('SavePassword', False))
        attributes = f' id="{number}" keepAlive="1" name="{escape(name)}" type="{int(KINDS[self.kind])}" refreshedVersion="8"'
        if keep:
            attributes += ' savePassword="1"'
        if get('BackgroundQuery', False):
            attributes += ' background="1"'
        if get('SaveData', True):
            attributes += ' saveData="1"'
        connection = self.connection if keep else without_passwords(self.connection)
        command = escape(self.command_text or '').replace('"', '&quot;')
        return f'<connection{attributes}><dbPr connection="{escape(connection).replace(chr(34), "&quot;")}" command="{command}"/></connection>'

    def fields_markup(self) -> str:
        if not self.fields:
            return ''
        fields = ''.join(f'<queryTableField id="{at}" name="{escape(one).replace(chr(34), "&quot;")}"/>' for at, one in enumerate(self.fields, start=1))
        return (f'<queryTableRefresh preserveSortFilterLayout="0" nextId="{len(self.fields) + 1}"><queryTableFields count="{len(self.fields)}">'
                f'{fields}</queryTableFields></queryTableRefresh>')


def _as_text(value: str) -> str:
    """Text a Formula set leaves as it is: one that would read as a number, a date or a formula gets a quote first."""
    return value if value == '' else "'" + value


def _checked(rows: object) -> list[list[object]]:
    """The function's rows: a header of names and rows as long, of values a cell can hold."""
    if isinstance(rows, (str, bytes)) or not isinstance(rows, (list, tuple)):
        raise error(1004, 'The function set for Refresh returned no rows: give a list of rows, the field names first')
    found: list[list[object]] = []
    for row in rows:  # pyright: ignore[reportUnknownVariableType]
        if isinstance(row, (str, bytes)) or not isinstance(row, (list, tuple)):
            raise error(1004, 'The function set for Refresh returned a row that is not a list')
        found.append(list(row))  # pyright: ignore[reportUnknownArgumentType]
    if not found or not found[0]:
        raise error(1004, 'The function set for Refresh returned no field names')
    width = len(found[0])
    for at, row in enumerate(found):
        if len(row) != width:
            raise error(1004, f'The function set for Refresh returned row {at} with {len(row)} values, where the field names are {width}')
        for value in row:
            if at == 0 and not isinstance(value, str):
                raise error(1004, 'The function set for Refresh returned a field name that is not text')
            if value is not None and not isinstance(value, (str, int, float, bool, Decimal, _dt.date, _dt.datetime)):
                raise error(1004, f'The function set for Refresh returned a {type(value).__name__}, which no cell holds')
    return found


def pending(application: Application) -> list[DatabaseQueryTable]:
    """The tables refreshing in the background, waiting for the macro to end."""
    found = getattr(application, 'pending_refreshes', None)
    if found is None:
        found = []
        setattr(application, 'pending_refreshes', found)
    return found


def finish(application: Application) -> list[Exception]:
    """Every waiting refresh run, as Excel runs them once no macro is, and each failure given back."""
    failures: list[Exception] = []
    waiting = pending(application)
    while waiting:
        table = waiting.pop(0)
        table.pending = False
        if not any(table is one for one in table.sheet.query_tables):
            continue
        try:
            table.refresh()
        except Exception as failed:  # noqa: BLE001 - each failure is reported to the caller, none swallowed
            failures.append(failed)
    return failures


__all__ = ['DatabaseQueryTable', 'finish', 'is_database', 'masked', 'pending', 'without_passwords']

_ = EMPTY
