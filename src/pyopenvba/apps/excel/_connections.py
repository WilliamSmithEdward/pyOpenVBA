"""Workbook.Connections: the data connections a file keeps in xl/connections.xml, read as Excel reads them.

Measured in live Excel (scripts/measure_connections.py, tests/fixtures/connections.json) on a workbook holding a
Power Query connection and on a new one, which holds none:

- Connections counts the file's connections and takes one by its number from 1 or its name; anything else is
  error 9.
- A connection reads its name and description, and for an OLE DB connection, the file's type 5, xlConnectionTypeOLEDB
  and an OLEDBConnection. That reads the connection string after OLEDB;, the command and its type, and the
  connection's settings from their attributes, keepAlive as MaintainConnection; ODBCConnection, TextConnection,
  ModelConnection and WorksheetDataConnection are error 1004 there.

Connections.Add2 adds an OLE DB or ODBC connection and Delete takes one away, each named and described as asked, with
Excel's rules for a name there already and for a connection string it does not take
(scripts/measure_connection_members.py, tests/fixtures/connection_members.json). A workbook keeps these for the session;
saving one reports itself.

A query table over a text file has a connection of its own, xlConnectionTypeTEXT, which TextFileConnection and
TextConnection answer for (see _query_tables, scripts/measure_query_table_connections.py,
tests/fixtures/query_table_connections.json). The list runs in the order of the names, ignoring case.

The model refreshes no other connection, and it reports itself unsupported, as does a connection of another type.
"""
from __future__ import annotations

from html import escape
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._query_tables import TextLink

PART = 'xl/connections.xml'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
#: XlConnectionType by the type a file gives a connection: OLE DB alone is measured.
TYPES = {'5': 1}


def entries(book: Workbook) -> list[XmlElement]:
    """The file's connection elements, in its order."""
    from pyopenvba.apps.excel._charts_file import xml_tree

    package = book.package
    from pyopenvba.apps.excel._query_tables import links

    # With a baseline every connection in the file is a query table's, which the links stand for.
    read = [] if package is None or not package.has(PART) or book.query_baseline is not None else xml_tree(package.read(PART)).findall(NS + 'connection')
    found = [entry for entry in read if entry.get('name', '') not in book.connections_deleted]
    found += [link.element for link in links(book) if not link.deleted]
    return sorted(found + book.connections_added, key=lambda entry: entry.get('name', '').lower())


def refuse_changed(book: Workbook) -> None:
    """Report a connection a save would have to add or take away."""
    if book.connections_added or book.connections_deleted:
        raise VBAUnsupportedError('saving a data connection a macro added or deleted is not implemented')


class Connections(VBACollection, ExcelObject):
    vba_type_name = 'Connections'

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        from pyopenvba.apps.excel._query_tables import links

        text = {id(link.element): link for link in links(self.book)}
        return [TextFileConnection(self.book, text[id(entry)]) if id(entry) in text else WorkbookConnection(self.book, entry)
                for entry in entries(self.book)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                assert isinstance(item, WorkbookConnection)
                if item.entry.get('name', '').casefold() == index.casefold():
                    return item
            raise error(9, 'Subscript out of range')
        return super().vba_lookup(index, items)

    @method
    def Add2(self, Name: object = MISSING, Description: object = MISSING, ConnectionString: object = MISSING,
             CommandText: object = MISSING, lCmdtype: object = MISSING, CreateModelConnection: object = MISSING,
             ImportRelationships: object = MISSING) -> object:
        """A new connection. An empty name is Connection; one there already, in any case, takes the next number."""
        from pyopenvba.apps.excel._charts_file import xml_tree

        if MISSING in (Name, Description, ConnectionString, CommandText, lCmdtype):
            raise error(450, 'Wrong number of arguments or invalid property assignment')
        if CreateModelConnection is not MISSING or ImportRelationships is not MISSING:
            raise VBAUnsupportedError('Connections.Add2 into the data model is not implemented')
        prefix, _, rest = to_text(ConnectionString).partition(';')
        kind = {'OLEDB': '5', 'ODBC': '1'}.get(prefix.upper())
        command_type = int(to_integer(lCmdtype, 'Long'))
        if kind is None or not 1 <= command_type <= 8:
            raise error(5, 'Invalid procedure call or argument')
        from pyopenvba.apps.excel._query_tables import links

        # A deleted connection's table still holds its name here, though a new table may take it.
        taken = {entry.get('name', '').casefold() for entry in entries(self.book)} | {link.name.casefold() for link in links(self.book)}
        base = to_text(Name) or 'Connection'
        name, number = base, 0
        while name.casefold() in taken:
            number += 1
            name = f'{base}{number}'
        entry = xml_tree(f'<connection xmlns="{NS[1:-1]}" name="{escape(name)}" description="{escape(to_text(Description))}" '
                         f'type="{kind}" background="1" added="1"><dbPr connection="{escape(rest)}" '
                         f'command="{escape(to_text(CommandText))}" commandType="{command_type}"/></connection>')
        self.book.connections_added.append(entry)
        self.book.saved = False
        return WorkbookConnection(self.book, entry)

    @method
    def AddFromFile(self, Filename: object = MISSING, CreateModelConnection: object = MISSING,
                    ImportRelationships: object = MISSING) -> object:
        self.book.application.outside('Connections.AddFromFile')
        if Filename is MISSING:
            raise error(449, 'Argument not optional')
        if Path(to_text(Filename)).is_file():
            raise VBAUnsupportedError('Connections.AddFromFile of a connection file is not implemented')
        raise error(-2147417848, "Method 'AddFromFile' of object 'Connections' failed")

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application


class WorkbookConnection(ExcelObject):
    vba_type_name = 'WorkbookConnection'

    def __init__(self, book: Workbook, entry: XmlElement) -> None:
        self.book, self.entry = book, entry

    def _type(self) -> int:
        found = TYPES.get(self.entry.get('type', ''))
        if found is None:
            raise VBAUnsupportedError('a data connection of this type is not implemented')
        return found

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    def _added(self) -> None:
        if self.entry not in self.book.connections_added:
            raise VBAUnsupportedError("changing a data connection a file holds is not implemented")

    @member(default=True)
    def Name(self) -> object:
        return self.entry.get('name', '')

    @setter('Name')
    def _set_name(self, value: object) -> None:
        self._added()
        self.entry.set('name', to_text(value))

    @member
    def Description(self) -> object:
        return self.entry.get('description', '')

    @setter('Description')
    def _set_description(self, value: object) -> None:
        self._added()
        self.entry.set('description', to_text(value))

    @member
    def Type(self) -> object:
        if self.entry.get('type') == '1':
            return VBAInt(2, 'Long')
        return VBAInt(self._type(), 'Long')

    @method
    def Delete(self) -> object:
        if self.entry in self.book.connections_added:
            self.book.connections_added.remove(self.entry)
        else:
            self.book.connections_deleted.add(self.entry.get('name', ''))
        self.book.saved = False
        return EMPTY

    @member
    def Ranges(self) -> object:
        from pyopenvba.apps.excel._empty_collections import EmptyCollection

        return EmptyCollection(self.book, self.book, 'Ranges', set())

    @member
    def DataFeedConnection(self) -> object:
        raise error(1004)

    @member
    def ModelTables(self) -> object:
        raise error(1004)

    @member
    def OLEDBConnection(self) -> object:
        self._type()
        return OLEDBConnection(self)

    @member
    def RefreshWithRefreshAll(self) -> object:
        self._type()
        if self.entry.find(NS + 'extLst') is not None:
            raise VBAUnsupportedError("a connection's refresh-all setting in its extension is not implemented")
        return True

    @member
    def InModel(self) -> object:
        self._type()
        return False

    def _other(self) -> object:
        self._type()
        raise error(1004)

    @member
    def ODBCConnection(self) -> object:
        return self._other()

    @member
    def TextConnection(self) -> object:
        return self._other()

    @member
    def ModelConnection(self) -> object:
        return self._other()

    @member
    def WorksheetDataConnection(self) -> object:
        return self._other()


#: Excel's automation error, which a connection left behind by its table raises, and a text connection's delimiters.
_AUTOMATION = (-2147417848, 'Automation error')


class TextFileConnection(WorkbookConnection):
    """The connection of a query table over a text file."""

    vba_type_name = 'WorkbookConnection'

    def __init__(self, book: Workbook, link: TextLink) -> None:
        super().__init__(book, link.element)
        self.link = link

    def _check_alive(self) -> None:
        from pyopenvba.apps.excel._query_tables import links

        if self.link not in links(self.book) or self.link.deleted:
            raise error(424, 'Object required')
        super()._check_alive()

    def _type(self) -> int:
        if self.link.left_behind:
            raise error(*_AUTOMATION)
        return 4

    @member(default=True)
    def Name(self) -> object:
        return self.link.name

    @setter('Name')
    def _set_name(self, value: object) -> None:
        if self.link.left_behind:
            raise VBAUnsupportedError('Naming a connection its query table left behind stops Excel, and is refused')
        wanted = to_text(value)
        others = {entry.get('name', '').casefold() for entry in entries(self.book) if entry is not self.entry}
        if not wanted or wanted.casefold() in others:
            raise error(5, 'Invalid procedure call or argument')
        self.entry.set('name', wanted)
        self.book.saved = False

    @member
    def Description(self) -> object:
        return self.link.read_description

    @setter('Description')
    def _set_description(self, value: object) -> None:
        self.link.description = to_text(value)
        self.book.saved = False

    @member
    def RefreshWithRefreshAll(self) -> object:
        return not self.link.excluded

    @setter('RefreshWithRefreshAll')
    def _set_refresh_with_refresh_all(self, value: object) -> None:
        self.link.excluded = not to_bool(value)
        self.book.saved = False

    @member
    def OLEDBConnection(self) -> object:
        return self._other()

    @member
    def TextConnection(self) -> object:
        self._type()
        return TextConnection(self)

    @method
    def Refresh(self) -> object:
        """The table again where it has been refreshed, whatever its BackgroundQuery; nothing where it never has."""
        if self.link.left_behind:
            raise error(-2147418113, 'Automation error')
        if self.link.table.refreshed:
            self.link.table.refresh()
        return EMPTY

    @method
    def Delete(self) -> object:
        if self.link.left_behind:
            self.book.text_links.remove(self.link)
        else:
            self.link.deleted = True
        self.book.saved = False
        return EMPTY


class TextConnection(ExcelObject):
    vba_type_name = 'TextConnection'

    def __init__(self, connection: TextFileConnection) -> None:
        self.connection = connection

    def table(self) -> Any:
        return self.connection.link.table

    @member
    def TextFileHeaderRow(self) -> object:
        raise VBAUnsupportedError('TextConnection.TextFileHeaderRow, which a query table does not have, is not implemented')

    @member
    def Parent(self) -> object:
        return self.connection

    @member
    def Application(self) -> object:
        return self.connection.book.application


def _shared(name: str) -> MemberSpec:
    """A member that is the query table's own."""

    def get(connection: TextConnection) -> object:
        return connection.table().vba_get(name, [])

    def put(connection: TextConnection, value: object) -> None:
        connection.table().vba_set(name, value)

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _delimiter(name: str) -> MemberSpec:
    """Excel reads none of these from here, and setting one stops it."""

    def get(connection: TextConnection) -> object:
        raise error(*_AUTOMATION)

    def put(connection: TextConnection, value: object) -> None:
        raise VBAUnsupportedError('Setting a delimiter on a TextConnection stops Excel, and is refused')

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


for _name in ('Connection', 'TextFileColumnDataTypes', 'TextFileConsecutiveDelimiter', 'TextFileDecimalSeparator', 'TextFileFixedColumnWidths',
              'TextFileOtherDelimiter', 'TextFileParseType', 'TextFilePlatform', 'TextFilePromptOnRefresh', 'TextFileStartRow',
              'TextFileTextQualifier', 'TextFileThousandsSeparator', 'TextFileTrailingMinusNumbers', 'TextFileVisualLayout'):
    TextConnection.vba_add_member(_shared(_name))
for _name in ('TextFileCommaDelimiter', 'TextFileSemicolonDelimiter', 'TextFileSpaceDelimiter', 'TextFileTabDelimiter'):
    TextConnection.vba_add_member(_delimiter(_name))


def _flag(attribute: str, default: bool, *, owner: str = 'connection') -> MemberSpec:
    """A setting the connection or its dbPr keeps as an attribute; a set is not implemented."""

    def get(connection: OLEDBConnection) -> object:
        element = connection.element(owner)
        return element.get(attribute, '1' if default else '0') in ('1', 'true')

    def put(connection: OLEDBConnection, value: object) -> None:
        raise VBAUnsupportedError("changing a data connection's settings is not implemented")

    return MemberSpec(name='', kind='property', getter=get, setter=put)


class OLEDBConnection(ExcelObject):
    vba_type_name = 'OLEDBConnection'

    def __init__(self, connection: WorkbookConnection) -> None:
        self.connection = connection

    def element(self, owner: str) -> XmlElement:
        if owner == 'connection':
            return self.connection.entry
        found = self.connection.entry.find(NS + 'dbPr')
        if found is None:
            raise VBAUnsupportedError('an OLE DB connection with no dbPr is not implemented')
        return found

    @member
    def Parent(self) -> object:
        return self.connection

    @member
    def Application(self) -> object:
        return self.connection.book.application

    @member
    def Connection(self) -> object:
        return 'OLEDB;' + self.element('dbPr').get('connection', '')

    @member
    def CommandText(self) -> object:
        return self.element('dbPr').get('command', '')

    @member
    def CommandType(self) -> object:
        return VBAInt(int(to_integer(self.element('dbPr').get('commandType', '2'), 'Long')), 'Long')

    @member
    def RefreshPeriod(self) -> object:
        return VBAInt(int(self.connection.entry.get('interval', '0')), 'Long')

    @member
    def SourceConnectionFile(self) -> object:
        return self.connection.entry.get('odcFile', '')

    @member
    def ServerCredentialsMethod(self) -> object:
        if self.connection.entry.get('credentials', 'integrated') != 'integrated':
            raise VBAUnsupportedError("a connection's other credential methods are not implemented")
        return VBAInt(0, 'Long')

    @member
    def IsConnected(self) -> object:
        return False

    @member
    def OLAP(self) -> object:
        return False

    @member
    def Refreshing(self) -> object:
        return False

    @member
    def RobustConnect(self) -> object:
        return VBAInt(0, 'Long')

    @member
    def ServerSSOApplicationID(self) -> object:
        return ''

    @member
    def SourceDataFile(self) -> object:
        return ''

    @method
    def CancelRefresh(self) -> object:
        raise error(1004, 'Not currently refreshing.')

    @method
    def Reconnect(self) -> object:
        return EMPTY

    @method
    def MakeConnection(self) -> object:
        raise error(1004)

    @member
    def UseLocalConnection(self) -> object:
        return False

    @member
    def LocalConnection(self) -> object:
        raise error(1004)


for _name, _attribute, _default in (('BackgroundQuery', 'background', False),
                                    ('RefreshOnFileOpen', 'refreshOnLoad', False),
                                    ('SavePassword', 'savePassword', False), ('MaintainConnection', 'keepAlive', False),
                                    ('AlwaysUseConnectionFile', 'onlyUseConnectionFile', False)):
    _spec = _flag(_attribute, _default)
    OLEDBConnection.vba_add_member(MemberSpec(name=_name, kind='property', getter=_spec.getter, setter=_spec.setter))


def _failing(name: str, number: int) -> MemberSpec:
    """A member that fails on a connection that is not to a data cube, or was never refreshed."""
    def get(connection: OLEDBConnection) -> object:
        raise error(number)

    return MemberSpec(name=name, kind='property', getter=get)


for _name, _number in (('LocaleID', 1004), ('MaxDrillthroughRecords', 1004), ('RetrieveInOfficeUILang', 1004),
                       ('RefreshDate', 1004), ('ADOConnection', 1004), ('CalculatedMembers', 1004),
                       ('ServerFillColor', 5), ('ServerFontStyle', 5), ('ServerNumberFormat', 5),
                       ('ServerTextColor', 5)):
    OLEDBConnection.vba_add_member(_failing(_name, _number))


def _enable_refresh() -> MemberSpec:
    def get(connection: OLEDBConnection) -> object:
        """True: a connection element has no attribute that turns refreshing off. A connection a macro added is off."""
        return connection.connection.entry.get('added') != '1'

    def put(connection: OLEDBConnection, value: object) -> None:
        raise VBAUnsupportedError("changing a data connection's settings is not implemented")

    return MemberSpec(name='EnableRefresh', kind='property', getter=get, setter=put)


OLEDBConnection.vba_add_member(_enable_refresh())
