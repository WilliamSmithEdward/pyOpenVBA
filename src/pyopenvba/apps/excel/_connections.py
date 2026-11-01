"""Workbook.Connections: the data connections a file keeps in xl/connections.xml, read as Excel reads them.

Measured in live Excel (scripts/measure_connections.py, tests/fixtures/connections.json) on a workbook holding a
Power Query connection and on a new one, which holds none:

- Connections counts the file's connections and takes one by its number from 1 or its name; anything else is
  error 9.
- A connection reads its name and description, and for an OLE DB connection, the file's type 5, xlConnectionTypeOLEDB
  and an OLEDBConnection. That reads the connection string after OLEDB;, the command and its type, and the
  connection's settings from their attributes, keepAlive as MaintainConnection; ODBCConnection, TextConnection,
  ModelConnection and WorksheetDataConnection are error 1004 there.

The model refreshes, adds and changes no connection: each reports itself unsupported, as does a connection of
another type.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec, VBACollection, member
from pyopenvba.interpreter._values import VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Workbook

PART = 'xl/connections.xml'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
#: XlConnectionType by the type a file gives a connection: OLE DB alone is measured.
TYPES = {'5': 1}


def entries(book: Workbook) -> list[XmlElement]:
    """The file's connection elements, in its order."""
    from pyopenvba.apps.excel._charts_file import xml_tree

    package = book.package
    if package is None or not package.has(PART):
        return []
    return list(xml_tree(package.read(PART)).findall(NS + 'connection'))


class Connections(VBACollection, ExcelObject):
    vba_type_name = 'Connections'

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return [WorkbookConnection(self.book, entry) for entry in entries(self.book)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for item in items:
                assert isinstance(item, WorkbookConnection)
                if item.entry.get('name', '').casefold() == index.casefold():
                    return item
            raise error(9, 'Subscript out of range')
        return super().vba_lookup(index, items)

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

    @member
    def Name(self) -> object:
        return self.entry.get('name', '')

    @member
    def Description(self) -> object:
        return self.entry.get('description', '')

    @member
    def Type(self) -> object:
        return VBAInt(self._type(), 'Long')

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


def _enable_refresh() -> MemberSpec:
    def get(connection: OLEDBConnection) -> object:
        """True: a connection element has no attribute that turns refreshing off."""
        return True

    def put(connection: OLEDBConnection, value: object) -> None:
        raise VBAUnsupportedError("changing a data connection's settings is not implemented")

    return MemberSpec(name='EnableRefresh', kind='property', getter=get, setter=put)


OLEDBConnection.vba_add_member(_enable_refresh())
