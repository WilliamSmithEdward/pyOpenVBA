"""Collections the model holds nothing in: QueryTables, OLEObjects and NamedSheetViews of a sheet, and XmlMaps,
XmlNamespaces, CustomViews, PublishObjects, SlicerCaches, Charts and the Excel 4 macro sheets of a workbook, as a new
workbook has them (scripts/measure_empty_collections.py, tests/fixtures/empty_collections.json).

Each counts nothing, and asking for an item is the error Excel gives, which differs between them and between a
number and a name. A workbook loaded with any of these in it is refused rather than read as empty: the loader notes
the kinds of part a workbook and each sheet came with (see _io), and a chart sheet or an Excel 4 macro sheet stops
the load itself.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet


@dataclass(frozen=True)
class Kind:
    """One collection: its type name, the errors an item asked for by number and by name raises, and the kinds of
    part (a relationship's type, or a marker the loader notes) that would put something in it."""

    type_name: str
    by_number: int
    by_name: int
    held_by: frozenset[str]


KINDS = {
    'QueryTables': Kind('QueryTables', 9, 9, frozenset({'queryTable'})),
    'OLEObjects': Kind('OLEObjects', 1004, 1004, frozenset({'oleObject', 'package', 'oleObjects', 'control'})),
    'NamedSheetViews': Kind('NamedSheetViewCollection', 438, 438, frozenset({'namedSheetView'})),
    'XmlMaps': Kind('XmlMaps', 9, 9, frozenset({'xmlMaps'})),
    'XmlNamespaces': Kind('XmlNamespaces', 9, 9, frozenset({'xmlMaps'})),
    'CustomViews': Kind('CustomViews', 5, 5, frozenset({'customWorkbookViews'})),
    'PublishObjects': Kind('PublishObjects', 9, 1004, frozenset({'webPublishObjects', 'webPublishItems'})),
    'SlicerCaches': Kind('SlicerCaches', 9, 5, frozenset({'slicerCache', 'timelineCache'})),
    'Charts': Kind('Sheets', 9, 9, frozenset()),
    'Excel4MacroSheets': Kind('Sheets', 9, 9, frozenset()),
    'Excel4IntlMacroSheets': Kind('Sheets', 9, 9, frozenset()),
}


class EmptyCollection(VBACollection, ItemDefault, Created):
    """One of the collections a new workbook holds nothing in."""

    def __init__(self, owner: Workbook | Worksheet, book: Workbook, name: str, held: set[str]) -> None:
        self.owner, self.book, self.name, self.kind = owner, book, name, KINDS[name]
        self.vba_type_name = self.kind.type_name
        found = self.kind.held_by & held
        if found:
            raise VBAUnsupportedError(f"{name} of a workbook that holds {sorted(found)[0]} parts is not implemented")

    def vba_items(self) -> list[object]:
        return []

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = self.kind.by_name if isinstance(index, str) else self.kind.by_number
        raise error(number)

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.book.application


__all__ = ['EmptyCollection', 'KINDS']
