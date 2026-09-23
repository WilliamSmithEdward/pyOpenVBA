"""What a VBA project sees when Excel is the host.

Excel promotes part of its Application onto the global namespace, which
is why a macro can write ``Range("A1")`` rather than
``Application.Range("A1")``.  Which part is not guesswork: the type
library has a class called Global whose members are exactly those, and
this asks the inventory for it rather than keeping a hand-written list
that would drift.
"""

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._constants_data import EXCEL_CONSTANTS, INTEGER_CONSTANTS, OFFICE_CONSTANTS, VBA_CONSTANTS
from pyopenvba.interpreter._inventory import members_of
from pyopenvba.interpreter._runtime import HostBridge, UNRESOLVED
from pyopenvba.interpreter._values import VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application


class ExcelBridge(HostBridge):
    """The Excel host, seen from inside a VBA project."""

    def __init__(self, application: Application) -> None:
        self.application = application
        self._global_members = members_of("Global", "excel")

    def global_object(self, name: str) -> object:
        if name == "application":
            return self.application
        from pyopenvba.apps.excel._model import Workbook

        # A code name, Sheet1 or ThisWorkbook, is the sheet or the workbook of the project's own workbook.
        book = self.application.vba_get("ThisWorkbook")
        if isinstance(book, Workbook):
            if book.code_name.lower() == name:
                return book
            for sheet in book.sheets_:
                if sheet.code_name and sheet.code_name.lower() == name:
                    return sheet
        return UNRESOLVED

    def has_global_member(self, name: str) -> bool:
        return name.lower() in self._global_members

    def call_global(self, name: str, args: list[object], named: dict[str, object]) -> object:
        return self.application.vba_get(name, args, named)

    def set_global(self, name: str, value: object, by_ref: bool) -> None:
        self.application.vba_set(name, value, [], {}, by_ref=by_ref)

    def constant(self, name: str) -> object:
        found = _by_lower_case().get(name.lower())
        return UNRESOLVED if found is None else _as_vba(name.lower(), found)

    def create(self, type_name_: str) -> object:
        raise VBAUnsupportedError(
            f"CreateObject({type_name_!r}) needs a real COM server, which pyOpenVBA does not start"
        )

    def evaluate_bracket(self, text: str) -> object:
        return self.application.evaluate_text(text)


@lru_cache(maxsize=1)
def _by_lower_case() -> dict[str, int | float | str]:
    """Every constant a macro may name, indexed the way VBA matches one."""
    out: dict[str, int | float | str] = {}
    for table in (OFFICE_CONSTANTS, EXCEL_CONSTANTS, VBA_CONSTANTS):
        for key, value in table.items():
            out[key.lower()] = value
    return out


def _as_vba(name: str, value: int | float | str) -> object:
    """A constant as VBA holds it: an enum's member a Long however small, a key code the Integer it is declared."""
    if isinstance(value, bool):  # pragma: no cover - filtered out when generated
        return value
    if isinstance(value, int):
        return VBAInt(value, "Integer" if name in INTEGER_CONSTANTS else "Long")
    return value
