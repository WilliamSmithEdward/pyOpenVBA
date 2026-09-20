"""What a VBA project sees when Word is the host.

Word promotes part of its Application onto the global namespace, which
is why a macro can write ``ActiveDocument`` rather than
``Application.ActiveDocument``.  The type library has a class called
Global whose members are exactly those, and this asks the inventory for
it.
"""

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._constants_data import OFFICE_CONSTANTS, VBA_CONSTANTS
from pyopenvba.interpreter._inventory import members_of
from pyopenvba.interpreter._runtime import UNRESOLVED, HostBridge
from pyopenvba.interpreter._values import VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.word._model import Application


class WordBridge(HostBridge):
    """The Word host, seen from inside a VBA project."""

    def __init__(self, application: Application) -> None:
        self.application = application
        self._global_members = members_of("Global", "word")

    def global_object(self, name: str) -> object:
        if name == "application":
            return self.application
        return UNRESOLVED

    def has_global_member(self, name: str) -> bool:
        return name.lower() in self._global_members

    def call_global(self, name: str, args: list[object], named: dict[str, object]) -> object:
        return self.application.vba_get(name, args, named)

    def set_global(self, name: str, value: object, by_ref: bool) -> None:
        self.application.vba_set(name, value, [], {}, by_ref=by_ref)

    def constant(self, name: str) -> object:
        found = _by_lower_case().get(name.lower())
        return UNRESOLVED if found is None else _as_vba(found)

    def create(self, type_name_: str) -> object:
        raise VBAUnsupportedError(
            f"CreateObject({type_name_!r}) needs a real COM server, which pyOpenVBA does not start"
        )


#: The Word constants a macro uses to place a shape and move about the
#: text.  The generated tables cover VBA and Office; these are Word's.
WORD_CONSTANTS: dict[str, int] = {
    "wdRelativeHorizontalPositionMargin": 0,
    "wdRelativeHorizontalPositionPage": 1,
    "wdRelativeHorizontalPositionColumn": 2,
    "wdRelativeHorizontalPositionCharacter": 3,
    "wdRelativeVerticalPositionMargin": 0,
    "wdRelativeVerticalPositionPage": 1,
    "wdRelativeVerticalPositionParagraph": 2,
    "wdRelativeVerticalPositionLine": 3,
    "wdWrapInline": 7,
    "wdWrapNone": 3,
    "wdWrapSquare": 0,
    "wdWrapThrough": 5,
    "wdWrapTight": 4,
    "wdWrapTopBottom": 1,
    "wdDoNotSaveChanges": 0,
    "wdSaveChanges": -1,
    "wdPromptToSaveChanges": -2,
    "wdFormatDocument": 0,
    "wdFormatXMLDocument": 12,
    "wdFormatXMLDocumentMacroEnabled": 13,
    "wdCharacter": 1,
    "wdWord": 2,
    "wdSentence": 3,
    "wdParagraph": 4,
    "wdStory": 6,
    "wdAlertsNone": 0,
    "wdAlertsAll": -1,
}


@lru_cache(maxsize=1)
def _by_lower_case() -> dict[str, object]:
    out: dict[str, object] = {}
    for table in (VBA_CONSTANTS, OFFICE_CONSTANTS, WORD_CONSTANTS):
        for name, value in table.items():
            out[name.lower()] = value
    return out


def _as_vba(value: object) -> object:
    if isinstance(value, bool):
        return VBAInt(-1 if value else 0, "Integer")
    if isinstance(value, int):
        return VBAInt(value, "Long")
    return value
