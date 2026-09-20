"""What a VBA project sees when PowerPoint is the host.

PowerPoint promotes part of its Application onto the global namespace,
which is why a macro can write ``ActivePresentation`` rather than
``Application.ActivePresentation``.  Which part is not guesswork: the
type library has a class called Global whose members are exactly those,
and this asks the inventory for it.
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
    from pyopenvba.apps.powerpoint._model import Application


class PowerPointBridge(HostBridge):
    """The PowerPoint host, seen from inside a VBA project."""

    def __init__(self, application: Application) -> None:
        self.application = application
        self._global_members = members_of("Global", "powerpoint")

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


#: The PowerPoint constants a macro uses to place and wire a shape.
#: Generated lists cover VBA and Office; these are PowerPoint's own.
POWERPOINT_CONSTANTS: dict[str, int] = {
    "ppLayoutBlank": 12,
    "ppLayoutText": 2,
    "ppLayoutTitleOnly": 11,
    "ppLayoutTitle": 1,
    "ppLayoutTwoColumnText": 3,
    "ppMouseClick": 1,
    "ppMouseOver": 2,
    "ppActionNone": 0,
    "ppActionNextSlide": 1,
    "ppActionPreviousSlide": 2,
    "ppActionFirstSlide": 3,
    "ppActionLastSlide": 4,
    "ppActionRunMacro": 8,
    "ppActionHyperlink": 7,
    "ppSaveAsPresentation": 1,
    "ppSaveAsOpenXMLPresentationMacroEnabled": 25,
    "ppViewNormal": 9,
    "ppViewSlide": 1,
}


@lru_cache(maxsize=1)
def _by_lower_case() -> dict[str, object]:
    out: dict[str, object] = {}
    for table in (VBA_CONSTANTS, OFFICE_CONSTANTS, POWERPOINT_CONSTANTS):
        for name, value in table.items():
            out[name.lower()] = value
    return out


def _as_vba(value: object) -> object:
    if isinstance(value, bool):
        return VBAInt(-1 if value else 0, "Integer")
    if isinstance(value, int):
        return VBAInt(value, "Long")
    return value
