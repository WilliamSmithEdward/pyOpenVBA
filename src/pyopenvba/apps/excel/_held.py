"""What an object answers when a macro still holds it after it has gone (tests/fixtures/shape_held_deleted.json,
tests/fixtures/held_objects.json). Like a held Range that has lost its cells, it still knows its type, and a member it never had still
raises 438, but a real member raises an error each family has its own rule for, 424 most often. A part leads to the object that decides
through whichever of these it holds, whichever way it was reached."""
from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, ClassVar

from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.shapes._values import Shape

_LINKS = ("range", "owner", "characters", "font", "frame", "format", "effect", "stop", "fill", "whole")
_MESSAGES = {424: "Object required", 1004: "Application-defined or object-defined error", 9: "Subscript out of range"}


def shape_here(sheet: Worksheet, shape: Shape) -> bool:
    """Whether the shape is on its sheet, at the top or within a group, the sheet still in its book. The shapes known are kept until
    the drawing next changes, so reaching a member does not walk them every time."""
    if sheet.deleted:
        return False
    known = sheet.live_shapes
    if known is None or id(shape) not in known:
        known = set[int]()
        waiting = list(sheet.shapes_)
        while waiting:
            one = waiting.pop()
            known.add(id(one))
            waiting.extend(one.children)
        sheet.live_shapes = known
    return id(shape) in known


class HeldPart(Created):
    """Something that can go while a macro holds it, or anything reached from one."""

    #: Members that still answer once it has gone.
    held_free: ClassVar[frozenset[str]] = frozenset()
    #: Whether its own rule goes on applying once its sheet has gone, where most things simply answer nothing.
    held_past_sheet: ClassVar[bool] = False

    def held_gone(self) -> bool:
        for name in _LINKS:
            link = getattr(self, name, None)
            if isinstance(link, HeldPart):
                return link.held_gone()
        return False

    def held_error(self, name: str) -> int:
        """The error a real member gives once it has gone, by its name in lower case; 0 if the member still answers."""
        return 424

    def held_set_error(self, name: str) -> int:
        """The error of setting a real member once it has gone, where that differs from reading it."""
        return self.held_error(name)

    def _raise_if_gone(self, name: str, *, setting: bool = False) -> None:
        lower = name.lower()
        sheet = getattr(self, "sheet", None)
        if not self.held_past_sheet and getattr(sheet, "deleted", False) and Created.vba_member(self, name) is not None:
            # With its sheet, a thing has nothing left to answer with, whichever member it is.
            raise error(424, _MESSAGES[424])
        if lower not in self.held_free and self.vba_member(name) is not None and self.held_gone():
            number = self.held_set_error(lower) if setting else self.held_error(lower)
            if number:
                raise error(number, _MESSAGES.get(number, "Automation error"))

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        self._raise_if_gone(name)
        return super().vba_get(name, args, named)

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        self._raise_if_gone(name, setting=True)
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def vba_call(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        self._raise_if_gone(name)
        return super().vba_call(name, args, named)
