"""Creator, which every object of Excel's answers with xlCreatorCode, 1480803660, "XCEL" read as a number
(tests/fixtures/small_members.json, 34 objects)."""
from __future__ import annotations

from pyopenvba.interpreter._objects import VBAObject, member
from pyopenvba.interpreter._values import VBAInt


class Created(VBAObject):
    """What every Excel class shares: Creator."""

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")
