"""Creator, which every object of Excel's answers with xlCreatorCode, 1480803660, "XCEL" read as a number
(tests/fixtures/small_members.json, 34 objects)."""
from __future__ import annotations

from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method
from pyopenvba.interpreter._values import MISSING, VBAInt


class Created(VBAObject):
    """What every Excel class shares: Creator."""

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")


class ItemDefault(VBAObject):
    """A collection's hidden _Default, which finds an item as Item does, errors included
    (tests/fixtures/hidden_defaults.json)."""

    @method(name="_Default")
    def hidden_default(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        assert isinstance(self, VBACollection)
        return self.vba_lookup(Index, self.vba_items())
