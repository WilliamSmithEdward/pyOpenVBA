"""Creator, which every object of Excel's answers with xlCreatorCode, 1480803660, "XCEL" read as a number
(tests/fixtures/small_members.json, 34 objects)."""
from __future__ import annotations

from typing import Any

from pyopenvba.interpreter._objects import MemberSpec, VBACollection, VBAObject, member, method
from pyopenvba.interpreter._values import MISSING, VBAInt


class Enumerator(VBAObject):
    """What a collection's hidden _NewEnum hands a macro that names it: an object it can do nothing with."""

    vba_type_name = "Unknown"


def _enumerator(self: VBAObject, *arguments: object) -> object:
    return Enumerator()


class Created(VBAObject):
    """What every Excel class shares: Creator, and the type library its members are looked up in."""

    vba_library = "excel"

    def __init_subclass__(cls, **arguments: Any) -> None:
        """A collection's hidden _NewEnum, which For Each uses. Named on a typed variable it answers an enumerator, and
        on an Object error 438; the model does not know how a variable was typed (tests/fixtures/hidden_members.json)."""
        super().__init_subclass__(**arguments)
        if issubclass(cls, VBACollection):
            cls.vba_add_member(MemberSpec(name="_NewEnum", kind="method", getter=_enumerator))

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
