"""Creator, which every object of Excel's answers with xlCreatorCode, 1480803660, "XCEL" read as a number
(tests/fixtures/small_members.json, 34 objects)."""
from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from pyopenvba.apps.excel._item_indexes import item_error, loose_index
from pyopenvba.apps.excel._read_only import let_error, listed_let_error, set_error, set_parent_error
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._inventory import member_exists, type_known
from pyopenvba.interpreter._objects import BINDING, MemberSpec, VBACollection, VBAObject, member, method
from pyopenvba.interpreter._values import MISSING, NULL, VBAArray, VBAInt, error


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
        # A few types leave Creator out of the type library, and Excel answers 438 on them, as for any member they lack:
        # Protection, Page and the conditional-format criteria (tests/fixtures/object_trio.json).
        if type_known(self.vba_type_name, self.vba_library) and not member_exists(
                self.vba_type_name, "Creator", self.vba_library):
            raise error(438, f"{self.vba_type_name} has no Creator")
        return VBAInt(1480803660, "Long")

    def vba_set(self, name: str, value: object, args: Sequence[object] = (), named: dict[str, object] | None = None,
                *, by_ref: bool = False) -> None:
        if by_ref and name.lower() == "parent" and self.vba_member(name) is not None:
            parent = self.vba_get(name)
            owner = parent.vba_type_name if isinstance(parent, VBAObject) else ""
            raise error(set_parent_error(self.vba_type_name, owner), f"{self.vba_type_name}.Parent cannot be set")
        spec = self.vba_member(name) if by_ref else None
        if spec is not None and spec.setter is None and spec.ref_setter is None and spec.getter is not None:
            # A read-only member has no Set to find, and what it answers says which error that is.
            if spec.parameters and (self.vba_type_name, spec.name) == ("Range", "End"):
                raise error(438, "Range.End cannot be assigned to")
            found = self.vba_get(name)
            raise error(set_error(self.vba_type_name, spec.name, found), f"{self.vba_type_name}.{name} cannot be assigned to")
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def vba_let_refused(self, name: str) -> Exception | None:
        listed = listed_let_error(self.vba_type_name, name)
        return None if listed is None else error(listed, f"{self.vba_type_name}.{name} cannot be assigned to")

    def vba_read_only_object(self, name: str, found: object) -> Exception:
        """An object with no value to take: 438, but for Application, which is 450 like a value (tests/fixtures/excel_everyday_14.json)."""
        number = 450 if isinstance(found, VBAObject) and found.vba_type_name == "Application" else 438
        return error(number, f"{self.vba_type_name}.{name} cannot be assigned to")

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() != "item" or not isinstance(self, VBACollection):
            return super().vba_get(name, args, named)
        if len(args) == 1 and not named:
            if self.vba_type_name == "Scenarios" and args[0] is NULL:
                if BINDING.late:
                    return self
                raise error(13)
            # o(Array(1)) on Hyperlinks is the collection, late bound (tests/fixtures/excel_everyday_14.json).
            if self.vba_type_name == "Hyperlinks" and isinstance(args[0], VBAArray):
                return self
            args = [loose_index(self.vba_type_name, args[0])]
        try:
            return super().vba_get(name, args, named)
        except VBARuntimeError as failed:
            index = (args[0] if args else None) if not named else None
            if not args and not named:
                index = MISSING
            number = None if index is None else item_error(self.vba_type_name, index, len(self.vba_items()))
            if number is None:
                raise
            raise error(number, failed.description) from failed

    def vba_read_only(self, name: str) -> Exception:
        """Assigning to a read-only member: a gap where the real class lets a macro set it, and otherwise Excel's error,
        which is not VBA's 438 for most of them (tests/fixtures/read_only_lets.json)."""
        found = super().vba_read_only(name)
        if isinstance(found, VBAUnsupportedError):
            return found
        return error(let_error(self.vba_type_name, name), f"{self.vba_type_name}.{name} cannot be assigned to")


class ItemDefault(VBAObject):
    """A collection's hidden _Default, which finds an item as Item does, errors included
    (tests/fixtures/hidden_defaults.json)."""

    @method(name="_Default")
    def hidden_default(self, Index: object = MISSING) -> object:  # noqa: N803 - native named argument
        assert isinstance(self, VBACollection)
        return self.vba_lookup(Index, self.vba_items())
