"""Creator, which every object of Excel's answers with xlCreatorCode, 1480803660, "XCEL" read as a number
(tests/fixtures/small_members.json, 34 objects)."""
from __future__ import annotations

from collections.abc import Sequence

from pyopenvba.apps.excel._item_indexes import item_error, loose_index
from pyopenvba.apps.excel._read_only import let_error, set_parent_error
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._inventory import member_exists, type_known
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method
from pyopenvba.interpreter._values import MISSING, NULL, VBAInt, error


class Created(VBAObject):
    """What every Excel class shares: Creator, and the type library its members are looked up in."""

    vba_library = "excel"

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
        super().vba_set(name, value, args, named, by_ref=by_ref)

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        if name.lower() != "item" or not isinstance(self, VBACollection):
            return super().vba_get(name, args, named)
        if len(args) == 1 and not named:
            if self.vba_type_name == "Scenarios" and args[0] is NULL:
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
