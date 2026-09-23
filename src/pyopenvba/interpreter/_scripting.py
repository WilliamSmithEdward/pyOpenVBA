"""Scripting.Dictionary, which CreateObject makes without starting the Scripting Runtime.

Measured in live Excel (tests/fixtures/vba_semantics/probes.txt). Keys
are matched as the real dictionary matches them: numbers by value
whatever their type -- 1, 1# and 1& are one key, and so are a Date and
its serial number, and True and -1 -- while the text "1" is another;
Empty and "" are one key; Null is a key of its own; an object is a key
by its identity; text is matched case and all, or without case under
vbTextCompare, which can only be set while the dictionary is empty.

Reading a key the dictionary has not got adds it, holding Empty; Exists
adds nothing. An item keeps the type it was given. Keys and Items are
zero-based Variant arrays in the order the keys went in, and For Each
walks the keys. Adding a key twice is error 457, removing or renaming
one it has not got 32811, renaming one onto another 457.
"""

from __future__ import annotations

from collections.abc import Hashable, Iterator
from decimal import Decimal

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject, member, method, setter
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBAArray,
    VBAErrorValue,
    VBAInt,
    error,
    to_integer,
    to_number,
    type_name,
)

#: The error for a key the dictionary has not got, which Remove and Key raise.
ERR_NOT_FOUND = 32811


class Dictionary(VBAObject):
    """Scripting.Dictionary."""

    vba_type_name = "Dictionary"
    vba_library = "scripting"

    def __init__(self) -> None:
        #: Each entry under the identity its key is matched by: the key as given, and the item.
        self._entries: dict[Hashable, tuple[object, object]] = {}
        self._mode = 0

    def _identity(self, key: object) -> Hashable:
        """What a key is matched by."""
        if key is EMPTY or key == "":
            return ("text", "")
        if key is NULL:
            return ("null",)
        if isinstance(key, str):
            return ("text", key.lower() if self._mode else key)
        if isinstance(key, VBAObject):
            return ("object", id(key))
        if isinstance(key, (VBAArray, VBAErrorValue)) or key is MISSING:
            raise VBAUnsupportedError(f"a Dictionary key that is {type_name(key)} is not implemented")
        number = to_number(key)
        return ("number", float(number) if not isinstance(number, Decimal) else number)

    @method
    def Add(self, Key: object = MISSING, Item: object = MISSING) -> object:
        if Key is MISSING or Item is MISSING:
            raise error(449)
        identity = self._identity(Key)
        if identity in self._entries:
            raise error(457)
        self._entries[identity] = (Key, Item)
        return EMPTY

    @method
    def Exists(self, Key: object = MISSING) -> object:
        return self._identity(Key) in self._entries

    @member(default=True)
    def Item(self, Key: object = MISSING) -> object:
        identity = self._identity(Key)
        if identity not in self._entries:
            # Reading a key the dictionary has not got adds it, holding Empty.
            self._entries[identity] = (Key, EMPTY)
        return self._entries[identity][1]

    @setter("Item")
    def _set_item(self, Key: object, value: object) -> None:
        identity = self._identity(Key)
        given = self._entries[identity][0] if identity in self._entries else Key
        self._entries[identity] = (given, value)

    @method
    def Items(self) -> object:
        return VBAArray([(0, len(self._entries) - 1)], items=[item for _, item in self._entries.values()])

    @method
    def Keys(self) -> object:
        return VBAArray([(0, len(self._entries) - 1)], items=[key for key, _ in self._entries.values()])

    @method
    def Remove(self, Key: object = MISSING) -> object:
        identity = self._identity(Key)
        if identity not in self._entries:
            raise error(ERR_NOT_FOUND, "Element not found")
        del self._entries[identity]
        return EMPTY

    @method
    def RemoveAll(self) -> object:
        self._entries.clear()
        return EMPTY

    @member
    def Count(self) -> object:
        return VBAInt(len(self._entries), "Long")

    @member
    def CompareMode(self) -> object:
        return VBAInt(self._mode, "Long")

    @setter("CompareMode")
    def _set_compare_mode(self, value: object) -> None:
        if self._entries:
            raise error(5)
        mode = int(to_integer(value, "Long"))
        if mode not in (0, 1):
            raise VBAUnsupportedError("a Dictionary's CompareMode other than vbBinaryCompare or vbTextCompare is "
                                      "not implemented")
        self._mode = mode

    @member
    def Key(self, Key: object = MISSING) -> object:
        raise error(451, "Dictionary.Key can only be assigned to")

    @setter("Key")
    def _set_key(self, Key: object, value: object) -> None:
        """Give an entry a new key, where it stands in the order."""
        identity = self._identity(Key)
        if identity not in self._entries:
            raise error(ERR_NOT_FOUND, "Element not found")
        renamed = self._identity(value)
        if renamed != identity and renamed in self._entries:
            raise error(457)
        self._entries = {(renamed if one == identity else one): ((value, entry[1]) if one == identity else entry)
                         for one, entry in self._entries.items()}

    def vba_iterate(self) -> Iterator[object]:
        yield from [key for key, _ in self._entries.values()]
