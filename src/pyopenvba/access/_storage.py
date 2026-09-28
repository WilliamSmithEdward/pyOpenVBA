"""`MSysAccessStorage`: the tree Access keeps its objects' data in.

Modules live under `Modules`, macros under `Scripts`, and both are laid
out the same way: a numbered folder per object, the object's bytes in a
row beneath it, and a `\\x03DirData` beside the folders listing what is
there.  This module owns the parts both need.
"""

from __future__ import annotations

import random
import string

STORAGE_TABLE = "MSysAccessStorage"
#: A storage row is either a folder (1) or a value (2).
TYPE_FOLDER = 1
TYPE_VALUE = 2
DIR_DATA = "\x03DirData"
STREAM_NAME_LENGTH = 28
#: Every object's storage folder holds this, unchanging, 13 bytes.
PROP_DATA = bytes.fromhex("00000000020000000000000000")

#: One entry in a `\\x03DirData` payload.
ENTRY_TAG = 4
ENTRY_TRAILER = 4
#: The `<u32 0>` a `\\x03DirData` payload and a folder list open with, so
#: a list this long holds nothing.
LIST_HEADER = 4


def next_folder(taken: set[str]) -> str:
    """The name Access gives a new object's storage folder: the lowest
    number no folder in the container has, counting from `0`, written in
    decimal.

    Measured with Access driven over COM alone.  Modules added to the
    blank template ({`0`}) took `1`, `2`, `3`; deleting the middle one
    and adding another reused `2`; a database that never held code gave
    its first module `0`.  `Forms`, `Reports` and `Scripts` count the
    same way, and past `9` every container Access filled held `10`, `11`
    and `12`, `\\x03DirData` numbering them the same (GitHub issue #34).

    An earlier rule started `Modules` at `4`.  It was measured through
    pyvbaharness, whose three injected modules held `1`, `2` and `3`
    while Access added the module it was asked for.
    """
    number = 0
    while str(number) in taken:
        number += 1
    return str(number)


def stream_row_name(rng: random.Random, taken: set[str]) -> str:
    """A module's storage row name: 28 random capitals, unused."""
    while True:
        name = "".join(rng.choice(string.ascii_uppercase) for _ in range(STREAM_NAME_LENGTH))
        if name not in taken:
            return name


# --- the container's `\x03DirData` -------------------------------------------
# `<u32 0>` and then one entry each:
#
#     04 <u8 payload length> <name UTF-16> <u32 folder>
#
# where the payload length counts the name's bytes plus the four of the
# folder number.  **The trailing four bytes name the object's storage
# folder**, not a terminator: a five-module project whose folders are
# 0, 4, 5, 6, 7 carries exactly those, in the order the next section
# describes, and a module that reused a freed folder carries the reused
# name.  Measured on six databases Access wrote.


# --- the order Access keeps a container's lists in -----------------------------
# Access loads a container's `\x03DirData` into an MSVC `std::unordered_map`
# keyed by the object's name, and the container's `PropData` folder list
# into another keyed by the folder's name, and writes each back in the
# map's own order, only when it changed.  Read from MSACCESS.EXE 16.0
# (the hash at 0x14006c060, the map's insert, grow and rehash) and checked
# against every order Access wrote in the probes: objects added one
# session each and many in one, deleted, renamed, and named past ASCII
# (docs/research/access_write/README.md).

#: The ASCII characters the name hash skips, and those it counts as a space.
_HASH_SKIPPED = (frozenset(range(0x20)) - {0x09}) | {0x22, 0x27, 0x7E, 0x7F}
_HASH_SPACES = frozenset({0x09, 0x20})
#: cp1252 small letters whose capitals differ in more than the case bit.
_CP1252_CAPITALS = {0x9A: 0x8A, 0x9C: 0x8C, 0x9E: 0x8E}
#: A new map's bucket count; a map below 512 buckets grows eightfold.
_FIRST_BUCKETS = 8
_EIGHTFOLD_BELOW = 512


def name_hash(name: str) -> int:
    """The 16-bit hash Access files a name under in a container's lists.

    Each character that counts adds its low five bits, `h = (h << 5) +
    (h >> 13) + 1 + bits` in 16 bits, so the hash ignores case and `0`
    hashes as `P`.  A leading `.` is skipped.  ASCII characters go through
    a table: quotes, `~` and control characters are skipped, and a space
    or tab counts as 0.  From the first character past ASCII, the rest of
    the name is taken a byte at a time in the system code page, upper-cased,
    every byte above 1 counting.  That code page is cp1252 on Western
    Windows, which this assumes.
    """
    if name.startswith("."):
        name = name[1:]
    values: list[int] = []
    for at, character in enumerate(name):
        code = ord(character)
        if code >= 0x80:
            rest = name[at:].encode("cp1252", errors="replace")
            values += [_CP1252_CAPITALS.get(byte, byte) & 0x1F for byte in rest if byte > 1]
            break
        if code not in _HASH_SKIPPED:
            values.append(0 if code in _HASH_SPACES else code & 0x1F)
    hashed = 0
    for value in values:
        hashed = (((hashed << 5) & 0xFFFF) + (hashed >> 13) + 1 + value) & 0xFFFF
    return hashed


class _ContainerMap:
    """Access's map of a container list's keys, in the order it keeps them.

    One list holds every key, each bucket's keys together.  A new key goes
    in front of the first key of its bucket, or at the end when its bucket
    is empty.  An insert that would leave more keys than buckets first
    grows the map, eightfold below 512 buckets, to a power of two, and the
    rehash walks the list moving each key to the front of its new bucket.
    Loading inserts the stored keys in their stored order.
    """

    def __init__(self, stored: list[str]) -> None:
        self.keys: list[str] = []
        self.buckets = _FIRST_BUCKETS
        for key in stored:
            self.insert(key)

    def _bucket(self, key: str) -> int:
        return name_hash(key) & (self.buckets - 1)

    def insert(self, key: str) -> None:
        if len(self.keys) + 1 > self.buckets:
            wanted = len(self.keys) + 1
            if self.buckets < _EIGHTFOLD_BELOW:
                wanted = max(wanted, self.buckets * 8)
            self.buckets = 1 << (wanted - 1).bit_length()
            chains: dict[int, list[str]] = {}
            for moved in self.keys:
                chains.setdefault(self._bucket(moved), []).insert(0, moved)
            self.keys = [moved for chain in chains.values() for moved in chain]
        bucket = self._bucket(key)
        first = next((at for at, other in enumerate(self.keys) if self._bucket(other) == bucket), None)
        if first is None:
            self.keys.append(key)
        else:
            self.keys.insert(first, key)


def access_order(stored: list[str], *, remove: tuple[str, ...] = (), add: tuple[str, ...] = ()) -> list[str]:
    """The keys of a container list after Access loads it as `stored`,
    erases `remove` and inserts `add`, in the order it writes them.  A
    rename is an erase and an insert."""
    table = _ContainerMap(stored)
    for key in remove:
        table.keys.remove(key)
    for key in add:
        table.insert(key)
    return table.keys


def dir_data_prefix(name: str) -> bytes:
    """An entry up to its folder number, which is what finding one needs."""
    text = name.encode("utf-16-le")
    return bytes((ENTRY_TAG, len(text) + ENTRY_TRAILER)) + text


def dir_data_entry(name: str, folder: str) -> bytes:
    return dir_data_prefix(name) + int(folder).to_bytes(ENTRY_TRAILER, "little")


def _dir_data_parts(payload: bytes) -> tuple[list[str], dict[str, bytes], bytes]:
    """The names in stored order, each name's whole entry, and whatever
    follows the last entry."""
    names: list[str] = []
    entries: dict[str, bytes] = {}
    at = LIST_HEADER
    while at + 2 <= len(payload) and payload[at] == ENTRY_TAG:
        end = at + 2 + payload[at + 1]
        name = payload[at + 2 : end - ENTRY_TRAILER].decode("utf-16-le")
        names.append(name)
        entries[name] = payload[at:end]
        at = end
    return names, entries, payload[at:]


def _dir_data_in(payload: bytes, order: list[str], entries: dict[str, bytes], tail: bytes) -> bytes:
    return payload[:LIST_HEADER] + b"".join(entries[name] for name in order) + tail


def add_to_dir_data(payload: bytes, name: str, folder: str) -> bytes:
    """List a new object where Access puts it (see `access_order`)."""
    names, entries, tail = _dir_data_parts(payload)
    entries[name] = dir_data_entry(name, folder)
    return _dir_data_in(payload, access_order(names, add=(name,)), entries, tail)


def remove_from_dir_data(payload: bytes, name: str) -> bytes:
    """Drop an entry, the four bytes that belong to it included, leaving
    the rest in the order Access writes them after the erase."""
    names, entries, tail = _dir_data_parts(payload)
    if name not in entries:
        raise LookupError(f"DirData holds no entry for {name!r}")
    return _dir_data_in(payload, access_order(names, remove=(name,)), entries, tail)


def rename_dir_data(payload: bytes, old: str, new: str) -> bytes:
    """Rename an entry, keeping the folder it names.  Access erases the
    old name and inserts the new one, so the entry moves."""
    names, entries, tail = _dir_data_parts(payload)
    if old not in entries:
        raise LookupError(f"DirData holds no entry for {old!r}")
    entries[new] = dir_data_prefix(new) + entries[old][-ENTRY_TRAILER:]
    return _dir_data_in(payload, access_order(names, remove=(old,), add=(new,)), entries, tail)


def dir_data_entries(payload: bytes) -> list[tuple[str, str]]:
    """`(name, folder)` for everything the container lists."""
    out: list[tuple[str, str]] = []
    at = LIST_HEADER
    while at + 2 <= len(payload) and payload[at] == ENTRY_TAG:
        size = payload[at + 1]
        body = payload[at + 2 : at + 2 + size]
        folder = int.from_bytes(body[-ENTRY_TRAILER:], "little")
        out.append((body[:-ENTRY_TRAILER].decode("utf-16-le"), str(folder)))
        at += 2 + size
    return out
