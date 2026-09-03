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


def next_folder(fixed_rows: int, taken: set[str]) -> str:
    """The name Access gives a new object's storage folder.

    It is computed, not chosen, and Access will not find an object in a
    folder by any other name: `AllModules(i).Name` fails on a module in
    the wrong one while the VBE still lists and runs it.  Names are
    allocated from a base of `chr(0x30 + <rows in the container that are
    not folders>)`, lowest free first.  `Modules` holds four such rows --
    `PropData`, `PropDataCopy`, `\\x03DirData` and `\\x03DirDataCopy` --
    so its folders start at `4`, which is why the second module in a
    database gets `4` and never `1`.  `Scripts` starts empty, so a
    database's first macro gets `0`.

    Measured across six cases, Access's own allocation each time: the
    blank template ({`0`}) gives `4`, then `5`, `6`, `7` as the project
    grows.  Deleting the last module and adding one gives its name back;
    deleting a *middle* module ({`0`, `5`}) and adding one reuses `4`
    rather than continuing upward, which is what rules out counting.
    """
    code = ord("0") + fixed_rows
    while chr(code) in taken:
        code += 1
    return chr(code)


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
# 0, 4, 5, 6, 7 carries exactly those, in whatever order the entries
# happen to be, and a module that reused a freed folder carries the
# reused name.  Measured on six databases Access wrote.


def dir_data_prefix(name: str) -> bytes:
    """An entry up to its folder number, which is what finding one needs."""
    text = name.encode("utf-16-le")
    return bytes((ENTRY_TAG, len(text) + ENTRY_TRAILER)) + text


def dir_data_entry(name: str, folder: str) -> bytes:
    return dir_data_prefix(name) + int(folder).to_bytes(ENTRY_TRAILER, "little")


def add_to_dir_data(payload: bytes, name: str, folder: str) -> bytes:
    return payload + dir_data_entry(name, folder)


def remove_from_dir_data(payload: bytes, name: str) -> bytes:
    """Drop an entry, the four bytes that belong to it included."""
    prefix = dir_data_prefix(name)
    at = payload.find(prefix)
    if at < 0:
        raise LookupError(f"DirData holds no entry for {name!r}")
    return payload[:at] + payload[at + len(prefix) + ENTRY_TRAILER :]


def rename_dir_data(payload: bytes, old: str, new: str) -> bytes:
    """Rewrite an entry's name, leaving the folder it names alone."""
    prefix = dir_data_prefix(old)
    if prefix not in payload:
        raise LookupError(f"DirData holds no entry for {old!r}")
    return payload.replace(prefix, dir_data_prefix(new))


def dir_data_entries(payload: bytes) -> list[tuple[str, str]]:
    """`(name, folder)` for everything the container lists."""
    out: list[tuple[str, str]] = []
    at = 4
    while at + 2 <= len(payload) and payload[at] == ENTRY_TAG:
        size = payload[at + 1]
        body = payload[at + 2 : at + 2 + size]
        folder = int.from_bytes(body[-ENTRY_TRAILER:], "little")
        out.append((body[:-ENTRY_TRAILER].decode("utf-16-le"), str(folder)))
        at += 2 + size
    return out
