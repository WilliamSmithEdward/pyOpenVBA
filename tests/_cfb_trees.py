"""A compound file's storage trees as the bytes hold them, and Windows' own compound-file code adding to one.

A storage's children form a red-black tree ([MS-CFB] 2.6.4).  Reading never
depends on the colours, but Windows inserts into the tree it reads and
refuses (STG_E_DOCFILECORRUPT) where they do not hold, which is how Excel's
editor came to refuse a new module in a project the writer had saved.
"""
from __future__ import annotations

import ctypes
import struct
from collections.abc import Sequence
from pathlib import Path
from typing import Any

_NOSTREAM = 0xFFFF_FFFF
_BLACK = 1
_STORAGE, _ROOT = 1, 5

#: ``ctypes`` seen as untyped: its Windows-only names (``WinDLL``,
#: ``WINFUNCTYPE``) are not in the stubs off Windows, and the tests are
#: type-checked everywhere.
_windows: Any = ctypes


def directory(raw: bytes) -> list[tuple[str, int, int, int, int, int]]:
    """Each directory entry: its name, object type, colour, left and right sibling, and child."""
    size: int = 1 << struct.unpack_from("<H", raw, 30)[0]
    n_fat: int = struct.unpack_from("<I", raw, 44)[0]
    assert n_fat <= 109, "the header lists every FAT sector of a file this small"
    fat: list[int] = []
    for fat_sector in struct.unpack_from(f"<{n_fat}I", raw, 76):
        fat.extend(struct.unpack_from(f"<{size // 4}I", raw, (fat_sector + 1) * size))
    entries: list[tuple[str, int, int, int, int, int]] = []
    sector: int = struct.unpack_from("<I", raw, 48)[0]
    while sector < 0xFFFF_FFFA:
        for at in range((sector + 1) * size, (sector + 2) * size, 128):
            name_bytes = struct.unpack_from("<H", raw, at + 64)[0]
            name = raw[at:at + max(0, name_bytes - 2)].decode("utf-16-le")
            left, right, child = struct.unpack_from("<III", raw, at + 68)
            entries.append((name, raw[at + 66], raw[at + 67], left, right, child))
        sector = fat[sector]
    return entries


def tree_problems(raw: bytes) -> list[str]:
    """Where a storage's tree breaks [MS-CFB] 2.6.4: its order (length, then upper case), a red
    entry at its top or under another red one, or paths down holding different numbers of black entries."""
    entries = directory(raw)
    problems: list[str] = []
    for storage, kind, _, _, _, top in entries:
        if kind not in (_STORAGE, _ROOT) or top == _NOSTREAM:
            continue
        if entries[top][2] != _BLACK:
            problems.append(f"{storage}: {entries[top][0]!r} at the top is red")
        black_counts: set[int] = set()
        # Each entry with the number of black entries above it, whether its parent is red, and the keys it lies between.
        stack: list[tuple[int, int, bool, tuple[int, str] | None, tuple[int, str] | None]] = [(top, 0, False, None, None)]
        while stack:
            node, blacks, parent_red, low, high = stack.pop()
            if node == _NOSTREAM:
                black_counts.add(blacks)
                continue
            name, _, colour, left, right, _ = entries[node]
            key = (len(name), name.upper())
            if (low is not None and key <= low) or (high is not None and key >= high):
                problems.append(f"{storage}: {name!r} is out of order")
            red = colour != _BLACK
            if red and parent_red:
                problems.append(f"{storage}: red {name!r} under a red entry")
            blacks += 0 if red else 1
            stack.append((left, blacks, red, low, key))
            stack.append((right, blacks, red, key, high))
        if len(black_counts) > 1:
            problems.append(f"{storage}: paths down hold {sorted(black_counts)} black entries")
    return problems


def _method(interface: ctypes.c_void_p, index: int, *argtypes: Any) -> Any:
    """Method ``index`` of a COM interface's table, callable with the interface first."""
    table = ctypes.cast(interface, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    return _windows.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)(table[index])


def windows_adds_modules(path: Path, names: Sequence[str]) -> str:
    """Add each of ``names`` to the file's VBA storage through ole32, as Excel's editor adds a
    module: a stream made as Module2, then renamed.  Every second one is then deleted.  Returns
    the first call Windows refused, with its HRESULT, or an empty string when it took them all."""
    ole32 = _windows.WinDLL("ole32")
    pointer = ctypes.POINTER(ctypes.c_void_p)
    ole32.StgOpenStorage.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p,
                                     ctypes.c_ulong, pointer]
    ole32.StgOpenStorage.restype = ctypes.c_long
    mode = 0x12  # STGM_READWRITE | STGM_SHARE_EXCLUSIVE: written straight through
    root, vba = ctypes.c_void_p(), ctypes.c_void_p()

    def release(interface: ctypes.c_void_p) -> None:
        if interface.value:
            _method(interface, 2)(interface)

    try:
        if hr := ole32.StgOpenStorage(str(path), None, mode, None, 0, ctypes.byref(root)):
            return f"StgOpenStorage: {hr & 0xFFFF_FFFF:#x}"
        open_storage = _method(root, 6, ctypes.c_wchar_p, ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p,
                               ctypes.c_ulong, pointer)
        if hr := open_storage(root, "VBA", None, mode, None, 0, ctypes.byref(vba)):
            return f"OpenStorage VBA: {hr & 0xFFFF_FFFF:#x}"
        create_stream = _method(vba, 3, ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, pointer)
        rename = _method(vba, 13, ctypes.c_wchar_p, ctypes.c_wchar_p)
        destroy = _method(vba, 12, ctypes.c_wchar_p)
        for name in names:
            stream = ctypes.c_void_p()
            if hr := create_stream(vba, "Module2", mode | 0x1000, 0, 0, ctypes.byref(stream)):  # | STGM_CREATE
                return f"CreateStream Module2 for {name}: {hr & 0xFFFF_FFFF:#x}"
            written = ctypes.c_ulong()
            _method(stream, 4, ctypes.c_char_p, ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong))(
                stream, b"x" * 100, 100, ctypes.byref(written))
            release(stream)
            if hr := rename(vba, "Module2", name):
                return f"RenameElement Module2 to {name}: {hr & 0xFFFF_FFFF:#x}"
        for name in names[::2]:
            if hr := destroy(vba, name):
                return f"DestroyElement {name}: {hr & 0xFFFF_FFFF:#x}"
        return ""
    finally:
        release(vba)
        release(root)
