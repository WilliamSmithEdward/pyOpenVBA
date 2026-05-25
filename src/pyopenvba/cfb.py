"""
Compound File Binary (CFB / OLE2) parser.

Implements a subset of [MS-CFB] sufficient to locate and extract
VBA project streams from vbaProject.bin.

Structure overview
------------------
Offset 0       : 512-byte header
  - Magic bytes (8)
  - Sector size (power-of-two exponent, usually 9 -> 512 bytes)
  - Mini-stream cutoff size
  - Number of FAT sectors, root entry start sector, ...
Sectors        : fixed-size blocks that follow the header
  - FAT (File Allocation Table) chains describe how data spans sectors
  - Mini-FAT chains describe how small streams span mini-sectors
  - Directory entries (128 bytes each) live in the directory sector chain
"""

from __future__ import annotations

import struct
from collections.abc import Callable
from dataclasses import dataclass
from typing import BinaryIO

from pyopenvba.exceptions import CFBError

# ---------------------------------------------------------------------------
# Constants from the MS-CFB specification
# ---------------------------------------------------------------------------

_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

_FREESECT: int = 0xFFFF_FFFF   # unused sector
_ENDOFCHAIN: int = 0xFFFF_FFFE  # end of a FAT chain
_FATSECT: int = 0xFFFF_FFFD    # sector is part of the FAT
_DIFSECT: int = 0xFFFF_FFFC    # sector is part of the DIFAT
_NOSTREAM: int = 0xFFFF_FFFF   # no sibling / no child

# Directory entry object types
_OBJTYPE_EMPTY: int = 0
_OBJTYPE_STORAGE: int = 1
_OBJTYPE_STREAM: int = 2
_OBJTYPE_ROOT: int = 5

# Spec field order: sig(8), CLSID(16), minorVer(2), majorVer(2), byteOrder(2),
# sectorShift(2), miniSectorShift(2), reserved(6), then nine DWORDs, then DIFAT(436)
_HEADER_FMT = "<8s16sHHHHH6sIIIIIIIII436s"
_HEADER_SIZE = struct.calcsize(_HEADER_FMT)  # == 512

# name(64), nameLen(2), type(1), color(1), left(4), right(4), child(4),
# CLSID(16), state(4), created(8), modified(8), startSector(4), sizeLow(4), sizeHigh(4)
_DIR_ENTRY_FMT = "<64sHBBIII16sIQQIII"
_DIR_ENTRY_SIZE = struct.calcsize(_DIR_ENTRY_FMT)  # == 128


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class DirEntry:
    name: str
    obj_type: int
    child_id: int
    left_sibling_id: int
    right_sibling_id: int
    start_sector: int
    size: int
    is_mini: bool = False  # True when size < mini-stream cutoff
    # Preserved verbatim for round-trip writes (CLSID, state, timestamps, color, etc.)
    raw: bytes = b""


class CFB:
    """Parsed Compound File Binary."""

    # ------------------------------------------------------------------
    # Constructor: build via from_bytes() / from_file()
    # ------------------------------------------------------------------

    def __init__(self, data: bytes) -> None:
        self._data: bytes = data
        self.sector_size: int = 512
        self.mini_sector_size: int = 64
        self.mini_stream_cutoff: int = 4096
        self._fat: list[int] = []
        self._minifat: list[int] = []
        self._directory: list[DirEntry] = []
        self._mini_stream: bytes = b""
        # Pending overrides for write-back; key = directory index, value = new stream bytes.
        self._stream_overrides: dict[int, bytes] = {}

    # ------------------------------------------------------------------
    # Factory
    # ------------------------------------------------------------------

    @classmethod
    def from_bytes(cls, data: bytes) -> CFB:
        cfb = cls(data)
        cfb._parse()
        return cfb

    @classmethod
    def from_file(cls, fh: BinaryIO) -> CFB:
        return cls.from_bytes(fh.read())

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def list_streams(self) -> list[str]:
        """Return the names of all stream entries (not storages)."""
        return [e.name for e in self._directory if e.obj_type == _OBJTYPE_STREAM]

    def get_stream(self, name: str) -> bytes:
        """Return the raw bytes of a stream by name (case-insensitive)."""
        needle = name.casefold()
        for entry in self._directory:
            if entry.obj_type == _OBJTYPE_STREAM and entry.name.casefold() == needle:
                return self._read_stream(entry)
        raise KeyError(f"Stream not found: {name!r}")

    def get_stream_in_storage(self, storage: str, name: str) -> bytes:
        """Return the raw bytes of a stream nested inside a named storage."""
        needle_s = storage.casefold()
        needle_n = name.casefold()
        # Find the storage entry
        storage_entry: DirEntry | None = None
        for entry in self._directory:
            if entry.obj_type == _OBJTYPE_STORAGE and entry.name.casefold() == needle_s:
                storage_entry = entry
                break
        if storage_entry is None:
            raise KeyError(f"Storage not found: {storage!r}")
        # Walk children of that storage
        for entry in self._directory:
            if entry.obj_type == _OBJTYPE_STREAM and entry.name.casefold() == needle_n:
                return self._read_stream(entry)
        raise KeyError(f"Stream {name!r} not found in storage {storage!r}")

    def list_storages(self) -> list[str]:
        """Return the names of all storage entries (excluding the root)."""
        return [
            e.name for e in self._directory
            if e.obj_type == _OBJTYPE_STORAGE
        ]

    def list_streams_in_storage(self, storage: str) -> list[str]:
        """
        Return the names of all streams whose case-folded parent storage matches
        ``storage``.

        .. note::
            The current scan is linear over the directory and does not walk the
            red-black sibling tree.  For Excel-produced ``vbaProject.bin`` this
            is sufficient because storage names are unique.
        """
        needle = storage.casefold()
        found_storage = False
        for e in self._directory:
            if e.obj_type == _OBJTYPE_STORAGE and e.name.casefold() == needle:
                found_storage = True
                break
        if not found_storage:
            raise KeyError(f"Storage not found: {storage!r}")
        return [
            e.name for e in self._directory
            if e.obj_type == _OBJTYPE_STREAM
        ]

    # ------------------------------------------------------------------
    # Internal parsing
    # ------------------------------------------------------------------

    def _parse(self) -> None:
        data = self._data
        if len(data) < _HEADER_SIZE:
            raise CFBError("File too small to be a valid CFB.")
        if data[:8] != _MAGIC:
            raise CFBError("Not a Compound File Binary (magic bytes mismatch).")

        (
            _magic,
            _clsid,
            _minor_ver,
            major_ver,
            _byte_order,
            sector_size_pow,
            mini_sector_size_pow,
            _reserved,
            _num_dir_sectors,
            _num_fat_sectors,
            root_dir_start,
            _sig,
            mini_cutoff,
            minifat_start,
            _num_minifat_sectors,
            difat_start,
            num_difat_sectors,
            difat_array_raw,
        ) = struct.unpack_from(_HEADER_FMT, data, 0)

        if major_ver not in (3, 4):
            raise CFBError(f"Unsupported CFB major version: {major_ver}")

        self.sector_size = 1 << sector_size_pow          # 512 or 4096
        self.mini_sector_size = 1 << mini_sector_size_pow  # 64
        self.mini_stream_cutoff = mini_cutoff             # usually 4096

        # Parse the initial DIFAT array from the header (up to 109 entries)
        difat: list[int] = [int(x) for x in struct.unpack_from("<109I", difat_array_raw)]

        # Follow DIFAT chain for large files (> 109 FAT sectors)
        if difat_start != _ENDOFCHAIN and num_difat_sectors > 0:
            sector = difat_start
            for _ in range(num_difat_sectors):
                if sector in (_ENDOFCHAIN, _FREESECT):
                    break
                sector_data = self._sector(sector)
                entries_per_difat = (self.sector_size // 4) - 1
                extra = [int(x) for x in struct.unpack_from(f"<{entries_per_difat}I", sector_data)]
                difat.extend(extra)
                sector = struct.unpack_from("<I", sector_data, self.sector_size - 4)[0]

        # Build FAT from the sectors listed in the DIFAT
        fat_raw = bytearray()
        for sect in difat:
            if sect in (_FREESECT, _ENDOFCHAIN, _FATSECT, _DIFSECT):
                break
            fat_raw.extend(self._sector(sect))
        fat_count = len(fat_raw) // 4
        self._fat = [int(x) for x in struct.unpack_from(f"<{fat_count}I", fat_raw)]

        # Build mini-FAT
        if minifat_start != _ENDOFCHAIN:
            minifat_raw = bytearray()
            for sect in self._chain(minifat_start):
                minifat_raw.extend(self._sector(sect))
            mf_count = len(minifat_raw) // 4
            self._minifat = [int(x) for x in struct.unpack_from(f"<{mf_count}I", minifat_raw)]

        # Parse directory entries
        dir_raw = bytearray()
        for sect in self._chain(root_dir_start):
            dir_raw.extend(self._sector(sect))
        entries_count = len(dir_raw) // _DIR_ENTRY_SIZE
        self._directory = []
        for i in range(entries_count):
            self._directory.append(self._parse_dir_entry(dir_raw, i))

        # Build mini-stream from the root entry's stream data
        if self._directory:
            root = self._directory[0]
            if root.start_sector not in (_ENDOFCHAIN, _FREESECT):
                ms_raw = bytearray()
                for sect in self._chain(root.start_sector):
                    ms_raw.extend(self._sector(sect))
                self._mini_stream = bytes(ms_raw)

    def _sector(self, index: int) -> bytes:
        offset = _HEADER_SIZE + index * self.sector_size
        return self._data[offset: offset + self.sector_size]

    def _chain(self, start: int) -> list[int]:
        sectors: list[int] = []
        current = start
        seen: set[int] = set()
        while current not in (_ENDOFCHAIN, _FREESECT):
            if current in seen or current >= len(self._fat):
                raise CFBError(f"Cycle or out-of-range sector in FAT chain at {current}.")
            seen.add(current)
            sectors.append(current)
            current = self._fat[current]
        return sectors

    def _mini_chain(self, start: int) -> list[int]:
        sectors: list[int] = []
        current = start
        seen: set[int] = set()
        while current not in (_ENDOFCHAIN, _FREESECT):
            if current in seen or current >= len(self._minifat):
                raise CFBError(f"Cycle or out-of-range sector in mini-FAT chain at {current}.")
            seen.add(current)
            sectors.append(current)
            current = self._minifat[current]
        return sectors

    def _parse_dir_entry(self, raw: bytes | bytearray, index: int) -> DirEntry:
        offset = index * _DIR_ENTRY_SIZE
        entry_raw = bytes(raw[offset: offset + _DIR_ENTRY_SIZE])
        (
            name_raw,
            name_len,
            obj_type,
            _color,
            left_id,
            right_id,
            child_id,
            _clsid,
            _state,
            _created,
            _modified,
            start_sector,
            size,
            _size_high,
        ) = struct.unpack_from(_DIR_ENTRY_FMT, raw, offset)

        # Name is UTF-16LE; name_len includes the null terminator
        name_bytes = name_raw[: max(0, name_len - 2)]
        name = name_bytes.decode("utf-16-le", errors="replace")

        return DirEntry(
            name=name,
            obj_type=obj_type,
            child_id=child_id,
            left_sibling_id=left_id,
            right_sibling_id=right_id,
            start_sector=start_sector,
            size=size,
            raw=entry_raw,
        )

    def _read_stream(self, entry: DirEntry) -> bytes:
        # Honor pending write-back overrides.
        for idx, e in enumerate(self._directory):
            if e is entry and idx in self._stream_overrides:
                return self._stream_overrides[idx]
        if entry.size < self.mini_stream_cutoff and entry.obj_type != _OBJTYPE_ROOT:
            # Read from mini-stream
            raw = bytearray()
            for sect in self._mini_chain(entry.start_sector):
                offset = sect * self.mini_sector_size
                raw.extend(self._mini_stream[offset: offset + self.mini_sector_size])
            return bytes(raw[: entry.size])
        # Read from regular stream
        raw = bytearray()
        for sect in self._chain(entry.start_sector):
            raw.extend(self._sector(sect))
        return bytes(raw[: entry.size])

    # ------------------------------------------------------------------
    # Mutation API
    # ------------------------------------------------------------------

    def write_stream(self, name: str, data: bytes) -> None:
        """
        Replace the bytes of an existing top-level stream by name.

        The change is held in memory until :meth:`to_bytes` is called.
        """
        needle = name.casefold()
        for idx, entry in enumerate(self._directory):
            if entry.obj_type == _OBJTYPE_STREAM and entry.name.casefold() == needle:
                self._stream_overrides[idx] = bytes(data)
                return
        raise KeyError(f"Stream not found: {name!r}")

    def write_stream_in_storage(self, storage: str, name: str, data: bytes) -> None:
        """
        Replace the bytes of an existing stream nested inside a storage.

        Currently scans the directory linearly; the first matching stream wins.
        """
        needle_s = storage.casefold()
        needle_n = name.casefold()
        if not any(
            e.obj_type == _OBJTYPE_STORAGE and e.name.casefold() == needle_s
            for e in self._directory
        ):
            raise KeyError(f"Storage not found: {storage!r}")
        for idx, entry in enumerate(self._directory):
            if entry.obj_type == _OBJTYPE_STREAM and entry.name.casefold() == needle_n:
                self._stream_overrides[idx] = bytes(data)
                return
        raise KeyError(f"Stream {name!r} not found in storage {storage!r}")

    def add_stream_to_storage(self, storage: str, name: str, data: bytes) -> None:
        """
        Create a new stream as a child of ``storage`` and seed it with ``data``.

        A new :class:`DirEntry` is appended to the directory (re-using an
        existing empty slot when one is available) and inserted into the
        storage's child subtree.  Raises ``KeyError`` if the storage does not
        exist and ``ValueError`` if a child stream of that name already exists.
        """
        if not name:
            raise ValueError("stream name must be non-empty")
        parent_idx = self._find_storage_index(storage)
        if self._find_child_stream_index(parent_idx, name) is not None:
            raise ValueError(
                f"Stream {name!r} already exists in storage {storage!r}"
            )
        # Prefer to recycle an existing EMPTY slot so the directory does not
        # grow unnecessarily; otherwise append a fresh entry.
        target_idx: int | None = None
        for i, e in enumerate(self._directory):
            if e.obj_type == _OBJTYPE_EMPTY:
                target_idx = i
                break
        if target_idx is None:
            self._directory.append(DirEntry(
                name=name,
                obj_type=_OBJTYPE_STREAM,
                child_id=_NOSTREAM,
                left_sibling_id=_NOSTREAM,
                right_sibling_id=_NOSTREAM,
                start_sector=_ENDOFCHAIN,
                size=0,
                raw=b"",
            ))
            target_idx = len(self._directory) - 1
        else:
            slot = self._directory[target_idx]
            slot.name = name
            slot.obj_type = _OBJTYPE_STREAM
            slot.child_id = _NOSTREAM
            slot.left_sibling_id = _NOSTREAM
            slot.right_sibling_id = _NOSTREAM
            slot.start_sector = _ENDOFCHAIN
            slot.size = 0
            slot.raw = b""
        self._stream_overrides[target_idx] = bytes(data)
        # Splice the new entry into the parent's child subtree and rebalance.
        siblings = self._collect_subtree(self._directory[parent_idx].child_id)
        siblings.append(target_idx)
        self._directory[parent_idx].child_id = self._rebuild_balanced_subtree(siblings)

    def add_substorage(self, parent: str, name: str) -> None:
        """
        Create a new storage as a child of ``parent``.

        Use :meth:`add_stream_to_storage` afterwards to place child streams
        inside the new sub-storage.  Raises ``KeyError`` if ``parent`` does
        not exist and ``ValueError`` if a sibling storage already uses ``name``.
        """
        if not name:
            raise ValueError("storage name must be non-empty")
        parent_idx = self._find_storage_index(parent)
        # Refuse duplicate sibling storages (case-insensitive).
        needle = name.casefold()
        for child_idx in self._collect_subtree(self._directory[parent_idx].child_id):
            sib = self._directory[child_idx]
            if sib.obj_type == _OBJTYPE_STORAGE and sib.name.casefold() == needle:
                raise ValueError(
                    f"Storage {name!r} already exists under {parent!r}"
                )
        target_idx: int | None = None
        for i, e in enumerate(self._directory):
            if e.obj_type == _OBJTYPE_EMPTY:
                target_idx = i
                break
        if target_idx is None:
            self._directory.append(DirEntry(
                name=name,
                obj_type=_OBJTYPE_STORAGE,
                child_id=_NOSTREAM,
                left_sibling_id=_NOSTREAM,
                right_sibling_id=_NOSTREAM,
                start_sector=_ENDOFCHAIN,
                size=0,
                raw=b"",
            ))
            target_idx = len(self._directory) - 1
        else:
            slot = self._directory[target_idx]
            slot.name = name
            slot.obj_type = _OBJTYPE_STORAGE
            slot.child_id = _NOSTREAM
            slot.left_sibling_id = _NOSTREAM
            slot.right_sibling_id = _NOSTREAM
            slot.start_sector = _ENDOFCHAIN
            slot.size = 0
            slot.raw = b""
        siblings = self._collect_subtree(self._directory[parent_idx].child_id)
        siblings.append(target_idx)
        self._directory[parent_idx].child_id = self._rebuild_balanced_subtree(siblings)

    # ------------------------------------------------------------------
    # Directory mutation (stream removal / drop)
    # ------------------------------------------------------------------

    def remove_stream_in_storage(self, storage: str, name: str) -> None:
        """
        Remove a stream from a named storage.

        The stream's directory entry is marked empty and unlinked from the
        storage's red-black sibling tree.  The change is held in memory until
        :meth:`to_bytes` is called.
        """
        parent_idx = self._find_storage_index(storage)
        target_idx = self._find_child_stream_index(parent_idx, name)
        if target_idx is None:
            raise KeyError(f"Stream {name!r} not found in storage {storage!r}")
        self._unlink_and_clear(parent_idx, target_idx)

    def remove_stream(self, name: str) -> None:
        """Remove a top-level stream (child of the root entry)."""
        target_idx = self._find_child_stream_index(0, name)
        if target_idx is None:
            raise KeyError(f"Stream not found: {name!r}")
        self._unlink_and_clear(0, target_idx)

    def rename_stream_in_storage(self, storage: str, old: str, new: str) -> None:
        """
        Rename a stream within a storage.

        The entry's name is replaced and the storage's child subtree is rebuilt
        to keep BST ordering ([MS-CFB] 2.6.4) valid.  Raises ``KeyError`` if
        the source stream is missing and ``ValueError`` if a sibling already
        uses ``new``.
        """
        if not new:
            raise ValueError("new stream name must be non-empty")
        if old.casefold() == new.casefold():
            # Same name (case-insensitive): just patch the displayed casing.
            parent_idx = self._find_storage_index(storage)
            target_idx = self._find_child_stream_index(parent_idx, old)
            if target_idx is None:
                raise KeyError(f"Stream {old!r} not found in storage {storage!r}")
            self._directory[target_idx].name = new
            return
        parent_idx = self._find_storage_index(storage)
        target_idx = self._find_child_stream_index(parent_idx, old)
        if target_idx is None:
            raise KeyError(f"Stream {old!r} not found in storage {storage!r}")
        if self._find_child_stream_index(parent_idx, new) is not None:
            raise ValueError(
                f"Stream {new!r} already exists in storage {storage!r}"
            )
        self._directory[target_idx].name = new
        # Rebuild the parent's subtree so BST ordering reflects the new name.
        siblings = self._collect_subtree(self._directory[parent_idx].child_id)
        self._directory[parent_idx].child_id = self._rebuild_balanced_subtree(siblings)

    def drop_streams_in_storage(
        self, storage: str, predicate: Callable[[str], bool]
    ) -> list[str]:
        """
        Remove every stream child of ``storage`` whose name satisfies ``predicate``.

        Returns the names of removed streams (in directory order).  Raises
        :class:`KeyError` if the storage does not exist.
        """
        parent_idx = self._find_storage_index(storage)
        children = self._collect_subtree(self._directory[parent_idx].child_id)
        removed: list[str] = []
        for child_idx in children:
            entry = self._directory[child_idx]
            if entry.obj_type == _OBJTYPE_STREAM and predicate(entry.name):
                removed.append(entry.name)
        for name in removed:
            self.remove_stream_in_storage(storage, name)
        return removed

    # ------------------------------------------------------------------
    # Directory tree helpers
    # ------------------------------------------------------------------

    def _find_storage_index(self, storage: str) -> int:
        needle = storage.casefold()
        for idx, entry in enumerate(self._directory):
            if entry.obj_type == _OBJTYPE_STORAGE and entry.name.casefold() == needle:
                return idx
        raise KeyError(f"Storage not found: {storage!r}")

    def _find_child_stream_index(self, parent_idx: int, name: str) -> int | None:
        needle = name.casefold()
        for child_idx in self._collect_subtree(self._directory[parent_idx].child_id):
            entry = self._directory[child_idx]
            if entry.obj_type == _OBJTYPE_STREAM and entry.name.casefold() == needle:
                return child_idx
        return None

    def _collect_subtree(self, root_id: int) -> list[int]:
        """Walk the sibling subtree rooted at ``root_id`` and return all node indices."""
        out: list[int] = []
        seen: set[int] = set()
        stack: list[int] = []
        if root_id != _NOSTREAM and root_id < len(self._directory):
            stack.append(root_id)
        while stack:
            node = stack.pop()
            if node in seen or node == _NOSTREAM or node >= len(self._directory):
                continue
            seen.add(node)
            out.append(node)
            entry = self._directory[node]
            if entry.left_sibling_id != _NOSTREAM:
                stack.append(entry.left_sibling_id)
            if entry.right_sibling_id != _NOSTREAM:
                stack.append(entry.right_sibling_id)
        return out

    def _rebuild_balanced_subtree(self, indices: list[int]) -> int:
        """
        Build a balanced BST from ``indices`` sorted by ``([MS-CFB] 2.6.4)``
        ordering: ``(len(name), upper(name))``.  Returns the new root id, or
        ``_NOSTREAM`` if empty.  Side-effect: rewrites the participating
        entries' left/right sibling pointers.
        """
        if not indices:
            return _NOSTREAM
        ordered = sorted(
            indices,
            key=lambda i: (len(self._directory[i].name), self._directory[i].name.upper()),
        )
        return self._build_from_sorted(ordered)

    def _build_from_sorted(self, ordered: list[int]) -> int:
        if not ordered:
            return _NOSTREAM
        mid = len(ordered) // 2
        root = ordered[mid]
        self._directory[root].left_sibling_id = self._build_from_sorted(ordered[:mid])
        self._directory[root].right_sibling_id = self._build_from_sorted(ordered[mid + 1:])
        return root

    def _unlink_and_clear(self, parent_idx: int, target_idx: int) -> None:
        parent = self._directory[parent_idx]
        siblings = [i for i in self._collect_subtree(parent.child_id) if i != target_idx]
        parent.child_id = self._rebuild_balanced_subtree(siblings)
        # Clear the removed entry so the serializer emits an empty slot.
        dead = self._directory[target_idx]
        dead.name = ""
        dead.obj_type = _OBJTYPE_EMPTY
        dead.left_sibling_id = _NOSTREAM
        dead.right_sibling_id = _NOSTREAM
        dead.child_id = _NOSTREAM
        dead.start_sector = _FREESECT
        dead.size = 0
        dead.raw = b""
        # Drop any pending override for the removed stream.
        self._stream_overrides.pop(target_idx, None)

    # ------------------------------------------------------------------
    # Serializer ([MS-CFB] write-path)
    # ------------------------------------------------------------------

    def to_bytes(self) -> bytes:
        """
        Serialize this CFB back to bytes, honoring any pending stream overrides.

        Output uses major version 3 (512-byte sectors, 64-byte mini-sectors,
        4096-byte mini-stream cutoff).  Preserves the original directory-tree
        topology and per-entry metadata (CLSID, state, color, timestamps).

        Layout, in sector order:
            [mini-stream data] [regular stream data] [directory] [mini-FAT] [FAT]

        Files large enough to require a DIFAT chain (more than 109 FAT sectors,
        about 27 MB of total payload) are rejected — vbaProject.bin is always
        well under this limit.
        """
        SECTOR = 512
        MINI = 64
        CUTOFF = 4096
        ENTRIES_PER_SECTOR = SECTOR // 4              # 128 FAT entries / sector
        DIR_ENTRIES_PER_SECTOR = SECTOR // _DIR_ENTRY_SIZE  # 4 dir entries / sector

        # ---- 1. Snapshot every directory entry's stream data ----
        stream_bytes: list[bytes] = []
        for i, entry in enumerate(self._directory):
            if entry.obj_type == _OBJTYPE_STREAM:
                if i in self._stream_overrides:
                    stream_bytes.append(self._stream_overrides[i])
                else:
                    stream_bytes.append(self._read_stream(entry))
            else:
                stream_bytes.append(b"")

        # ---- 2. Pack mini-eligible streams into the mini-stream + mini-FAT ----
        minifat: list[int] = []
        mini_stream = bytearray()
        mini_starts: dict[int, int] = {}   # dir_idx -> first mini-sector
        for i, entry in enumerate(self._directory):
            if entry.obj_type != _OBJTYPE_STREAM:
                continue
            data = stream_bytes[i]
            if len(data) == 0 or len(data) >= CUTOFF:
                continue
            n = (len(data) + MINI - 1) // MINI
            first = len(minifat)
            for k in range(n):
                minifat.append(first + k + 1 if k + 1 < n else _ENDOFCHAIN)
            mini_stream.extend(data)
            if len(data) % MINI != 0:
                mini_stream.extend(b"\x00" * (MINI - len(data) % MINI))
            mini_starts[i] = first

        mini_stream_used_bytes = len(minifat) * MINI  # logical length for Root.size

        # Pad mini-stream to a 512-byte sector boundary.
        if len(mini_stream) % SECTOR != 0:
            mini_stream.extend(b"\x00" * (SECTOR - len(mini_stream) % SECTOR))

        # Pad mini-FAT array to a full sector.
        while len(minifat) % ENTRIES_PER_SECTOR != 0:
            minifat.append(_FREESECT)

        # ---- 3. Assign sector indices ----
        sector_payloads: list[bytes] = []

        # 3a. Mini-stream sectors
        n_mini_stream_sectors = len(mini_stream) // SECTOR
        if n_mini_stream_sectors > 0:
            mini_stream_first = len(sector_payloads)
            for k in range(n_mini_stream_sectors):
                sector_payloads.append(bytes(mini_stream[k * SECTOR:(k + 1) * SECTOR]))
        else:
            mini_stream_first = _ENDOFCHAIN

        # 3b. Regular stream sectors
        reg_starts: dict[int, int] = {}
        for i, entry in enumerate(self._directory):
            if entry.obj_type != _OBJTYPE_STREAM:
                continue
            data = stream_bytes[i]
            if len(data) < CUTOFF:
                continue
            n = (len(data) + SECTOR - 1) // SECTOR
            first = len(sector_payloads)
            padded = data + b"\x00" * (n * SECTOR - len(data))
            for k in range(n):
                sector_payloads.append(padded[k * SECTOR:(k + 1) * SECTOR])
            reg_starts[i] = first

        # 3c. Directory sectors
        n_dir_entries = len(self._directory)
        n_dir_sectors = max(
            1,
            (n_dir_entries + DIR_ENTRIES_PER_SECTOR - 1) // DIR_ENTRIES_PER_SECTOR,
        )
        dir_buf = bytearray()
        for i, entry in enumerate(self._directory):
            dir_buf.extend(self._build_dir_entry_bytes(
                entry=entry,
                data_size=len(stream_bytes[i]),
                mini_starts=mini_starts,
                reg_starts=reg_starts,
                mini_stream_first=mini_stream_first,
                mini_stream_used_bytes=mini_stream_used_bytes,
                idx=i,
            ))
        # Pad remaining entries in the last dir sector with empty entries.
        empty_entry = self._empty_dir_entry_bytes()
        while len(dir_buf) < n_dir_sectors * SECTOR:
            dir_buf.extend(empty_entry)

        dir_first = len(sector_payloads)
        for k in range(n_dir_sectors):
            sector_payloads.append(bytes(dir_buf[k * SECTOR:(k + 1) * SECTOR]))

        # 3d. Mini-FAT sectors
        if minifat:
            minifat_first = len(sector_payloads)
            minifat_bytes = struct.pack(f"<{len(minifat)}I", *minifat)
            n_minifat_sectors = len(minifat_bytes) // SECTOR
            for k in range(n_minifat_sectors):
                sector_payloads.append(minifat_bytes[k * SECTOR:(k + 1) * SECTOR])
        else:
            minifat_first = _ENDOFCHAIN
            n_minifat_sectors = 0

        # ---- 4. Decide how many FAT sectors we need (fixed-point) ----
        n_data = len(sector_payloads)
        n_fat = 1
        while True:
            needed = ((n_data + n_fat) + ENTRIES_PER_SECTOR - 1) // ENTRIES_PER_SECTOR
            if needed <= n_fat:
                break
            n_fat = needed
        if n_fat > 109:
            raise CFBError(
                "CFB writer requires DIFAT chain support for files this large; "
                "not implemented."
            )

        fat_first = n_data
        fat_sector_ids = list(range(fat_first, fat_first + n_fat))

        # ---- 5. Build the FAT ----
        fat = [_FREESECT] * (n_fat * ENTRIES_PER_SECTOR)

        def chain(start: int, n: int) -> None:
            for k in range(n):
                fat[start + k] = (start + k + 1) if k + 1 < n else _ENDOFCHAIN

        if n_mini_stream_sectors > 0:
            chain(mini_stream_first, n_mini_stream_sectors)
        for i, entry in enumerate(self._directory):
            if entry.obj_type == _OBJTYPE_STREAM:
                data = stream_bytes[i]
                if len(data) >= CUTOFF:
                    n_s = (len(data) + SECTOR - 1) // SECTOR
                    chain(reg_starts[i], n_s)
        chain(dir_first, n_dir_sectors)
        if n_minifat_sectors > 0:
            chain(minifat_first, n_minifat_sectors)
        for s in fat_sector_ids:
            fat[s] = _FATSECT

        fat_bytes = struct.pack(f"<{len(fat)}I", *fat)
        for k in range(n_fat):
            sector_payloads.append(fat_bytes[k * SECTOR:(k + 1) * SECTOR])

        # ---- 6. Build the header ----
        difat = list(fat_sector_ids) + [_FREESECT] * (109 - len(fat_sector_ids))
        difat_bytes = struct.pack("<109I", *difat)

        header = struct.pack(
            _HEADER_FMT,
            _MAGIC,
            b"\x00" * 16,         # CLSID — must be zero per spec
            0x003E,               # minor version
            3,                    # major version
            0xFFFE,               # byte-order mark (little-endian)
            9,                    # sector shift: 2^9 = 512
            6,                    # mini-sector shift: 2^6 = 64
            b"\x00" * 6,          # reserved
            0,                    # num_dir_sectors (must be 0 for v3)
            n_fat,                # num_fat_sectors
            dir_first,            # first_dir_sector
            0,                    # transaction signature
            CUTOFF,               # mini-stream cutoff size
            minifat_first,        # first mini-FAT sector
            n_minifat_sectors,    # num mini-FAT sectors
            _ENDOFCHAIN,          # first DIFAT sector (none)
            0,                    # num DIFAT sectors
            difat_bytes,          # 436-byte DIFAT array
        )

        out = bytearray(header)
        for s in sector_payloads:
            out.extend(s)
        return bytes(out)

    # ------------------------------------------------------------------
    # Serializer helpers
    # ------------------------------------------------------------------

    def _build_dir_entry_bytes(
        self,
        entry: DirEntry,
        data_size: int,
        mini_starts: dict[int, int],
        reg_starts: dict[int, int],
        mini_stream_first: int,
        mini_stream_used_bytes: int,
        idx: int,
    ) -> bytes:
        """
        Re-emit a 128-byte directory entry, patching in updated start_sector and
        size while preserving the original CLSID, color, state, and timestamps.
        """
        CUTOFF = 4096

        if entry.obj_type == _OBJTYPE_EMPTY:
            return self._empty_dir_entry_bytes()

        if entry.obj_type == _OBJTYPE_ROOT:
            start_sector = mini_stream_first
            size = mini_stream_used_bytes
        elif entry.obj_type == _OBJTYPE_STREAM:
            if data_size == 0:
                start_sector = _ENDOFCHAIN
                size = 0
            elif data_size < CUTOFF:
                start_sector = mini_starts[idx]
                size = data_size
            else:
                start_sector = reg_starts[idx]
                size = data_size
        else:
            # Storage entry: no stream data.
            start_sector = _ENDOFCHAIN
            size = 0

        # Start from the preserved raw entry (or a synthesised blank if missing).
        if len(entry.raw) == _DIR_ENTRY_SIZE:
            buf = bytearray(entry.raw)
        else:
            buf = bytearray(self._synth_dir_entry_bytes(entry))

        # Patch sibling/child pointers and name in case they have been edited.
        name_utf16 = entry.name.encode("utf-16-le") + b"\x00\x00"
        if len(name_utf16) > 64:
            raise CFBError(f"Directory entry name too long: {entry.name!r}")
        name_bytes = name_utf16 + b"\x00" * (64 - len(name_utf16))
        struct.pack_into("<64sHBB", buf, 0,
                         name_bytes, len(name_utf16), entry.obj_type, buf[67])
        struct.pack_into("<III", buf, 68,
                         entry.left_sibling_id,
                         entry.right_sibling_id,
                         entry.child_id)
        # start_sector at offset 116, size_low at 120, size_high at 124.
        struct.pack_into("<III", buf, 116,
                         start_sector & 0xFFFFFFFF,
                         size & 0xFFFFFFFF,
                         (size >> 32) & 0xFFFFFFFF)
        return bytes(buf)

    def _synth_dir_entry_bytes(self, entry: DirEntry) -> bytes:
        """Build a fresh directory entry from scratch when no preserved raw exists."""
        return struct.pack(
            _DIR_ENTRY_FMT,
            b"\x00" * 64,             # name (patched in by caller)
            0,                        # name_len
            entry.obj_type,
            0,                        # color (red)
            entry.left_sibling_id,
            entry.right_sibling_id,
            entry.child_id,
            b"\x00" * 16,             # CLSID
            0,                        # state
            0,                        # created
            0,                        # modified
            _ENDOFCHAIN,              # start_sector (patched in by caller)
            0,                        # size_low (patched in by caller)
            0,                        # size_high
        )

    def _empty_dir_entry_bytes(self) -> bytes:
        """A fully-empty directory entry slot ([MS-CFB] 2.6.4)."""
        return struct.pack(
            _DIR_ENTRY_FMT,
            b"\x00" * 64, 0, _OBJTYPE_EMPTY, 0,
            _NOSTREAM, _NOSTREAM, _NOSTREAM,
            b"\x00" * 16, 0, 0, 0,
            _FREESECT, 0, 0,
        )
