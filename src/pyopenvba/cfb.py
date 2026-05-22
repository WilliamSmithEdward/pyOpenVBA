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

    # ------------------------------------------------------------------
    # Factory
    # ------------------------------------------------------------------

    @classmethod
    def from_bytes(cls, data: bytes) -> "CFB":
        cfb = cls(data)
        cfb._parse()
        return cfb

    @classmethod
    def from_file(cls, fh: BinaryIO) -> "CFB":
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
        )

    def _read_stream(self, entry: DirEntry) -> bytes:
        if entry.size < self.mini_stream_cutoff and entry.obj_type != _OBJTYPE_ROOT:
            # Read from mini-stream
            raw = bytearray()
            for sect in self._mini_chain(entry.start_sector):
                offset = sect * self.mini_sector_size
                raw.extend(self._mini_stream[offset: offset + self.mini_sector_size])
            return bytes(raw[: entry.size])
        else:
            # Read from regular stream
            raw = bytearray()
            for sect in self._chain(entry.start_sector):
                raw.extend(self._sector(sect))
            return bytes(raw[: entry.size])
