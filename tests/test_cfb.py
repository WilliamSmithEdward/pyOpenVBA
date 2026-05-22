"""Tests for the CFB parser."""

from __future__ import annotations

import struct
import pytest

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import CFBError

_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


# ---------------------------------------------------------------------------
# Helpers to build minimal valid CFB bytes for unit testing
# ---------------------------------------------------------------------------

def _make_minimal_cfb() -> bytes:
    """
    Build the smallest possible valid CFB (version 3, 512-byte sectors)
    with a single empty root directory entry and no streams.
    """
    sector_size = 512

    # We need at least:
    #   1 FAT sector  (sector 0)
    #   1 directory sector (sector 1)
    #   root entry only

    # -- Directory entry (128 bytes, root entry) --
    root_name = "Root Entry".encode("utf-16-le")
    root_name_padded = root_name + b"\x00" * (64 - len(root_name))
    root_name_len = len(root_name) + 2  # includes null terminator
    dir_entry = struct.pack(
        "<64sHBBIII16sIQQIII",
        root_name_padded,      # name (64 bytes)
        root_name_len,         # name length
        5,                     # obj_type = root
        1,                     # color = red
        0xFFFF_FFFF,           # left sibling
        0xFFFF_FFFF,           # right sibling
        0xFFFF_FFFF,           # child
        b"\x00" * 16,          # CLSID
        0,                     # state bits
        0,                     # created
        0,                     # modified
        0xFFFF_FFFE,           # start sector = ENDOFCHAIN (no mini stream)
        0,                     # size_low = 0
        0,                     # size_high = 0
    )
    dir_sector = dir_entry + b"\x00" * (sector_size - len(dir_entry))

    # -- FAT sector: sector 0 = FATSECT, sector 1 = ENDOFCHAIN (dir chain) --
    fat_entries = [0xFFFF_FFFD, 0xFFFF_FFFE] + [0xFFFF_FFFF] * (sector_size // 4 - 2)
    fat_sector = struct.pack(f"<{len(fat_entries)}I", *fat_entries)

    # -- DIFAT array: first entry points to sector 0 (FAT), rest FREESECT --
    difat = [0] + [0xFFFF_FFFF] * 108
    difat_packed = struct.pack("<109I", *difat)

    # -- Header (512 bytes) --
    # Spec field order: sig(8), CLSID(16), minorVer, majorVer, byteOrder,
    # sectorShift, miniSectorShift (all H), reserved(6s), then 9 DWORDs, DIFAT(436)
    header = struct.pack(
        "<8s16sHHHHH6sIIIIIIIII436s",
        _MAGIC,                # signature
        b"\x00" * 16,          # CLSID
        0x003E,                # minor version
        3,                     # major version
        0xFFFE,                # byte order (little-endian)
        9,                     # sector shift (2^9 = 512)
        6,                     # mini sector shift (2^6 = 64)
        b"\x00" * 6,           # reserved
        0,                     # num dir sectors (v3 = 0)
        1,                     # num FAT sectors
        1,                     # root dir start sector
        0,                     # transaction signature
        0x1000,                # mini stream cutoff = 4096
        0xFFFF_FFFE,           # miniFAT start = ENDOFCHAIN
        0,                     # num miniFAT sectors
        0xFFFF_FFFE,           # DIFAT start = ENDOFCHAIN
        0,                     # num DIFAT sectors
        difat_packed,
    )

    return header + fat_sector + dir_sector


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestCFBParsing:
    def test_parse_minimal_cfb(self) -> None:
        data = _make_minimal_cfb()
        cfb = CFB.from_bytes(data)
        assert cfb.sector_size == 512
        assert cfb.mini_sector_size == 64
        assert cfb.mini_stream_cutoff == 0x1000

    def test_directory_has_root(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        assert cfb.list_streams() == []
        # The root entry exists — list_streams returns [] for a CFB with only a root
        assert isinstance(cfb, CFB)

    def test_list_streams_empty(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        assert cfb.list_streams() == []

    def test_bad_magic_raises(self) -> None:
        data = bytearray(_make_minimal_cfb())
        data[0] = 0x00  # corrupt the magic
        with pytest.raises(CFBError, match="magic"):
            CFB.from_bytes(bytes(data))

    def test_too_small_raises(self) -> None:
        with pytest.raises(CFBError, match="too small"):
            CFB.from_bytes(b"\xd0\xcf\x11\xe0")

    def test_get_stream_missing_raises(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        with pytest.raises(KeyError):
            cfb.get_stream("nonexistent")
