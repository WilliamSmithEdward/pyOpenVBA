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


# ---------------------------------------------------------------------------
# Writer tests
# ---------------------------------------------------------------------------

class TestCFBWriter:
    """
    Black-box tests for CFB.to_bytes().

    Strategy: build a minimal CFB (no streams), serialize it, reparse, and verify
    the structure round-trips.  Then exercise stream-bearing CFBs by writing the
    output of to_bytes() and reparsing.

    To get a CFB with actual streams to test against, we use the writer itself
    to *construct* a fixture: we take a minimal CFB, monkey-attach a fake
    DirEntry list, and serialize.  This is a closed test loop, but it validates
    every layer of the writer pipeline against the existing reader.
    """

    def test_minimal_roundtrip(self) -> None:
        """An empty CFB serializes to valid bytes that reparse identically."""
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.sector_size == 512
        assert cfb2.mini_sector_size == 64
        assert cfb2.mini_stream_cutoff == 4096
        assert cfb2.list_streams() == []

    def test_output_starts_with_magic(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        assert out[:8] == _MAGIC

    def test_output_is_sector_aligned(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        assert len(out) >= 512
        assert (len(out) - 512) % 512 == 0

    def test_header_version_is_3(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        # major version at offset 26 (after sig=8, clsid=16, minorVer=2)
        major = struct.unpack_from("<H", out, 26)[0]
        assert major == 3

    def test_header_byte_order_marker(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        # byte-order field at offset 28
        bom = struct.unpack_from("<H", out, 28)[0]
        assert bom == 0xFFFE

    def test_header_mini_cutoff_is_4096(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        out = cfb.to_bytes()
        # mini stream cutoff is at offset 56
        cutoff = struct.unpack_from("<I", out, 56)[0]
        assert cutoff == 4096

    def test_write_stream_replaces_data(self) -> None:
        """
        Build a CFB that has one mini-stream by constructing a v3 file with
        a single stream entry, then round-trip a replacement.
        """
        original = _make_cfb_with_one_stream("Hello", b"original data")
        cfb = CFB.from_bytes(original)
        assert cfb.get_stream("Hello") == b"original data"

        cfb.write_stream("Hello", b"replacement!")
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("Hello") == b"replacement!"

    def test_write_stream_unknown_name_raises(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        with pytest.raises(KeyError):
            cfb.write_stream("NoSuchStream", b"x")

    def test_mini_stream_roundtrip(self) -> None:
        """A small stream (<4096 bytes) round-trips through the mini-FAT."""
        original = _make_cfb_with_one_stream("S", b"abc" * 30)   # 90 bytes
        cfb = CFB.from_bytes(original)
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("S") == b"abc" * 30

    def test_regular_stream_roundtrip(self) -> None:
        """A large stream (>=4096 bytes) round-trips through the regular FAT."""
        big = b"X" * 5000
        original = _make_cfb_with_one_stream("Big", big)
        cfb = CFB.from_bytes(original)
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("Big") == big

    def test_replace_then_grow_stream(self) -> None:
        """Replacing a small stream with a large one moves it from mini-FAT to FAT."""
        original = _make_cfb_with_one_stream("Grow", b"tiny")
        cfb = CFB.from_bytes(original)
        cfb.write_stream("Grow", b"Y" * 6000)
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("Grow") == b"Y" * 6000

    def test_replace_then_shrink_stream(self) -> None:
        """Replacing a large stream with a small one moves it to the mini-FAT."""
        original = _make_cfb_with_one_stream("Shrink", b"Z" * 6000)
        cfb = CFB.from_bytes(original)
        cfb.write_stream("Shrink", b"small")
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("Shrink") == b"small"

    def test_empty_stream_roundtrip(self) -> None:
        """A zero-byte stream round-trips correctly."""
        original = _make_cfb_with_one_stream("Empty", b"")
        cfb = CFB.from_bytes(original)
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert cfb2.get_stream("Empty") == b""

    def test_fat_marks_fat_sectors_correctly(self) -> None:
        """Every FAT sector must be marked FATSECT (0xFFFFFFFD) in the FAT."""
        cfb = CFB.from_bytes(_make_cfb_with_one_stream("X", b"data"))
        out = cfb.to_bytes()
        # Parse header to find first FAT sector via DIFAT
        first_fat = struct.unpack_from("<I", out, 76)[0]  # DIFAT starts at offset 76
        fat_offset = 512 + first_fat * 512
        # The FAT entry for `first_fat` itself must be FATSECT
        entry = struct.unpack_from("<I", out, fat_offset + first_fat * 4)[0]
        assert entry == 0xFFFF_FFFD

    def test_directory_first_entry_is_root(self) -> None:
        cfb = CFB.from_bytes(_make_cfb_with_one_stream("X", b"data"))
        out = cfb.to_bytes()
        # first_dir_sector is at offset 48
        (dir_sector,) = struct.unpack_from("<I", out, 48)
        dir_offset: int = 512 + int(dir_sector) * 512
        # object type byte is at offset 66 in dir entry
        obj_type = out[dir_offset + 66]
        assert obj_type == 5    # OBJTYPE_ROOT


# ---------------------------------------------------------------------------
# Fixture builder: CFB with one stream
# ---------------------------------------------------------------------------

def _make_cfb_with_one_stream(name: str, data: bytes) -> bytes:
    """
    Build a v3 CFB containing a root entry and one stream entry.

    Small streams (<4096 bytes) go in the mini-stream; larger ones go in
    sectors via the regular FAT.  Keeps the layout intentionally simple so
    the fixture is independent of CFB.to_bytes().
    """
    SECTOR = 512
    MINI = 64
    CUTOFF = 4096

    def _dir_entry(
        name: str, obj_type: int, child_id: int,
        start_sector: int, size: int, color: int = 1,
    ) -> bytes:
        name_utf16 = name.encode("utf-16-le")
        name_bytes = name_utf16 + b"\x00" * (64 - len(name_utf16))
        return struct.pack(
            "<64sHBBIII16sIQQIII",
            name_bytes,
            len(name_utf16) + 2,
            obj_type,
            color,
            0xFFFF_FFFF, 0xFFFF_FFFF, child_id,
            b"\x00" * 16,
            0, 0, 0,
            start_sector,
            size & 0xFFFFFFFF,
            (size >> 32) & 0xFFFFFFFF,
        )

    sectors: list[bytes] = []
    is_mini = 0 < len(data) < CUTOFF

    if is_mini:
        # ---- Mini path ----
        # Sector layout: [mini-stream] [dir] [mini-FAT] [FAT]
        n_mini = (len(data) + MINI - 1) // MINI if data else 0
        mini_padded = data + b"\x00" * (n_mini * MINI - len(data))
        # Pad mini-stream to one full sector
        mini_stream = mini_padded + b"\x00" * (SECTOR - len(mini_padded))
        sectors.append(mini_stream)            # sector 0 = mini-stream
        mini_stream_first = 0
        mini_stream_size = n_mini * MINI

        stream_start = 0           # mini-sector index inside mini-stream
        stream_size = len(data)

        # Directory sector (sector 1)
        root = _dir_entry("Root Entry", 5, 1, mini_stream_first, mini_stream_size)
        child = _dir_entry(name, 2, 0xFFFF_FFFF, stream_start, stream_size)
        empty = _dir_entry("", 0, 0xFFFF_FFFF, 0xFFFF_FFFE, 0, color=0)
        dir_sector = root + child + empty + empty
        sectors.append(dir_sector)             # sector 1 = directory
        dir_first = 1

        # Mini-FAT sector (sector 2)
        mini_fat = [
            (i + 1) if i + 1 < n_mini else 0xFFFF_FFFE for i in range(n_mini)
        ]
        mini_fat += [0xFFFF_FFFF] * (SECTOR // 4 - len(mini_fat))
        sectors.append(struct.pack(f"<{len(mini_fat)}I", *mini_fat))
        minifat_first = 2

        # FAT sector (sector 3) — chains: [mini-stream]=ENDOFCHAIN,
        # [dir]=ENDOFCHAIN, [minifat]=ENDOFCHAIN, [FAT]=FATSECT
        fat = [0xFFFF_FFFE, 0xFFFF_FFFE, 0xFFFF_FFFE, 0xFFFF_FFFD]
        fat += [0xFFFF_FFFF] * (SECTOR // 4 - len(fat))
        sectors.append(struct.pack(f"<{len(fat)}I", *fat))
        fat_first = 3

        num_minifat = 1
    else:
        # ---- Regular FAT path (large stream) ----
        # Sector layout: [stream sectors] [dir] [FAT]
        if len(data) == 0:
            stream_first = 0xFFFF_FFFE
            stream_size = 0
            stream_sectors = 0
        else:
            stream_sectors = (len(data) + SECTOR - 1) // SECTOR
            padded = data + b"\x00" * (stream_sectors * SECTOR - len(data))
            for k in range(stream_sectors):
                sectors.append(padded[k * SECTOR:(k + 1) * SECTOR])
            stream_first = 0
            stream_size = len(data)

        dir_first = len(sectors)
        root = _dir_entry("Root Entry", 5, 1, 0xFFFF_FFFE, 0)
        child = _dir_entry(name, 2, 0xFFFF_FFFF, stream_first, stream_size)
        empty = _dir_entry("", 0, 0xFFFF_FFFF, 0xFFFF_FFFE, 0, color=0)
        sectors.append(root + child + empty + empty)

        fat_first = len(sectors)
        # FAT for: stream chain (stream_sectors entries), dir (1), FAT (1)
        fat: list[int] = []
        for k in range(stream_sectors):
            fat.append(k + 1 if k + 1 < stream_sectors else 0xFFFF_FFFE)
        fat.append(0xFFFF_FFFE)   # dir sector chain end
        fat.append(0xFFFF_FFFD)   # FAT sector itself
        fat += [0xFFFF_FFFF] * (SECTOR // 4 - len(fat))
        sectors.append(struct.pack(f"<{len(fat)}I", *fat))

        minifat_first = 0xFFFF_FFFE
        num_minifat = 0

    # ---- Header ----
    difat = [fat_first] + [0xFFFF_FFFF] * 108
    difat_packed = struct.pack("<109I", *difat)
    header = struct.pack(
        "<8s16sHHHHH6sIIIIIIIII436s",
        _MAGIC, b"\x00" * 16, 0x003E, 3, 0xFFFE, 9, 6,
        b"\x00" * 6,
        0, 1, dir_first, 0, CUTOFF, minifat_first, num_minifat,
        0xFFFF_FFFE, 0,
        difat_packed,
    )
    return header + b"".join(sectors)
