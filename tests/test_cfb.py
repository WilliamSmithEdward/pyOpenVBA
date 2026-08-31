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


class TestCFBStreamRemoval:
    """Coverage for the directory-mutation API (remove_stream / drop_streams_in_storage)."""

    def test_remove_top_level_stream(self) -> None:
        original = _make_cfb_with_one_stream("Doomed", b"goodbye")
        cfb = CFB.from_bytes(original)
        assert "Doomed" in cfb.list_streams()
        cfb.remove_stream("Doomed")
        # Even before serialization the stream is gone from listings.
        assert "Doomed" not in cfb.list_streams()
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert "Doomed" not in cfb2.list_streams()
        with pytest.raises(KeyError):
            cfb2.get_stream("Doomed")

    def test_remove_stream_unknown_name_raises(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        with pytest.raises(KeyError):
            cfb.remove_stream("nope")

    def test_drop_streams_in_storage_predicate(self) -> None:
        import zipfile
        from pathlib import Path

        live = Path(__file__).parent / "live_excel_testing" / "test_macro_workbook.xlsm"
        if not live.exists():
            pytest.skip("live workbook fixture not available")
        with zipfile.ZipFile(live) as zf:
            raw = zf.read("xl/vbaProject.bin")
        cfb = CFB.from_bytes(raw)
        # Live fixture is known to ship with __SRP_0..__SRP_3 under VBA.
        before = [s for s in cfb.list_streams() if s.startswith("__SRP_")]
        assert before, "live fixture lacks __SRP_* streams; refresh the fixture"
        removed = cfb.drop_streams_in_storage("VBA", lambda n: n.startswith("__SRP_"))
        assert sorted(removed) == sorted(before)
        out = cfb.to_bytes()
        cfb2 = CFB.from_bytes(out)
        assert not [s for s in cfb2.list_streams() if s.startswith("__SRP_")]
        # Surviving streams must still be intact.
        assert cfb2.get_stream("dir") == cfb.get_stream("dir")

    def test_drop_streams_unknown_storage_raises(self) -> None:
        cfb = CFB.from_bytes(_make_minimal_cfb())
        with pytest.raises(KeyError):
            cfb.drop_streams_in_storage("NoSuchStorage", lambda n: True)


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


class TestStorageScopedLookup:
    """get/write/list *_in_storage operate on the storage's own child
    subtree only.  Previously they linear-scanned the whole directory,
    so same-named streams in different storages could shadow each other
    and root-level streams appeared to belong to every storage."""

    @staticmethod
    def _cfb_with_duplicate_stream_names() -> CFB:
        import zipfile
        from pathlib import Path

        live = Path(__file__).parent / "live_excel_testing" / "test_macro_workbook.xlsm"
        if not live.exists():
            pytest.skip("live workbook fixture not available")
        with zipfile.ZipFile(live) as zf:
            raw = zf.read("xl/vbaProject.bin")
        cfb = CFB.from_bytes(raw)
        cfb.add_substorage("VBA", "StoreA")
        cfb.add_substorage("VBA", "StoreB")
        cfb.add_stream_to_storage("StoreA", "dup", b"content-a")
        cfb.add_stream_to_storage("StoreB", "dup", b"content-b")
        # Round-trip through the serializer so the test also covers the
        # rebuilt directory tree, not just in-memory state.
        return CFB.from_bytes(cfb.to_bytes())

    def test_same_named_streams_resolve_per_storage(self) -> None:
        cfb = self._cfb_with_duplicate_stream_names()
        assert cfb.get_stream_in_storage("StoreA", "dup") == b"content-a"
        assert cfb.get_stream_in_storage("StoreB", "dup") == b"content-b"

    def test_write_targets_only_the_named_storage(self) -> None:
        cfb = self._cfb_with_duplicate_stream_names()
        cfb.write_stream_in_storage("StoreB", "dup", b"rewritten-b")
        out = CFB.from_bytes(cfb.to_bytes())
        assert out.get_stream_in_storage("StoreA", "dup") == b"content-a"
        assert out.get_stream_in_storage("StoreB", "dup") == b"rewritten-b"

    def test_list_returns_only_direct_children(self) -> None:
        cfb = self._cfb_with_duplicate_stream_names()
        assert cfb.list_streams_in_storage("StoreA") == ["dup"]

    def test_root_level_streams_are_not_storage_children(self) -> None:
        cfb = self._cfb_with_duplicate_stream_names()
        # PROJECT and PROJECTwm are siblings of the VBA storage, not
        # children of it.
        vba_children = cfb.list_streams_in_storage("VBA")
        assert "dir" in vba_children
        assert "PROJECT" not in vba_children
        assert "PROJECTwm" not in vba_children
        with pytest.raises(KeyError):
            cfb.get_stream_in_storage("StoreA", "PROJECT")
        with pytest.raises(KeyError):
            cfb.write_stream_in_storage("StoreA", "PROJECT", b"x")

    def test_lookup_stays_case_insensitive(self) -> None:
        cfb = self._cfb_with_duplicate_stream_names()
        assert cfb.get_stream_in_storage("storea", "DUP") == b"content-a"


# ---------------------------------------------------------------------------
# Path-addressed navigation
# ---------------------------------------------------------------------------

class TestCFBPathNavigation:
    """A UserForm's nested storages repeat names -- two forms each own an
    ``i06`` holding an ``f`` -- so navigating by name alone finds the
    wrong one.  These walk from the root instead."""

    @staticmethod
    def _nested_cfb() -> CFB:
        import zipfile
        from pathlib import Path

        live = Path(__file__).parent / "live_excel_testing" / "nested_form.xlsm"
        if not live.exists():
            pytest.skip("nested form fixture not available")
        with zipfile.ZipFile(live) as zf:
            return CFB.from_bytes(zf.read("xl/vbaProject.bin"))

    def test_empty_path_lists_the_root(self) -> None:
        cfb = self._nested_cfb()
        assert set(cfb.list_storages_at()) == {"VBA", "FrmNested"}
        assert "PROJECT" in cfb.list_streams_at()

    def test_walks_into_nested_storages(self) -> None:
        cfb = self._nested_cfb()
        assert cfb.list_storages_at(["FrmNested"]) == ["i02", "i06"]
        assert cfb.list_storages_at(["FrmNested", "i06"]) == ["i08", "i09"]

    def test_same_named_streams_at_different_depths_stay_distinct(self) -> None:
        cfb = self._nested_cfb()
        outer = cfb.get_stream_at(["FrmNested"], "f")
        page = cfb.get_stream_at(["FrmNested", "i06", "i08"], "f")
        assert outer != page
        assert len(outer) > len(page)

    def test_a_step_that_is_not_a_child_raises(self) -> None:
        """``i08`` exists, but not directly under the form."""
        cfb = self._nested_cfb()
        with pytest.raises(KeyError, match="i08"):
            cfb.list_streams_at(["FrmNested", "i08"])

    def test_missing_stream_names_the_path(self) -> None:
        cfb = self._nested_cfb()
        with pytest.raises(KeyError, match="FrmNested/i06"):
            cfb.get_stream_at(["FrmNested", "i06"], "nosuchstream")

    def test_a_storage_is_not_returned_as_a_stream(self) -> None:
        cfb = self._nested_cfb()
        assert "i06" not in cfb.list_streams_at(["FrmNested"])
        assert "f" not in cfb.list_storages_at(["FrmNested"])

    def test_lookup_is_case_insensitive(self) -> None:
        cfb = self._nested_cfb()
        assert cfb.get_stream_at(["frmnested", "I06"], "F") == cfb.get_stream_at(
            ["FrmNested", "i06"], "f"
        )


class TestCFBPathEditing:
    """Creating a UserForm container means creating a storage with the
    right CLSID; removing one means removing it and everything under it,
    because a child storage no site claims makes the next read refuse."""

    @staticmethod
    def _nested_cfb() -> CFB:
        import zipfile
        from pathlib import Path

        live = Path(__file__).parent / "live_excel_testing" / "nested_form.xlsm"
        if not live.exists():
            pytest.skip("nested form fixture not available")
        with zipfile.ZipFile(live) as zf:
            return CFB.from_bytes(zf.read("xl/vbaProject.bin"))

    _FRAME_CLSID = bytes.fromhex("2020186e60f4ce119bcd00aa00608e01")

    @staticmethod
    def _clsid_of(raw: bytes, name: str) -> bytes:
        """Find a directory entry in serialized bytes and read its CLSID.

        Read from the output rather than the model, so this checks what
        actually lands on disk.  Entries are 128 bytes and sector-aligned,
        with a UTF-16LE name at offset 0 and the CLSID at 80.
        """
        needle = name.encode("utf-16-le") + b"\x00\x00"
        for offset in range(0, len(raw) - 128, 128):
            entry = raw[offset:offset + 128]
            if entry[:len(needle)] == needle and entry[64] == len(needle):
                return entry[80:96]
        raise AssertionError(f"no directory entry named {name!r}")

    def test_a_new_storage_keeps_its_clsid_through_a_round_trip(self) -> None:
        cfb = self._nested_cfb()
        cfb.add_substorage_at(["FrmNested"], "i99", self._FRAME_CLSID)
        cfb.add_stream_at(["FrmNested", "i99"], "f", b"payload")
        raw = cfb.to_bytes()
        out = CFB.from_bytes(raw)
        assert "i99" in out.list_storages_at(["FrmNested"])
        assert out.get_stream_at(["FrmNested", "i99"], "f") == b"payload"
        assert self._clsid_of(raw, "i99") == self._FRAME_CLSID

    def test_removing_a_storage_takes_its_subtree(self) -> None:
        cfb = self._nested_cfb()
        # i06 is the MultiPage: it owns i08 and i09.
        cfb.remove_storage_at(["FrmNested"], "i06")
        out = CFB.from_bytes(cfb.to_bytes())
        assert out.list_storages_at(["FrmNested"]) == ["i02"]
        # The sibling that was not removed is still intact.
        assert len(out.get_stream_at(["FrmNested", "i02"], "f")) == 208

    def test_a_duplicate_storage_name_is_refused(self) -> None:
        cfb = self._nested_cfb()
        with pytest.raises(ValueError, match="already exists"):
            cfb.add_substorage_at(["FrmNested"], "i02")

    def test_a_duplicate_stream_name_is_refused(self) -> None:
        cfb = self._nested_cfb()
        with pytest.raises(ValueError, match="already exists"):
            cfb.add_stream_at(["FrmNested"], "f", b"")

    def test_removing_something_that_is_not_there_raises(self) -> None:
        cfb = self._nested_cfb()
        with pytest.raises(KeyError, match="i99"):
            cfb.remove_storage_at(["FrmNested"], "i99")

    def test_a_clsid_of_the_wrong_length_is_refused(self) -> None:
        cfb = self._nested_cfb()
        with pytest.raises(ValueError, match="16 bytes"):
            cfb.add_substorage_at(["FrmNested"], "i99", b"\x00" * 8)
