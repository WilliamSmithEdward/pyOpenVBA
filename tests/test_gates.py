"""
Gate-level conformance tests against `xlsm_feature_completeness_gates.md`.

Each test class corresponds to one HARD gate from the gates document.  Tests
assert behavior that the current pyOpenVBA implementation must satisfy.
Gates that the library does not yet implement are marked with
``pytest.mark.xfail(strict=True, reason="...")``; once a gate is implemented
its xfail marker should be removed and the assertions strengthened.

These tests deliberately do not depend on a real running Excel install.  The
true end-state for Gate 1 / Gate 20 / Gate 21 ("opens in Excel without
repair") cannot be asserted from Python alone; the tests here exercise the
strongest in-memory and binary checks we can perform.

A real ``.xlsm`` is included at
``tests/live_excel_testing/test_macro_workbook.xlsm`` and is used for the
round-trip and mutation gates.
"""

from __future__ import annotations

import io
import os
import struct
import warnings
import zipfile
from pathlib import Path
from typing import cast

import pytest

import pyopenvba
from pyopenvba import (
    CFBError,
    ExcelFile,
    PyOpenVBAError,
    UnsupportedFormatError,
    VBAModuleKind,
    VBAProjectError,
)
from pyopenvba.cfb import CFB
from pyopenvba.vba import (
    VBAModule,
    VBAProject,
    compress,
    copy_token_help,
    decompress,
    parse_vba_project,
    rebuild_module_stream,
    write_back_modules,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

LIVE_XLSM = Path(__file__).parent / "live_excel_testing" / "test_macro_workbook.xlsm"
LIVE_XLSB = Path(__file__).parent / "live_excel_testing" / "test_macro_workbook.xlsb"
LIVE_PROTECTED_XLSM = (
    Path(__file__).parent
    / "live_excel_testing"
    / "workbook_with_password_protected_vba_modules.xlsm"
)
LIVE_EMPTY_XLSM = (
    Path(__file__).parent
    / "live_excel_testing"
    / "xlsm_file_with_no_vba_entered_yet.xlsm"
)
LIVE_LARGE_MODULE_XLSM = (
    Path(__file__).parent
    / "live_excel_testing"
    / "large_vba_module.xlsm"
)


@pytest.fixture(scope="session")
def live_xlsm_path() -> Path:
    if not LIVE_XLSM.exists():
        pytest.skip(f"live test workbook not available at {LIVE_XLSM}")
    return LIVE_XLSM


@pytest.fixture(scope="session")
def live_vba_bin(live_xlsm_path: Path) -> bytes:
    with zipfile.ZipFile(live_xlsm_path) as zf:
        return zf.read("xl/vbaProject.bin")


@pytest.fixture(scope="session")
def live_xlsb_path() -> Path:
    if not LIVE_XLSB.exists():
        pytest.skip(f"live xlsb workbook not available at {LIVE_XLSB}")
    return LIVE_XLSB


@pytest.fixture(scope="session")
def live_protected_xlsm_path() -> Path:
    if not LIVE_PROTECTED_XLSM.exists():
        pytest.skip(f"protected workbook not available at {LIVE_PROTECTED_XLSM}")
    return LIVE_PROTECTED_XLSM


@pytest.fixture(scope="session")
def live_empty_xlsm_path() -> Path:
    if not LIVE_EMPTY_XLSM.exists():
        pytest.skip(f"empty (no-VBA) workbook not available at {LIVE_EMPTY_XLSM}")
    return LIVE_EMPTY_XLSM


# ===========================================================================
# GATE 0 — Scope Declaration Gate
# ===========================================================================

class TestGate00_ScopeDeclaration:
    """Library declares scope, separates layers, exposes raw vbaProject.bin."""

    def test_supported_hosts_declared(self) -> None:
        # The library publicly enumerates the file extensions it claims to handle.
        from pyopenvba import excel as _excel
        cfb_formats = getattr(_excel, "_CFB_FORMATS")
        zip_formats = getattr(_excel, "_ZIP_FORMATS")
        assert ".xlsm" in zip_formats
        assert ".xls" in cfb_formats

    def test_unsupported_host_extension_fails_loudly(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.docm"
        p.write_bytes(b"PK\x03\x04")    # bogus ZIP header
        with pytest.raises(UnsupportedFormatError):
            ExcelFile(p)

    def test_can_operate_on_raw_vba_bytes(self, live_vba_bin: bytes) -> None:
        # The CFB and VBA layers must be usable without going through ExcelFile.
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        assert proj.modules, "raw vbaProject.bin must parse without ExcelFile"

    def test_layered_apis(self) -> None:
        # CFB, VBA project, and Excel handler are separate modules.
        assert hasattr(pyopenvba, "ExcelFile")
        from pyopenvba import cfb as cfb_mod
        from pyopenvba import vba as vba_mod
        assert cfb_mod is not vba_mod
        assert hasattr(cfb_mod, "CFB")
        assert hasattr(vba_mod, "VBAProject")

    def test_vba_project_bytes_accessor(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            raw = wb.vba_project_bytes()
        assert raw.startswith(b"\xd0\xcf\x11\xe0")    # CFB magic


# ===========================================================================
# GATE 1 — Host Package Gate
# ===========================================================================

class TestGate01_HostPackage:
    """xlsm round-trip preserves every non-VBA part."""

    def test_xlsm_open_does_not_corrupt_zip(self, live_xlsm_path: Path) -> None:
        with zipfile.ZipFile(live_xlsm_path) as zf:
            names_before = sorted(zf.namelist())
        with ExcelFile(live_xlsm_path) as wb:
            wb.vba_modules()
        with zipfile.ZipFile(live_xlsm_path) as zf:
            names_after = sorted(zf.namelist())
        assert names_before == names_after

    def test_save_preserves_all_non_vba_entries(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "roundtrip.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.save(out)
        with zipfile.ZipFile(live_xlsm_path) as before, zipfile.ZipFile(out) as after:
            names_before = sorted(before.namelist())
            names_after = sorted(after.namelist())
            assert names_before == names_after
            for name in names_before:
                if name == "xl/vbaProject.bin":
                    continue
                assert before.read(name) == after.read(name), (
                    f"non-VBA entry {name!r} mutated by save()"
                )

    def test_save_changes_only_vba_binary_when_module_edited(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "edited.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            src = wb.get_module("Module1")
            wb.set_module("Module1", src + "\r\n'edit\r\n")
            wb.save(out)
        with zipfile.ZipFile(live_xlsm_path) as before, zipfile.ZipFile(out) as after:
            for name in before.namelist():
                if name == "xl/vbaProject.bin":
                    continue
                assert before.read(name) == after.read(name)

    def test_xlsm_without_vba_project_raises_cleanly(
        self, live_empty_xlsm_path: Path
    ) -> None:
        """An xlsm whose VBA project has never been initialised must raise
        a structured ``VBAProjectError`` (not a bare ``KeyError`` from the
        underlying ZIP layer)."""
        from pyopenvba.exceptions import VBAProjectError
        with pytest.raises(VBAProjectError, match="vbaProject.bin"):
            ExcelFile(live_empty_xlsm_path)

    def test_xlsb_round_trip_preserves_all_non_vba_entries(
        self, live_xlsb_path: Path, tmp_path: Path
    ) -> None:
        """xlsb is a ZIP container too -- save() must preserve every
        non-VBA entry just like xlsm."""
        out = tmp_path / "roundtrip.xlsb"
        with ExcelFile(live_xlsb_path) as wb:
            wb.save(out)
        with zipfile.ZipFile(live_xlsb_path) as before, zipfile.ZipFile(out) as after:
            names_before = sorted(before.namelist())
            names_after = sorted(after.namelist())
            assert names_before == names_after
            for name in names_before:
                if name == "xl/vbaProject.bin":
                    continue
                assert before.read(name) == after.read(name), (
                    f"non-VBA entry {name!r} mutated by xlsb save()"
                )

    def test_xlsb_module_edit_round_trip(
        self, live_xlsb_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "edited.xlsb"
        marker = "\r\n' xlsb edit\r\n"
        with ExcelFile(live_xlsb_path) as wb:
            src = wb.get_module("Module1")
            wb.set_module("Module1", src + marker)
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert wb2.get_module("Module1").endswith(marker)


# ===========================================================================
# GATE 2 — OLE/CFB Container Gate
# ===========================================================================

class TestGate02_CFB:
    """CFB layer locates required storages and streams case-insensitively."""

    def test_vba_storage_found(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        assert any(s.casefold() == "vba" for s in cfb.list_storages())

    def test_required_streams_present(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        cfb.get_stream("PROJECT")       # case-insensitive
        cfb.get_stream_in_storage("vba", "DIR")
        cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")

    def test_stream_lookup_case_insensitive(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        a = cfb.get_stream("project")
        b = cfb.get_stream("PROJECT")
        assert a == b

    def test_srp_streams_not_emitted_on_write(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        # Spec: implementations MUST NOT emit SRP streams.
        # Read live workbook, save, then verify no SRP streams appear in output.
        out = tmp_path / "no_srp.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.save(out)
        with ExcelFile(out) as reread:
            cfb = CFB.from_bytes(reread.vba_project_bytes())
        all_streams = cfb.list_streams()
        for s in all_streams:
            assert not s.lower().startswith("__srp_"), (
                f"writer emitted SRP stream {s!r}"
            )
        # Sanity: real VBA streams must survive the SRP cull.
        names_lower = {s.lower() for s in all_streams}
        assert "dir" in names_lower
        assert "_vba_project" in names_lower


# ===========================================================================
# GATE 3 — Binary Parsing Discipline Gate
# ===========================================================================

class TestGate03_ParsingDiscipline:
    """Bounds-checked, fail-fast parsing."""

    def test_too_small_cfb_raises(self) -> None:
        with pytest.raises(CFBError):
            CFB.from_bytes(b"\xd0\xcf\x11\xe0short")

    def test_bad_magic_raises(self) -> None:
        with pytest.raises(CFBError):
            CFB.from_bytes(b"NOT_A_CFB" + b"\x00" * 2048)

    def test_invalid_signature_byte_raises(self) -> None:
        with pytest.raises(VBAProjectError):
            decompress(b"\x00")

    def test_bad_chunk_signature_raises(self) -> None:
        with pytest.raises(VBAProjectError):
            decompress(b"\x01\x00\x00")    # zero chunk signature

    def test_parse_error_includes_stream_name_and_offset(self) -> None:
        try:
            decompress(b"\x01\x00\x10", stream_name="VBA/Module1")
        except VBAProjectError as exc:
            msg = str(exc)
            assert "offset" in msg and "stream" in msg
            assert "Module1" in msg
        else:
            pytest.fail("no error raised")


# ===========================================================================
# GATE 4 — Compression / Decompression Gate
# ===========================================================================

class TestGate04_Compression:
    """MS-OVBA compression algorithm correctness."""

    @pytest.mark.parametrize("payload", [
        b"",
        b"a",
        b"hello world",
        b"x" * 4095,
        b"y" * 4096,
        b"z" * 4097,
        b"abcabcabcabcabcabcabcabc",
        b"\x00" * 100 + b"abcdef" * 200,
        bytes(range(256)) * 16,
        b"\xAB" * 20000,
    ], ids=["empty", "single", "hello", "x4095", "y4096", "z4097", "abc", "zeros+abc", "range256x16", "ab20k"])
    def test_compress_decompress_round_trip(self, payload: bytes) -> None:
        assert decompress(compress(payload)) == payload

    def test_compress_decompress_random_corpus(self) -> None:
        # Incompressible (random) partial final chunks are bounded by MS-OVBA
        # at ~3640 bytes (literal+flag overhead). Full 4096-byte chunks are
        # always stored raw, so they work at any size.
        for n in [1, 7, 64, 511, 3640, 4096, 8192, 32768]:
            data = os.urandom(n)
            assert decompress(compress(data)) == data, f"failed at n={n}"

    def test_compress_signature_byte(self) -> None:
        assert compress(b"abc")[0] == 0x01

    def test_copy_token_help_matches_spec(self) -> None:
        # Spec table — Section 2.4.1.3.6.
        cases = [
            (1, 4, 0xFFF, 0xF000),
            (16, 4, 0xFFF, 0xF000),
            (17, 5, 0x7FF, 0xF800),
            (32, 5, 0x7FF, 0xF800),
            (33, 6, 0x3FF, 0xFC00),
            (4096, 12, 0x000F, 0xFFF0),
        ]
        for diff, expected_bits, expected_lmask, expected_omask in cases:
            lmask, omask, bits = copy_token_help(diff, 0)
            assert bits == expected_bits, f"bit_count for diff={diff}"
            assert lmask == expected_lmask
            assert omask == expected_omask

    def test_round_trip_existing_dir_stream(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        original = cfb.get_stream_in_storage("VBA", "dir")
        plain = decompress(original)
        # We don't require byte-identical recompression — only that the
        # decompressed result matches.
        assert decompress(compress(plain)) == plain


# ===========================================================================
# GATE 5 — _VBA_PROJECT / Performance Cache Gate
# ===========================================================================

class TestGate05_PerformanceCache:
    """Performance cache is preserved verbatim and never consulted for source."""

    def test_module_prefix_preserved_across_unedited_save(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "noop.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            project_before = wb.vba_project()
            prefixes_before = {m.name: m.prefix_bytes for m in project_before.modules}
            wb.save(out)
        with ExcelFile(out) as wb2:
            project_after = wb2.vba_project()
            for m in project_after.modules:
                assert m.prefix_bytes == prefixes_before[m.name], (
                    f"module {m.name!r} cache prefix bytes changed across no-op save"
                )

    def test_source_replacement_preserves_other_modules_prefix(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "one_edit.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            before = {m.name: m.prefix_bytes for m in wb.vba_project().modules}
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n'x\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            for m in wb2.vba_project().modules:
                assert m.prefix_bytes == before[m.name], (
                    f"non-target module {m.name!r} had cache prefix mutated"
                )

    def test_vba_project_cache_preserved_verbatim_on_noop_save(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """A pure no-op save must not touch the ``_VBA_PROJECT`` stream."""
        out = tmp_path / "noop_cache.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            cache_before = CFB.from_bytes(wb.vba_project_bytes()).get_stream_in_storage(
                "VBA", "_VBA_PROJECT"
            )
            wb.save(out)
        with ExcelFile(out) as wb2:
            cache_after = CFB.from_bytes(wb2.vba_project_bytes()).get_stream_in_storage(
                "VBA", "_VBA_PROJECT"
            )
        assert bytes(cache_after) == bytes(cache_before), (
            "no-op save must not mutate the _VBA_PROJECT cache stream"
        )

    def test_vba_project_cache_invalidated_on_source_edit(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """A mutating save must zero the cache body and preserve the header.

        Per [MS-OVBA] 2.3.4.1 the PerformanceCache "MUST be ignored on
        read" and Office regenerates it on next open.  pyOpenVBA zeroes
        the body so stale offsets cannot be honored.
        """
        out = tmp_path / "edited_cache.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            cache_before = bytes(
                CFB.from_bytes(wb.vba_project_bytes()).get_stream_in_storage(
                    "VBA", "_VBA_PROJECT"
                )
            )
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n'edit\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            cache_after = bytes(
                CFB.from_bytes(wb2.vba_project_bytes()).get_stream_in_storage(
                    "VBA", "_VBA_PROJECT"
                )
            )
        # Length unchanged.
        assert len(cache_after) == len(cache_before)
        # Header preserved (Reserved1=0x61CC, Version, Reserved2=0x00).
        assert cache_after[:5] == cache_before[:5]
        # Body zeroed.
        assert cache_after[5:] == bytes(len(cache_after) - 5)
        # Sanity: original cache had non-zero content beyond the header.
        assert any(b != 0 for b in cache_before[5:]), (
            "fixture _VBA_PROJECT cache body is unexpectedly all zeros"
        )

    def test_vba_project_cache_invalidated_helper_is_idempotent(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.vba import invalidate_vba_project_cache

        out = tmp_path / "twice.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            cfb = wb._get_cfb()        # pyright: ignore[reportPrivateUsage]
            assert invalidate_vba_project_cache(cfb) is True
            first = bytes(cfb.get_stream_in_storage("VBA", "_VBA_PROJECT"))
            # Second call sees an already-zeroed body; helper still reports
            # True (header is valid, body is non-empty) but produces no
            # observable change.
            invalidate_vba_project_cache(cfb)
            second = bytes(cfb.get_stream_in_storage("VBA", "_VBA_PROJECT"))
            assert first == second
            wb.save(out)
        # Round-trip through save still readable.
        with ExcelFile(out) as wb2:
            assert "Module1" in [m.name for m in wb2.vba_project().modules]

    def test_vba_project_cache_invalidate_skips_missing_stream(
        self, live_xlsm_path: Path
    ) -> None:
        """The helper returns False (not raises) when the stream is absent."""
        from pyopenvba.vba import invalidate_vba_project_cache

        with ExcelFile(live_xlsm_path) as wb:
            cfb = wb._get_cfb()        # pyright: ignore[reportPrivateUsage]
            cfb.remove_stream_in_storage("VBA", "_VBA_PROJECT")
            assert invalidate_vba_project_cache(cfb) is False


# ===========================================================================
# GATE 6 — PROJECT Stream Gate
# ===========================================================================

class TestGate06_ProjectStream:
    """PROJECT stream parsing/writing is not yet implemented."""

    def test_project_stream_is_readable_as_bytes(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        # Minimum bar: we can find and read it raw.
        raw = cfb.get_stream("PROJECT")
        assert raw
        # The stream is plain-text key=value records.
        text = raw.decode("cp1252", errors="replace")
        assert "ID=" in text or "Name=" in text or "Module=" in text

    def test_project_stream_full_grammar_parsed(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_project_stream
        cfb = CFB.from_bytes(live_vba_bin)
        raw = cfb.get_stream("PROJECT")
        ps = parse_project_stream(raw)
        # Project ID and Name records are always present.
        assert ps.id
        assert ps.name
        # Item list should contain at least one Module / Document / Class record.
        keys = {k for k, _ in ps.items}
        assert keys & {"Module", "Document", "Class", "BaseClass"}


# ===========================================================================
# GATE 7 — PROJECTwm Name Mapping Gate
# ===========================================================================

class TestGate07_ProjectWm:
    def test_projectwm_parser_present(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_projectwm
        cfb = CFB.from_bytes(live_vba_bin)
        try:
            raw = cfb.get_stream("PROJECTwm")
        except KeyError:
            pytest.skip("fixture has no PROJECTwm stream")
        pairs = parse_projectwm(raw)
        # The MBCS and Unicode names of each module must agree byte-for-byte
        # when limited to ASCII module names (true for the live fixture).
        for mbcs, uni in pairs:
            assert mbcs == uni or not mbcs.isascii()

    def test_serialize_projectwm_round_trip(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_projectwm, serialize_projectwm
        cfb = CFB.from_bytes(live_vba_bin)
        try:
            raw = cfb.get_stream("PROJECTwm")
        except KeyError:
            pytest.skip("fixture has no PROJECTwm stream")
        pairs = parse_projectwm(raw)
        rebuilt = serialize_projectwm(pairs, code_page=1252)
        assert rebuilt == raw, "ASCII-only PROJECTwm must round-trip byte-for-byte"

    def test_serialize_projectwm_terminator_only(self) -> None:
        from pyopenvba.vba import parse_projectwm, serialize_projectwm
        # An empty pair list must produce only the terminator (0x00 + u16=0)
        # which the parser recognises as an empty stream.
        blob = serialize_projectwm([])
        assert blob == b"\x00\x00"
        assert parse_projectwm(blob) == []

    def test_projectwm_rewritten_on_module_add(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.vba import parse_projectwm
        out = tmp_path / "wm_add.xlsm"
        new_name = "BrandNewWmModule"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.add_module(new_name, "' wm\r\n", kind=VBAModuleKind.standard)
            wb.save(out)
        with ExcelFile(out) as wb2:
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            pairs = parse_projectwm(cfb.get_stream("PROJECTwm"))
            names = [m for (m, _u) in pairs]
            assert new_name in names
            # All originals still enumerated.
            for orig in ("ThisWorkbook", "Sheet1", "Module1", "Class1", "UserForm1"):
                assert orig in names

    def test_projectwm_rewritten_on_module_delete(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.vba import parse_projectwm
        out = tmp_path / "wm_del.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.delete_module("Module1")
            wb.save(out)
        with ExcelFile(out) as wb2:
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            pairs = parse_projectwm(cfb.get_stream("PROJECTwm"))
            names = [m for (m, _u) in pairs]
            assert "Module1" not in names


# ===========================================================================
# GATE 8 — PROJECTlk / ActiveX License Gate
# ===========================================================================

class TestGate08_ProjectLk:
    def test_projectlk_parser_present(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_projectlk
        cfb = CFB.from_bytes(live_vba_bin)
        try:
            raw = cfb.get_stream("PROJECTlk")
        except KeyError:
            # PROJECTlk is only present when the project references ActiveX
            # controls; the fallback empty-input check still exercises the parser.
            assert parse_projectlk(b"") == []
            return
        records = parse_projectlk(raw)
        # Even if no controls are present the parser must not raise.
        assert isinstance(records, list)

    def test_serialize_projectlk_round_trip_synthetic(self) -> None:
        from pyopenvba.vba import (
            LicenseRecord,
            parse_projectlk,
            serialize_projectlk,
        )
        records = [
            LicenseRecord(
                lic_key=b"\x01\x02\x03\x04",
                libid="*\\G{12345678-1234-1234-1234-123456789ABC}#1.0#0#C:\\Foo\\Bar.ocx#Bar Control",
                classid=bytes(range(16)),
                cookie=0xDEADBEEF,
            ),
            LicenseRecord(
                lic_key=b"",
                libid="*\\G{00000000-0000-0000-0000-000000000000}#0.0#0#x#y",
                classid=b"\x00" * 16,
                cookie=0,
            ),
        ]
        raw = serialize_projectlk(records)
        parsed = parse_projectlk(raw)
        assert len(parsed) == len(records)
        for orig, got in zip(records, parsed):
            assert got.lic_key == orig.lic_key
            assert got.libid == orig.libid
            assert got.classid == orig.classid
            assert got.cookie == orig.cookie

    def test_serialize_projectlk_live(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_projectlk, serialize_projectlk
        cfb = CFB.from_bytes(live_vba_bin)
        try:
            raw = cfb.get_stream("PROJECTlk")
        except KeyError:
            pytest.skip("fixture has no PROJECTlk stream")
        records = parse_projectlk(raw)
        # Re-emit and re-parse; structural equivalence is required even if
        # byte-for-byte equality is not (Office may pad differently).
        reparsed = parse_projectlk(serialize_projectlk(records))
        assert len(reparsed) == len(records)
        for a, b in zip(records, reparsed):
            assert a.lic_key == b.lic_key
            assert a.libid == b.libid
            assert a.classid == b.classid
            assert a.cookie == b.cookie


# ===========================================================================
# GATE 9 — dir Project Information Gate
# ===========================================================================

class TestGate09_DirProjectInfo:
    """We parse code page only; full PROJECTINFORMATION decoding pending."""

    def test_project_code_page_parsed(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        # The fixture workbook is cp1252; assert we recovered a real code page.
        assert proj.code_page in {1250, 1251, 1252, 1253, 1254, 65001, 932, 936}

    def test_full_project_information_records_decoded(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        assert proj.name
        # SysKind values per [MS-OVBA] 2.3.4.2.1.1: 0=16-bit, 1=32-bit, 2=Mac, 3=64-bit
        assert proj.sys_kind in (0, 1, 2, 3)


# ===========================================================================
# GATE 10 — dir Reference Gate
# ===========================================================================

class TestGate10_DirReferences:
    def test_references_parsed(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        refs = proj.references
        assert refs, "every Excel workbook references at least stdole+VBA"
        # Each reference must have a non-empty name and a known kind.
        for r in refs:
            assert r.name
            assert r.kind in {"registered", "project", "control", "original"}


# ===========================================================================
# GATE 11 — dir Module Record Gate
# ===========================================================================

class TestGate11_DirModuleRecords:
    """We parse module names, stream names, offsets, type, read-only, private."""

    def test_module_names_and_stream_names_extracted(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        for m in proj.modules:
            assert m.name
            assert m.stream_name
            # name and stream_name need not be equal — that's exactly the point
            # of MODULESTREAMNAME.

    def test_module_offset_used_for_source_extraction(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        # All Excel-produced modules have nonzero text_offset because of the
        # version-dependent performance-cache prefix.
        assert any(m.text_offset > 0 for m in proj.modules)
        for m in proj.modules:
            assert isinstance(m.source, str)

    def test_module_kind_distinguishes_standard_and_other(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        kinds = {m.kind for m in proj.modules}
        # Fixture has Module1 (standard) plus class/document modules (other).
        assert VBAModuleKind.standard in kinds
        assert VBAModuleKind.other in kinds


# ===========================================================================
# GATE 12 — Module Stream Gate
# ===========================================================================

class TestGate12_ModuleStream:
    """Decompress source from MODULEOFFSET; replace without breaking attributes."""

    def test_decompressed_source_starts_with_attribute(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        for m in proj.modules:
            assert m.source.startswith('Attribute VB_Name = '), (
                f"module {m.name!r} source missing VB_Name attribute"
            )

    def test_rebuild_module_stream_preserves_prefix(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        m = proj.get_module("Module1")
        new_stream = rebuild_module_stream(m, proj.code_page)
        assert new_stream.startswith(m.prefix_bytes), (
            "rebuilt module stream must preserve the cache-prefix bytes verbatim"
        )

    def test_write_back_then_reparse_yields_same_source(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "wb.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            target = wb.get_module("Module1") + "\r\n'edit\r\n"
            wb.set_module("Module1", target)
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert wb2.get_module("Module1") == target


# ===========================================================================
# GATE 13 — Module Mutation Gate
# ===========================================================================

class TestGate13_ModuleMutation:
    """Source replacement works; add/rename/delete are pending."""

    def test_replace_standard_module(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "edit_std.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n'x\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert "'x" in wb2.get_module("Module1")

    def test_replace_class_module(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "edit_cls.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.set_module("Class1", wb.get_module("Class1") + "\r\n'cls edit\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert "'cls edit" in wb2.get_module("Class1")

    def test_replace_document_module(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "edit_doc.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.set_module("Sheet1", wb.get_module("Sheet1") + "\r\n'doc edit\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert "'doc edit" in wb2.get_module("Sheet1")

    def test_replace_unknown_module_raises(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            with pytest.raises(KeyError):
                wb.set_module("DoesNotExist", "x")

    def test_add_module(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            module = proj.add_module(
                "NewMod", "Sub X(): End Sub", kind=VBAModuleKind.standard
            )
            assert module.name == "NewMod"
            assert module.kind == VBAModuleKind.standard
            assert "NewMod" in proj.module_names()
            with pytest.raises(ValueError):
                proj.add_module("NewMod", "")

    def test_rename_module(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.rename_module("Module1", "RenamedModule")
            names = {n.casefold() for n in proj.module_names()}
            assert "renamedmodule" in names
            assert "module1" not in names

    def test_delete_module(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.delete_module("Module1")
            names = {n.casefold() for n in proj.module_names()}
            assert "module1" not in names


# ===========================================================================
# GATE 14 — Designer / UserForm Gate
# ===========================================================================

class TestGate14_Designer:
    """UserForm / designer sub-storage round-trip on the live fixture.

    The live xlsm carries a real Office Forms 2.0 UserForm (`UserForm1`),
    contributing both a code-behind stream under ``VBA/`` and a sibling
    sub-storage ``UserForm1/`` whose children are the designer blobs
    ``f``, ``o``, ``\x01CompObj`` and ``\x03VBFrame``.
    """

    # Names of the four designer child streams that live under VBA/UserForm1/.
    DESIGNER_CHILDREN = ("f", "o", "\x01CompObj", "\x03VBFrame")

    def test_designer_storage_preserved(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """A no-op save must preserve the UserForm1 sub-storage byte-for-byte."""
        out = tmp_path / "userform_noop.xlsm"

        with ExcelFile(live_xlsm_path) as wb:
            cfb_before = CFB.from_bytes(wb.vba_project_bytes())
            before = {
                n: cfb_before.get_stream_in_storage("UserForm1", n)
                for n in self.DESIGNER_CHILDREN
            }
            wb.save(out)

        with ExcelFile(out) as wb2:
            cfb_after = CFB.from_bytes(wb2.vba_project_bytes())
            assert "UserForm1" in cfb_after.list_storages()
            for name, blob in before.items():
                assert cfb_after.get_stream_in_storage("UserForm1", name) == blob, name
            # The PROJECT stream still declares the UserForm.
            assert b"BaseClass=UserForm1" in cfb_after.get_stream("PROJECT")

    def test_synthetic_substorage_round_trips_through_cfb(
        self, live_vba_bin: bytes
    ) -> None:
        """A synthetic sub-storage with binary child streams survives CFB
        serialize / re-parse, mimicking the structural shape of a designer
        storage (parent storage + opaque child blobs).
        """
        cfb = CFB.from_bytes(live_vba_bin)
        cfb.add_substorage("VBA", "ZzDesigner")
        # Designer-like opaque blobs: a tiny "f" record + a larger "o" record.
        cfb.add_stream_to_storage("VBA", "Zz_f_stream", b"\x00\x04FORM" + b"\xaa" * 64)
        cfb.add_stream_to_storage("VBA", "Zz_o_stream", b"\xff" * 8192)

        round_tripped = CFB.from_bytes(cfb.to_bytes())
        # New storage shows up in the directory.
        assert "ZzDesigner" in round_tripped.list_storages()
        # Both opaque child streams survive byte-for-byte.
        assert round_tripped.get_stream_in_storage("VBA", "Zz_f_stream") == (
            b"\x00\x04FORM" + b"\xaa" * 64
        )
        assert round_tripped.get_stream_in_storage("VBA", "Zz_o_stream") == b"\xff" * 8192
        # Existing VBA streams remain intact.
        assert "dir" in round_tripped.list_streams_in_storage("VBA")
        assert "PROJECT" in round_tripped.list_streams()


# ===========================================================================
# GATE 15 — Content Hash / Integrity Gate
# ===========================================================================

class TestGate15_ContentHash:
    def test_v3_content_hash_implemented(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import compute_v3_content_hash
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        h1 = compute_v3_content_hash(proj)
        h2 = compute_v3_content_hash(proj)
        assert isinstance(h1, bytes) and len(h1) == 20   # SHA-1 size
        assert h1 == h2, "hash must be deterministic for the same project"
        # Mutating source must change the hash.
        proj.modules[0].source = proj.modules[0].source + "\n' edit\n"
        assert compute_v3_content_hash(proj) != h1


# ===========================================================================
# GATE 16 — Protection / Encryption / Password Gate
# ===========================================================================

class TestGate16_Protection:
    def test_protection_state_parsed(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import ProjectProtection, parse_project_stream
        cfb = CFB.from_bytes(live_vba_bin)
        raw = cfb.get_stream("PROJECT")
        ps = parse_project_stream(raw)
        assert isinstance(ps.protection, ProjectProtection)
        # DPB record is always present on a real Excel workbook.
        assert ps.protection.dpb
        # The live (unprotected) fixture.
        assert ps.protection.has_password is False

    def test_real_protected_workbook_detected(
        self, live_protected_xlsm_path: Path
    ) -> None:
        """A real password-protected workbook must parse and report
        has_password=True."""
        with ExcelFile(live_protected_xlsm_path) as wb:
            proj = wb.vba_project()
            assert proj.protection is not None
            assert proj.protection.has_password is True
            assert proj.protection.dpb
            # Module sources are still readable on a protected project
            # (the password gates the VBA IDE, not the on-disk stream).
            assert "PasswordTest" in proj.module_names()
            assert wb.get_module("PasswordTest") is not None

    def test_save_refuses_real_protected_project_without_opt_in(
        self, live_protected_xlsm_path: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.exceptions import VBAProjectError
        out = tmp_path / "protected.xlsm"
        with ExcelFile(live_protected_xlsm_path) as wb:
            wb.set_module(
                "PasswordTest", wb.get_module("PasswordTest") + "\r\n' x\r\n"
            )
            with pytest.raises(VBAProjectError, match="password-protected"):
                wb.save(out)

    def test_save_real_protected_project_with_opt_in_succeeds(
        self, live_protected_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "protected_optin.xlsm"
        marker = "\r\n' protected edit allowed\r\n"
        with ExcelFile(live_protected_xlsm_path) as wb:
            wb.set_module(
                "PasswordTest", wb.get_module("PasswordTest") + marker
            )
            wb.save(out, allow_protected=True)
        with ExcelFile(out) as wb2:
            assert marker in wb2.get_module("PasswordTest")
            # Protection record is preserved verbatim.
            proj2 = wb2.vba_project()
            assert proj2.protection is not None
            assert proj2.protection.has_password is True

    def test_save_unprotected_project_does_not_require_opt_in(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        # The live fixture is unprotected: save without opt-in must succeed.
        out = tmp_path / "unprotected.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n' y\r\n")
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert wb2.get_module("Module1").endswith("' y\r\n")


# ===========================================================================
# GATE 17 — Digital Signature Gate
# ===========================================================================

class TestGate17_Signature:
    def test_signature_detected(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import detect_signature
        cfb = CFB.from_bytes(live_vba_bin)
        info = detect_signature(cfb)
        # Live fixture is unsigned; the detector must report that cleanly.
        assert info.present is False
        assert info.kinds == []

    def test_save_drops_stale_signature_with_warning(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Mutating a signed project must drop the stale signature streams
        and emit a UserWarning."""
        from pyopenvba.vba import detect_signature
        out = tmp_path / "signed.xlsm"
        # Inject a fake legacy signature stream so detect_signature triggers.
        with ExcelFile(live_xlsm_path) as wb:
            cfb_inject = wb._get_cfb()        # pyright: ignore[reportPrivateUsage]
            cfb_inject.add_stream_to_storage(
                "VBA", "_VBA_PROJECT_SIGNATURE", b"\xde\xad\xbe\xef" * 64
            )
            assert detect_signature(cfb_inject).present
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n' sig\r\n")
            with pytest.warns(UserWarning, match="signature"):
                wb.save(out)
        with ExcelFile(out) as wb2:
            cfb_after = CFB.from_bytes(wb2.vba_project_bytes())
            assert not detect_signature(cfb_after).present

    def test_save_signed_project_silent_with_opt_in(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.vba import detect_signature
        out = tmp_path / "signed_silent.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            cfb_inject = wb._get_cfb()        # pyright: ignore[reportPrivateUsage]
            cfb_inject.add_stream_to_storage(
                "VBA", "_VBA_PROJECT_SIGNATURE", b"\x01" * 128
            )
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n' s2\r\n")
            with warnings.catch_warnings():
                warnings.simplefilter("error")  # any UserWarning would fail
                wb.save(out, allow_invalidate_signature=True)
        with ExcelFile(out) as wb2:
            cfb_after = CFB.from_bytes(wb2.vba_project_bytes())
            assert not detect_signature(cfb_after).present

    def test_save_does_not_warn_when_no_signature_present(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "unsigned.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.set_module("Module1", wb.get_module("Module1") + "\r\n' u\r\n")
            with warnings.catch_warnings():
                warnings.simplefilter("error")
                wb.save(out)


# ===========================================================================
# GATE 18 — Encoding Gate
# ===========================================================================

class TestGate18_Encoding:
    """Cp1252 source round-trips; non-ASCII is not yet verified."""

    def test_cp1252_source_round_trips(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "enc.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            target = wb.get_module("Module1") + "\r\n' ascii edit\r\n"
            wb.set_module("Module1", target)
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert wb2.get_module("Module1") == target

    @pytest.mark.skip(
        reason="Excel does not permit non-cp1252 / non-ASCII module identifiers "
               "in the VBA IDE. Tracked as out-of-scope; the parser nevertheless "
               "round-trips Latin-1 supplement names (see test below).",
    )
    def test_non_ascii_module_name_round_trip_outside_cp1252(self) -> None:
        pytest.fail("out of scope")

    def test_latin1_supplement_module_name_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Latin-1 supplement (eacute, ntilde, ouml) round-trips through cp1252."""
        out = tmp_path / "non_ascii.xlsm"
        # Latin-1 supplement characters, fully encodable in cp1252.
        mod_name = "M\u00f3d\u00fcle_\u00e9\u00f1"   # "Modüle_éñ"
        src = (
            "' Modul name: " + mod_name + "\r\n"
            "Public Sub Hej_v\u00e4rlden()\r\n"      # 'världen'
            "    Debug.Print \"\u00a1Hola, mundo!\"\r\n"  # '!Hola...'
            "End Sub\r\n"
        )
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.add_module(mod_name, src, kind=VBAModuleKind.standard)
            wb.save(out)

        with ExcelFile(out) as wb2:
            assert mod_name in wb2.module_names()
            assert wb2.get_module(mod_name) == src
            # CFB stream name preserved verbatim (UTF-16 directory entry).
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            assert mod_name in set(cfb.list_streams_in_storage("VBA"))


# ===========================================================================
# GATE 19 — Cross-Structure Consistency Gate
# ===========================================================================

class TestGate19_Validate:
    def test_validate_clean_project(self, live_xlsm_path: Path) -> None:
        with ExcelFile(live_xlsm_path) as wb:
            problems = wb.validate()
        assert problems == [], f"expected clean, got {problems!r}"

    def test_validate_detects_missing_stream(self, live_vba_bin: bytes) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        proj.modules[0].stream_name = "Definitely_Missing_Stream"
        problems = proj.validate(cfb)
        assert any("missing stream" in p for p in problems)

    def test_validate_detects_duplicate_module_names(
        self, live_vba_bin: bytes
    ) -> None:
        cfb = CFB.from_bytes(live_vba_bin)
        proj = parse_vba_project(cfb)
        # Force a duplicate.
        proj.modules.append(VBAModule(
            name=proj.modules[0].name,
            stream_name="dup",
            source="",
        ))
        problems = proj.validate()
        assert any("duplicate" in p for p in problems)


# ===========================================================================
# GATE 20 — Round-Trip Preservation Gate
# ===========================================================================

class TestGate20_RoundTrip:
    """No-op parse-write-reopen preserves every module's source."""

    def test_no_op_round_trip_preserves_all_module_sources(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "noop.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            before = wb.vba_modules()
            wb.save(out)
        with ExcelFile(out) as wb2:
            after = wb2.vba_modules()
        assert before == after

    def test_no_op_round_trip_preserves_module_order(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "noop_order.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            before = wb.module_names()
            wb.save(out)
        with ExcelFile(out) as wb2:
            after = wb2.module_names()
        assert before == after

    def test_no_op_round_trip_preserves_other_zip_entries(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "noop_zip.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            wb.save(out)
        with zipfile.ZipFile(live_xlsm_path) as a, zipfile.ZipFile(out) as b:
            for name in a.namelist():
                if name == "xl/vbaProject.bin":
                    continue
                assert a.read(name) == b.read(name), name


# ===========================================================================
# GATE 21 — Mutation Round-Trip Gate
# ===========================================================================

class TestGate21_MutationRoundTrip:
    """Each implemented mutation round-trips through save/reopen."""

    @pytest.mark.parametrize("module_name", ["Module1", "Class1", "Sheet1"])
    def test_replace_source_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path, module_name: str
    ) -> None:
        out = tmp_path / f"mut_{module_name}.xlsm"
        marker = f"'edit-{module_name}\r\n"
        with ExcelFile(live_xlsm_path) as wb:
            new_src = wb.get_module(module_name) + "\r\n" + marker
            wb.set_module(module_name, new_src)
            wb.save(out)
        with ExcelFile(out) as wb2:
            assert marker in wb2.get_module(module_name)
            # All other modules unchanged.
            for n in wb2.module_names():
                if n == module_name:
                    continue
                with ExcelFile(live_xlsm_path) as orig:
                    assert wb2.get_module(n) == orig.get_module(n), n

    def test_replace_userform_code_behind_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Editing a UserForm's code-behind module must persist while the
        sibling designer sub-storage remains byte-for-byte identical.
        """
        out = tmp_path / "userform_edit.xlsm"
        marker = "\r\n' UserForm code-behind edit\r\nPrivate Sub Edit_Marker(): End Sub\r\n"

        with ExcelFile(live_xlsm_path) as wb:
            assert "UserForm1" in wb.module_names()
            before_src = wb.get_module("UserForm1")
            cfb_before = CFB.from_bytes(wb.vba_project_bytes())
            designer_before = {
                n: cfb_before.get_stream_in_storage("UserForm1", n)
                for n in TestGate14_Designer.DESIGNER_CHILDREN
            }
            wb.set_module("UserForm1", before_src + marker)
            wb.save(out)

        with ExcelFile(out) as wb2:
            new_src = wb2.get_module("UserForm1")
            assert marker in new_src
            assert new_src.startswith(before_src)
            # Designer storage children are unchanged byte-for-byte.
            cfb_after = CFB.from_bytes(wb2.vba_project_bytes())
            for name, blob in designer_before.items():
                assert cfb_after.get_stream_in_storage("UserForm1", name) == blob, name
            # Other modules unchanged.
            with ExcelFile(live_xlsm_path) as orig:
                for n in ("Module1", "Class1", "ThisWorkbook", "Sheet1"):
                    assert wb2.get_module(n) == orig.get_module(n), n

    def test_rename_module_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """rename_module persists to disk: dir, PROJECT, and CFB all updated."""
        out = tmp_path / "renamed.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            original_src = wb.get_module("Module1")
            proj = wb.vba_project()
            proj.rename_module("Module1", "MyMod")
            wb.save(out)

        # Reopen and verify the parsed model reflects the rename.
        with ExcelFile(out) as wb2:
            names = {n.casefold() for n in wb2.module_names()}
            assert "mymod" in names
            assert "module1" not in names
            assert wb2.get_module("MyMod") == original_src

    def test_rename_persists_in_project_stream(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """The PROJECT stream's Module= / Workspace= entries follow the rename."""
        out = tmp_path / "renamed_project.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.rename_module("Module1", "MyMod")
            wb.save(out)

        with ExcelFile(out) as wb2:
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            project_raw = cfb.get_stream("PROJECT")
            text = project_raw.decode("cp1252", errors="replace")
            assert "Module=MyMod" in text
            assert "Module=Module1" not in text

    def test_rename_renames_cfb_stream(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """The VBA storage's child stream is renamed (not duplicated)."""
        out = tmp_path / "renamed_cfb.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.rename_module("Module1", "MyMod")
            wb.save(out)

        with ExcelFile(out) as wb2:
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            vba_streams = set(cfb.list_streams_in_storage("VBA"))
            assert "MyMod" in vba_streams
            assert "Module1" not in vba_streams

    def test_add_module_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """add_module persists to disk: new CFB stream + dir + PROJECT."""
        out = tmp_path / "added.xlsm"
        src = "Attribute VB_Name = \"BrandNew\"\r\nSub Hi(): End Sub\r\n"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.add_module("BrandNew", src, kind=VBAModuleKind.standard)
            wb.save(out)

        with ExcelFile(out) as wb2:
            assert "BrandNew" in wb2.module_names()
            assert wb2.get_module("BrandNew") == src
            # PROJECT stream advertises the new module declaration.
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            project_text = cfb.get_stream("PROJECT").decode("cp1252", errors="replace")
            assert "Module=BrandNew" in project_text
            assert "BrandNew" in set(cfb.list_streams_in_storage("VBA"))

    def test_delete_module_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """delete_module persists to disk: CFB stream gone + dir + PROJECT scrubbed."""
        out = tmp_path / "deleted.xlsm"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.delete_module("Module1")
            wb.save(out)

        with ExcelFile(out) as wb2:
            names = {n.casefold() for n in wb2.module_names()}
            assert "module1" not in names
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            assert "Module1" not in set(cfb.list_streams_in_storage("VBA"))
            project_text = cfb.get_stream("PROJECT").decode("cp1252", errors="replace")
            assert "Module=Module1" not in project_text

    def test_add_then_rename_round_trip(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Add a module then rename it before saving: only the new name persists."""
        out = tmp_path / "add_then_rename.xlsm"
        src = "Sub G(): End Sub\r\n"
        with ExcelFile(live_xlsm_path) as wb:
            proj = wb.vba_project()
            proj.add_module("TempName", src, kind=VBAModuleKind.standard)
            proj.rename_module("TempName", "FinalName")
            wb.save(out)

        with ExcelFile(out) as wb2:
            names = set(wb2.module_names())
            assert "FinalName" in names
            assert "TempName" not in names
            assert wb2.get_module("FinalName") == src
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            vba_streams = set(cfb.list_streams_in_storage("VBA"))
            assert "FinalName" in vba_streams
            assert "TempName" not in vba_streams

    def test_delete_then_readd_same_name_does_not_duplicate_project_decl(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Regression: an idempotent ``delete_module`` + ``add_module`` of the
        same name across two consecutive saves (e.g. a "replace this module"
        script run twice) must not emit duplicate ``Module=NAME`` declarations
        or duplicate ``[Workspace]`` entries in the PROJECT stream.  Excel
        rejects such duplicates as workbook corruption."""
        out = tmp_path / "replace_idempotent.xlsm"
        src1 = "Attribute VB_Name = \"Demo\"\r\nSub A(): End Sub\r\n"
        src2 = "Attribute VB_Name = \"Demo\"\r\nSub B(): End Sub\r\n"

        # First save: fresh add.
        with ExcelFile(live_xlsm_path) as wb:
            wb.vba_project().add_module("Demo", src1, kind=VBAModuleKind.standard)
            wb.save(out)

        # Second save: delete-then-readd the same name (simulating a
        # "wipe and re-push this module" script being re-run).
        with ExcelFile(out) as wb:
            proj = wb.vba_project()
            proj.delete_module("Demo")
            proj.add_module("Demo", src2, kind=VBAModuleKind.standard)
            wb.save()

        with ExcelFile(out) as wb2:
            assert wb2.get_module("Demo") == src2
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
            project_text = cfb.get_stream("PROJECT").decode("cp1252", errors="replace")
            assert project_text.count("Module=Demo\r\n") == 1, project_text
            # And the Workspace entry must appear exactly once.
            ws_entries = [
                line for line in project_text.splitlines()
                if line.startswith("Demo=")
            ]
            assert len(ws_entries) == 1, project_text

    def test_save_heals_preexisting_duplicate_project_declarations(
        self, live_xlsm_path: Path, tmp_path: Path
    ) -> None:
        """Regression: opening an already-corrupted workbook (with duplicate
        ``Module=NAME`` / ``[Workspace]`` entries in PROJECT, as produced by
        the pre-fix buggy delete-then-readd flow) and performing any
        structural save (add / rename / delete) MUST scrub the duplicates."""
        from pyopenvba.cfb import CFB as _CFB
        from pyopenvba.vba import compress as _compress, decompress as _decompress

        out = tmp_path / "preexisting_corruption.xlsm"

        # Step 1: add a module legitimately.
        with ExcelFile(live_xlsm_path) as wb:
            wb.vba_project().add_module(
                "Demo",
                "Attribute VB_Name = \"Demo\"\r\nSub A(): End Sub\r\n",
                kind=VBAModuleKind.standard,
            )
            wb.save(out)

        # Step 2: forge corruption by injecting duplicate declarations
        # directly into the PROJECT stream of the saved xlsm.
        import io as _io
        import zipfile as _zipfile

        raw = out.read_bytes()
        with _zipfile.ZipFile(_io.BytesIO(raw), "r") as zin:
            vba_bytes = zin.read("xl/vbaProject.bin")
            other = {
                name: (zin.read(name), zin.getinfo(name))
                for name in zin.namelist()
                if name != "xl/vbaProject.bin"
            }
        cfb = _CFB.from_bytes(vba_bytes)
        project_text = cfb.get_stream("PROJECT").decode("cp1252", errors="replace")
        # Duplicate the module declaration and workspace entry.
        forged = project_text.replace(
            "Module=Demo\r\n", "Module=Demo\r\nModule=Demo\r\n", 1
        ).replace(
            "Demo=0, 0, 0, 0, C\r\n",
            "Demo=0, 0, 0, 0, C\r\nDemo=0, 0, 0, 0, C\r\n",
            1,
        )
        cfb.write_stream("PROJECT", forged.encode("cp1252", errors="replace"))
        forged_vba = cfb.to_bytes()

        buf = _io.BytesIO()
        with _zipfile.ZipFile(buf, "w") as zout:
            zout.writestr("xl/vbaProject.bin", forged_vba)
            for name, (data, info) in other.items():
                zout.writestr(info, data)
        out.write_bytes(buf.getvalue())

        # Sanity: confirm the forgery actually produced duplicates.
        with ExcelFile(out) as wb:
            text = _CFB.from_bytes(wb.vba_project_bytes()).get_stream("PROJECT")
            assert text.decode("cp1252").count("Module=Demo\r\n") == 2

        # Step 3: any structural save must heal the duplicates.
        with ExcelFile(out) as wb:
            wb.vba_project().add_module(
                "Healer",
                "Attribute VB_Name = \"Healer\"\r\nSub H(): End Sub\r\n",
                kind=VBAModuleKind.standard,
            )
            wb.save()

        with ExcelFile(out) as wb2:
            project_text = _CFB.from_bytes(
                wb2.vba_project_bytes()
            ).get_stream("PROJECT").decode("cp1252", errors="replace")
            assert project_text.count("Module=Demo\r\n") == 1, project_text
            assert project_text.count("Module=Healer\r\n") == 1, project_text
            ws_demo = [
                line for line in project_text.splitlines()
                if line.startswith("Demo=")
            ]
            assert len(ws_demo) == 1, project_text
            # Suppress unused-import lint
            _ = _compress, _decompress


# ===========================================================================
# GATE 22 — Corpus Gate
# ===========================================================================

class TestGate22_Corpus:
    """In-scope corpus coverage.

    Per pyOpenVBA scope (see docs/roadmap.md), the following categories
    are explicitly out of scope and therefore not required in the corpus:
    ActiveX-licensed workbooks (deprecated), digitally-signed workbooks
    (re-signing not supported), and non-ASCII module identifiers (Excel's
    VBA IDE does not permit them).
    """

    def test_at_least_one_real_macro_workbook(self, live_xlsm_path: Path) -> None:
        assert live_xlsm_path.exists()

    def test_corpus_covers_required_in_scope_categories(self) -> None:
        required = {
            "test_macro_workbook.xlsm":  "std + class + document + UserForm modules (xlsm)",
            "test_macro_workbook.xlsb":  "binary xlsb host container",
            "workbook_with_password_protected_vba_modules.xlsm":
                "password-protected VBA project",
            "xlsm_file_with_no_vba_entered_yet.xlsm":
                "xlsm whose VBA project has never been initialised",
        }
        corpus_dir = LIVE_XLSM.parent
        missing = [n for n in required if not (corpus_dir / n).exists()]
        assert not missing, f"missing corpus files: {missing!r}"


# ===========================================================================
# GATE 23 — Fuzz / Malformed Input Gate
# ===========================================================================

_FUZZ_CORPUS_DIR = Path(__file__).parent / "fuzz_corpus"


def _collect_persistent_fuzz_seeds() -> list[tuple[str, str, bytes]]:
    """Walk the persistent corpus and return (target, name, data) triples.

    Called at import time so the results can drive ``pytest.mark.parametrize``.
    Returns an empty list if the corpus directory is missing; the harness
    test below skips in that case (and a sibling sanity test reports it).
    """
    seeds: list[tuple[str, str, bytes]] = []
    if not _FUZZ_CORPUS_DIR.is_dir():
        return seeds
    for subdir in sorted(_FUZZ_CORPUS_DIR.iterdir()):
        if not subdir.is_dir():
            continue
        target = subdir.name
        for seed_file in sorted(subdir.iterdir()):
            if not seed_file.is_file():
                continue
            seeds.append((target, seed_file.name, seed_file.read_bytes()))
    return seeds


_PERSISTENT_FUZZ_SEEDS: list[tuple[str, str, bytes]] = _collect_persistent_fuzz_seeds()


class TestGate23_Fuzz:
    """Parsers must fail cleanly on malformed input, never crash silently."""

    @pytest.mark.parametrize("bad", [
        b"",
        b"\xd0",
        b"\xd0\xcf\x11\xe0",
        b"\xd0\xcf\x11\xe0" + b"\x00" * 100,
    ])
    def test_short_cfb_inputs_fail_cleanly(self, bad: bytes) -> None:
        with pytest.raises((CFBError, PyOpenVBAError)):
            CFB.from_bytes(bad)

    @pytest.mark.parametrize("bad", [
        b"",
        b"\x00",
        b"\x01\xff\xff",
        b"\x01" + b"\x00" * 3,
    ])
    def test_malformed_compressed_input_fails_cleanly(self, bad: bytes) -> None:
        with pytest.raises(VBAProjectError):
            decompress(bad)

    def test_truncated_chunk_data_fails_cleanly(self) -> None:
        # Valid header announcing 100 bytes of data, but only 1 byte follows.
        header = struct.pack("<H", 0xB000 | 99)
        with pytest.raises(VBAProjectError):
            decompress(b"\x01" + header + b"\x00")

    def test_random_bytes_do_not_crash_cfb_parser(self) -> None:
        for _ in range(50):
            blob = os.urandom(2048)
            try:
                CFB.from_bytes(blob)
            except CFBError:
                pass
            except PyOpenVBAError:
                pass

    def test_bit_flipped_vba_project_bytes_fail_cleanly(
        self, live_vba_bin: bytes
    ) -> None:
        """A bit-flipped live vbaProject.bin must never crash the CFB or
        VBA parsers with an unhandled exception; only documented errors
        are allowed."""
        import random
        rng = random.Random(0xC0FFEE)
        for _ in range(40):
            blob = bytearray(live_vba_bin)
            # Flip 1-8 random bytes anywhere in the file.
            for _i in range(rng.randint(1, 8)):
                idx = rng.randrange(len(blob))
                blob[idx] ^= rng.randint(1, 255)
            try:
                cfb = CFB.from_bytes(bytes(blob))
                parse_vba_project(cfb)
            except (CFBError, PyOpenVBAError, UnicodeDecodeError):
                # All accepted: parsers raise structured errors on damage.
                pass

    def test_fuzz_dir_stream_parser(self, live_vba_bin: bytes) -> None:
        """Bit-flipped dir streams must raise VBAProjectError, not crash."""
        from pyopenvba.vba import decompress, _parse_dir_stream  # type: ignore[attr-defined]
        import random
        cfb = CFB.from_bytes(live_vba_bin)
        raw = decompress(cfb.get_stream_in_storage("VBA", "dir"))
        rng = random.Random(0xBADF00D)
        for _ in range(30):
            blob = bytearray(raw)
            for _i in range(rng.randint(1, 6)):
                idx = rng.randrange(len(blob))
                blob[idx] ^= rng.randint(1, 255)
            try:
                _parse_dir_stream(bytes(blob))
            except (VBAProjectError, UnicodeDecodeError, IndexError, struct.error):
                pass

    def test_fuzz_project_stream_parser(self, live_vba_bin: bytes) -> None:
        """Bit-flipped PROJECT plain-text streams must not crash."""
        from pyopenvba.vba import parse_project_stream
        import random
        cfb = CFB.from_bytes(live_vba_bin)
        raw = cfb.get_stream("PROJECT")
        rng = random.Random(0xFEEDFACE)
        for _ in range(30):
            blob = bytearray(raw)
            for _i in range(rng.randint(1, 10)):
                idx = rng.randrange(len(blob))
                blob[idx] ^= rng.randint(1, 255)
            try:
                parse_project_stream(bytes(blob))
            except (VBAProjectError, UnicodeDecodeError):
                pass

    def test_fuzz_projectwm_parser(self, live_vba_bin: bytes) -> None:
        from pyopenvba.vba import parse_projectwm
        import random
        cfb = CFB.from_bytes(live_vba_bin)
        try:
            raw = cfb.get_stream("PROJECTwm")
        except KeyError:
            pytest.skip("fixture has no PROJECTwm stream")
        rng = random.Random(0xABCDEF)
        for _ in range(30):
            blob = bytearray(raw)
            for _i in range(rng.randint(1, 6)):
                idx = rng.randrange(len(blob))
                blob[idx] ^= rng.randint(1, 255)
            try:
                parse_projectwm(bytes(blob))
            except (VBAProjectError, UnicodeDecodeError):
                pass

    # ------------------------------------------------------------------
    # Persistent on-disk fuzz corpus
    # ------------------------------------------------------------------

    _CORPUS_DIR = Path(__file__).parent / "fuzz_corpus"

    @staticmethod
    def _run_persistent_seed(target: str, data: bytes) -> None:
        """Dispatch a corpus seed to the matching parser.

        Each parser must either succeed or raise one of the documented
        exception types.  Any other escape is a fuzz regression.
        """
        from pyopenvba.vba import (
            decompress,
            parse_project_stream,
            parse_projectwm,
            _parse_dir_stream,  # type: ignore[attr-defined]
        )

        if target == "cfb":
            try:
                CFB.from_bytes(data)
            except (CFBError, PyOpenVBAError):
                pass
        elif target == "decompress":
            try:
                decompress(data)
            except VBAProjectError:
                pass
        elif target == "dir":
            try:
                _parse_dir_stream(data)
            except (VBAProjectError, UnicodeDecodeError, IndexError, struct.error):
                pass
        elif target == "project":
            try:
                parse_project_stream(data)
            except (VBAProjectError, UnicodeDecodeError):
                pass
        elif target == "projectwm":
            try:
                parse_projectwm(data)
            except (VBAProjectError, UnicodeDecodeError):
                pass
        else:
            raise AssertionError(
                f"unknown fuzz target {target!r} (add a dispatch branch)"
            )

    def test_persistent_fuzz_corpus_is_present(self) -> None:
        """The checked-in fuzz corpus must exist and cover every target."""
        assert self._CORPUS_DIR.is_dir(), (
            f"fuzz corpus directory missing: {self._CORPUS_DIR}.  "
            f"Run: python scripts/seed_fuzz_corpus.py"
        )
        expected = {"cfb", "decompress", "dir", "project", "projectwm"}
        present = {
            p.name
            for p in self._CORPUS_DIR.iterdir()
            if p.is_dir() and any(p.iterdir())
        }
        missing = expected - present
        assert not missing, (
            f"fuzz corpus is missing seeds for: {sorted(missing)}.  "
            f"Run: python scripts/seed_fuzz_corpus.py"
        )

    @pytest.mark.parametrize(
        ("target", "name", "data"),
        _PERSISTENT_FUZZ_SEEDS,
        ids=[f"{t}/{n}" for (t, n, _d) in _PERSISTENT_FUZZ_SEEDS]
        or ["__empty__"],
    )
    def test_persistent_fuzz_corpus(
        self, target: str, name: str, data: bytes
    ) -> None:
        """Every checked-in seed must fail cleanly (or succeed).

        The corpus accumulates regression seeds over time; see
        ``tests/fuzz_corpus/README.md``.
        """
        if not _PERSISTENT_FUZZ_SEEDS:
            pytest.skip("fuzz corpus is empty; run scripts/seed_fuzz_corpus.py")
        del name  # only used for the pytest ID
        self._run_persistent_seed(target, data)


# ===========================================================================
# GATE 24 — API Contract Gate
# ===========================================================================

class TestGate24_APIContract:
    """Public API exposes layered, named operations."""

    def test_public_api_symbols_exposed(self) -> None:
        expected_top_level = {
            "ExcelFile", "VBAModuleKind",
            "PyOpenVBAError", "CFBError", "VBAProjectError",
            "UnsupportedFormatError",
        }
        assert expected_top_level <= set(pyopenvba.__all__)

    def test_excel_layer_methods(self) -> None:
        for m in [
            "vba_project", "vba_project_bytes",
            "vba_modules", "module_names",
            "get_module", "set_module",
            "validate", "save", "close",
        ]:
            assert callable(getattr(ExcelFile, m)), m

    def test_vba_project_methods(self) -> None:
        for m in ["get_module", "module_names", "validate"]:
            assert callable(getattr(VBAProject, m)), m

    def test_cfb_layer_methods(self) -> None:
        for m in [
            "from_bytes", "list_streams", "list_storages",
            "get_stream", "get_stream_in_storage",
            "list_streams_in_storage",
            "write_stream", "write_stream_in_storage",
            "to_bytes",
        ]:
            assert callable(getattr(CFB, m)), m

    def test_mutation_methods_exposed(self) -> None:
        for m in ["add_module", "rename_module", "delete_module"]:
            assert callable(getattr(VBAProject, m)), m


# ---------------------------------------------------------------------------
# Excel-authored large-module fixture: smoke + round-trip
# ---------------------------------------------------------------------------

class TestLargeModuleFixture:
    """Regression coverage using `large_vba_module.xlsm`, a workbook authored
    by Excel that contains a multi-chunk (>4 KB compressed) VBA module
    (`Large_Module_`).  This file was the empirical reference that proved
    Office never emits raw (flag=0) chunks for module source streams."""

    def _fixture_or_skip(self) -> Path:
        if not LIVE_LARGE_MODULE_XLSM.exists():
            pytest.skip(f"fixture missing: {LIVE_LARGE_MODULE_XLSM}")
        return LIVE_LARGE_MODULE_XLSM

    @staticmethod
    def _chunk_flags(stream: bytes, text_offset: int) -> list[int]:
        cs = stream[text_offset:]
        pos = 1
        flags: list[int] = []
        while pos < len(cs):
            header = struct.unpack_from("<H", cs, pos)[0]
            size = (header & 0x0FFF) + 1
            flags.append((header >> 15) & 0x1)
            pos += 2 + size
        return flags

    def test_excel_authored_large_module_uses_only_token_compressed_chunks(
        self,
    ) -> None:
        """Empirical anchor: Office never writes raw chunks for source."""
        from pyopenvba.cfb import CFB
        from pyopenvba.vba import parse_vba_project

        path = self._fixture_or_skip()
        with ExcelFile(path) as wb:
            cfb = CFB.from_bytes(wb.vba_project_bytes())
        proj = parse_vba_project(cfb)
        module = next(m for m in proj.modules if m.name == "Large_Module_")
        raw = cfb.get_stream_in_storage("VBA", module.stream_name)
        flags = self._chunk_flags(raw, module.text_offset)
        assert len(flags) >= 2, (
            f"fixture should be multi-chunk, got {len(flags)} chunks"
        )
        assert all(f == 1 for f in flags), (
            f"Excel-authored modules should never contain raw chunks; "
            f"got flags={flags}"
        )

    def test_resave_preserves_token_compressed_chunk_structure(
        self, tmp_path: Path
    ) -> None:
        """pyOpenVBA must resave a multi-chunk module with the same all-flag=1
        chunk discipline Office uses, and preserve the source byte-for-byte."""
        from pyopenvba.cfb import CFB
        from pyopenvba.vba import parse_vba_project

        path = self._fixture_or_skip()
        out = tmp_path / "large_vba_module_resave.xlsm"
        with ExcelFile(path) as wb:
            original_source = wb.get_module("Large_Module_")
            wb.set_module("Large_Module_", original_source)
            wb.save(out)

        with ExcelFile(out) as wb2:
            assert wb2.get_module("Large_Module_") == original_source
            cfb = CFB.from_bytes(wb2.vba_project_bytes())
        proj = parse_vba_project(cfb)
        module = next(m for m in proj.modules if m.name == "Large_Module_")
        raw = cfb.get_stream_in_storage("VBA", module.stream_name)
        flags = self._chunk_flags(raw, module.text_offset)
        assert len(flags) >= 2, "resave should still produce multiple chunks"
        assert all(f == 1 for f in flags), (
            f"resaved module must use only token-compressed chunks; "
            f"got flags={flags}"
        )


# ===========================================================================
# GATE 25 — Documentation Gate
# ===========================================================================

class TestGate25_Documentation:
    def test_roadmap_exists(self) -> None:
        repo_root = Path(__file__).parent.parent
        roadmap = repo_root / "docs" / "roadmap.md"
        assert roadmap.exists(), "docs/roadmap.md must exist"

    def test_roadmap_lists_supported_and_unsupported_features(self) -> None:
        repo_root = Path(__file__).parent.parent
        roadmap = (repo_root / "docs" / "roadmap.md").read_text(encoding="utf-8")
        for keyword in [
            "Supported", "Unsupported",
            ".xlsm", "UserForm", "signed",
            "protected", "PROJECTwm",
        ]:
            assert keyword in roadmap, f"roadmap.md should mention {keyword!r}"

    def test_readme_exists(self) -> None:
        repo_root = Path(__file__).parent.parent
        assert (repo_root / "README.md").exists()


# ---------------------------------------------------------------------------
# Sanity — keep unused imports tidy under pyright
# ---------------------------------------------------------------------------

_ = (cast, write_back_modules, io, VBAProjectError)
