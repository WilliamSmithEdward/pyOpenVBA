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


@pytest.fixture(scope="session")
def live_xlsm_path() -> Path:
    if not LIVE_XLSM.exists():
        pytest.skip(f"live test workbook not available at {LIVE_XLSM}")
    return LIVE_XLSM


@pytest.fixture(scope="session")
def live_vba_bin(live_xlsm_path: Path) -> bytes:
    with zipfile.ZipFile(live_xlsm_path) as zf:
        return zf.read("xl/vbaProject.bin")


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
    @pytest.mark.xfail(
        strict=True,
        reason="UserForm-bearing fixture not available; designer storage "
               "round-trip support is verbatim only.",
    )
    def test_designer_storage_preserved(self) -> None:
        pytest.fail("no UserForm test fixture")


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
        # The live fixture is unprotected.
        assert ps.protection.has_password is False


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

    @pytest.mark.xfail(
        strict=True,
        reason="No non-ASCII test fixture; round-trip of non-ASCII source/names "
               "is unverified.",
    )
    def test_non_ascii_module_name_round_trip(self) -> None:
        pytest.fail("no non-ASCII fixture")


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

    @pytest.mark.xfail(strict=True, reason="UserForm round-trip not exercised")
    def test_replace_userform_code_behind_round_trip(self) -> None:
        pytest.fail("no UserForm fixture")

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


# ===========================================================================
# GATE 22 — Corpus Gate
# ===========================================================================

class TestGate22_Corpus:
    """Test corpus coverage."""

    def test_at_least_one_real_macro_workbook(self, live_xlsm_path: Path) -> None:
        assert live_xlsm_path.exists()

    @pytest.mark.xfail(
        strict=True,
        reason="Corpus does not yet include UserForm, ActiveX, non-ASCII, "
               "password-protected, or signed workbooks.",
    )
    def test_corpus_covers_required_categories(self) -> None:
        required = [
            "empty_project.xlsm",
            "userform.xlsm",
            "activex.xlsm",
            "non_ascii_names.xlsm",
            "password_protected.xlsm",
            "signed.xlsm",
        ]
        corpus_dir = LIVE_XLSM.parent
        for name in required:
            assert (corpus_dir / name).exists(), f"missing corpus file: {name}"


# ===========================================================================
# GATE 23 — Fuzz / Malformed Input Gate
# ===========================================================================

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
