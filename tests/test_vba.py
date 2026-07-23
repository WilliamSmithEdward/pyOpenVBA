"""Tests for the MS-OVBA compressor and decompressor."""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from pyopenvba.exceptions import VBAProjectError
from pyopenvba.vba import (
    CLASS_MODULE_CLSID,
    VBAModuleKind,
    VBAProject,
    compress,
    copy_token_help,
    decompress,
    normalize_class_source,
    split_attribute_header,
    synthesize_class_header,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_raw_chunk(data: bytes) -> bytes:
    """Build one raw (uncompressed) chunk.  data must be exactly 4096 bytes."""
    assert len(data) == 4096
    # header: flag=0, sig=0b011, size-1=0x0FFF → 0x3FFF
    return struct.pack("<H", 0x3FFF) + data


def _make_literal_chunk(data: bytes) -> bytes:
    """Build one token-compressed chunk using only literal tokens."""
    assert len(data) <= 3640, "literal-only encoding may exceed 4096 for larger chunks"
    payload = bytearray()
    i = 0
    while i < len(data):
        group = data[i: i + 8]
        payload.append(0x00)    # all-literal flag byte
        payload.extend(group)
        i += len(group)
    # header: flag=1, sig=0b011, size-1 = len(payload)-1 → 0xB000 | (size-1)
    header = 0xB000 | (len(payload) - 1)
    return struct.pack("<H", header) + bytes(payload)


def _make_container(*chunks: bytes) -> bytes:
    return b"\x01" + b"".join(chunks)


# ---------------------------------------------------------------------------
# _copy_token_help
# ---------------------------------------------------------------------------

class TestCopyTokenHelp:
    """[MS-OVBA] 2.4.1.3.6 — verify the bit-count / mask table."""

    def _bit_count(self, difference: int) -> int:
        _, _, bc = copy_token_help(difference, 0)
        return bc

    def test_difference_1_gives_bit_count_4(self) -> None:
        assert self._bit_count(1) == 4

    def test_difference_16_gives_bit_count_4(self) -> None:
        assert self._bit_count(16) == 4

    def test_difference_17_gives_bit_count_5(self) -> None:
        assert self._bit_count(17) == 5

    def test_difference_32_gives_bit_count_5(self) -> None:
        assert self._bit_count(32) == 5

    def test_difference_33_gives_bit_count_6(self) -> None:
        assert self._bit_count(33) == 6

    def test_difference_4096_gives_bit_count_12(self) -> None:
        assert self._bit_count(4096) == 12

    def test_masks_are_complementary(self) -> None:
        for diff in (1, 16, 17, 256, 4096):
            lm, om, _ = copy_token_help(diff, 0)
            assert (lm & om) == 0
            assert (lm | om) == 0xFFFF

    def test_offset_bits_grow_with_difference(self) -> None:
        """More bytes already decoded → more bits needed for offset."""
        _, _, bc1 = copy_token_help(16, 0)
        _, _, bc2 = copy_token_help(4096, 0)
        assert bc2 > bc1


# ---------------------------------------------------------------------------
# Decompressor
# ---------------------------------------------------------------------------

class TestDecompress:
    def test_missing_signature_raises(self) -> None:
        with pytest.raises(VBAProjectError, match="signature"):
            decompress(b"\x02some data")

    def test_empty_bytes_raises(self) -> None:
        with pytest.raises(VBAProjectError, match="signature"):
            decompress(b"")

    def test_empty_payload(self) -> None:
        """Signature byte only — valid empty container."""
        result = decompress(b"\x01")
        assert result == b""

    def test_literal_round_trip_small(self) -> None:
        original = b"Hello, VBA world!"
        container = _make_container(_make_literal_chunk(original))
        result = decompress(container)
        assert result[:len(original)] == original

    def test_literal_round_trip_all_zeros(self) -> None:
        original = b"\x00" * 64
        result = decompress(_make_container(_make_literal_chunk(original)))
        assert result[:64] == original

    def test_raw_chunk_round_trip(self) -> None:
        original = bytes(range(256)) * 16    # exactly 4096 bytes
        result = decompress(_make_container(_make_raw_chunk(original)))
        assert result == original

    def test_bad_chunk_signature_raises(self) -> None:
        # Bits 12-14 must be 0b011 = 3; use 0b101 = 5 → header with sig=5
        bad_header = (5 << 12) | 0x0000   # flag=0, sig=5
        data = b"\x01" + struct.pack("<H", bad_header) + b"\x00" * 4096
        with pytest.raises(VBAProjectError, match="signature"):
            decompress(data)

    def test_raw_chunk_wrong_size_raises(self) -> None:
        # flag=0 but size field says 10 bytes (not 4096)
        bad_header = 0x3000 | 9   # flag=0, sig=0b011, size-1=9
        data = b"\x01" + struct.pack("<H", bad_header) + b"\x00" * 10
        with pytest.raises(VBAProjectError, match=r"[Rr]aw chunk"):
            decompress(data)

    def test_two_consecutive_literal_chunks(self) -> None:
        chunk1 = b"AAAA"
        chunk2 = b"BBBB"
        container = _make_container(
            _make_literal_chunk(chunk1),
            _make_literal_chunk(chunk2),
        )
        result = decompress(container)
        assert result[:4] == chunk1
        assert result[4:8] == chunk2


# ---------------------------------------------------------------------------
# Compressor
# ---------------------------------------------------------------------------

class TestCompress:
    def test_empty_input(self) -> None:
        result = compress(b"")
        assert result == b"\x01"

    def test_round_trip_small(self) -> None:
        original = b"Sub Hello()\nMsgBox \"Hi\"\nEnd Sub\n"
        assert decompress(compress(original)) == original

    def test_round_trip_all_zeros(self) -> None:
        original = b"\x00" * 100
        assert decompress(compress(original)) == original

    def test_round_trip_exact_4096(self) -> None:
        original = bytes(range(256)) * 16   # exactly 4096 bytes
        assert decompress(compress(original)) == original

    def test_round_trip_4097_bytes(self) -> None:
        original = b"A" * 4096 + b"B"
        assert decompress(compress(original)) == original

    def test_round_trip_repetitive(self) -> None:
        original = b"ABCD" * 1000
        compressed = compress(original)
        assert decompress(compressed) == original
        # Repetitive data should compress well.
        assert len(compressed) < len(original)

    def test_signature_byte_present(self) -> None:
        assert compress(b"hello")[0] == 0x01

    def test_final_chunk_not_raw_unless_exact_4096(self) -> None:
        """Non-4096-byte final chunks must NOT be raw chunks (would add null padding)."""
        data = b"X" * 100
        compressed = compress(data)
        # After the signature byte, read the first chunk header.
        header = struct.unpack_from("<H", compressed, 1)[0]
        chunk_flag = (header >> 15) & 0x1
        assert chunk_flag == 1, "partial final chunk must be token-compressed, not raw"

    def test_full_chunk_emitted_as_token_compressed_not_raw(self) -> None:
        """Regression: full 4096-byte chunks of realistic VBA source must be
        emitted as token-compressed (flag=1), not raw (flag=0).  Excel's VBE
        rejects modules whose CompressedSourceCode begins with a raw chunk
        with "An error occurred while loading <Module>", even though raw
        chunks are spec-legal.  Office itself never emits raw chunks for
        module source, so we must match that behaviour."""
        # Construct a 5000-byte VBA-like source that spans two chunks.
        line = b"    Debug.Print \"line %04d -- pyOpenVBA round-trip\"\r\n"
        body = b"".join(line.replace(b"%04d", f"{i:04d}".encode()) for i in range(100))
        data = b"Sub LongRoutine()\r\n" + body + b"End Sub\r\n"
        assert len(data) > 4096, "test fixture must span more than one chunk"

        compressed = compress(data)
        # Walk all chunks and confirm every one is flag=1.
        pos = 1
        chunk_flags: list[int] = []
        while pos < len(compressed):
            header = struct.unpack_from("<H", compressed, pos)[0]
            size = (header & 0x0FFF) + 1
            chunk_flags.append((header >> 15) & 0x1)
            pos += 2 + size
        assert chunk_flags and all(f == 1 for f in chunk_flags), (
            f"every emitted chunk must be token-compressed; got flags={chunk_flags}"
        )
        # Round-trip must still hold.
        assert decompress(compressed) == data

    def test_long_module_round_trip_through_excel_save(
        self, tmp_path: Path
    ) -> None:
        """Regression for the bug that crashed VBE on long fresh-add modules.

        Adding a module whose compressed source spans more than one 4096-byte
        chunk (e.g. the ``DemoShowcase`` push demo) used to emit a raw first
        chunk that Excel rejected as "An error occurred while loading
        <Module>".  This test exercises the full add-then-save-then-reopen
        round-trip with such a module."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import VBAModuleKind

        # Use the same live fixture as the rest of the gate tests.
        live = (
            Path(__file__).resolve().parent
            / "live_excel_testing"
            / "test_macro_workbook.xlsm"
        )
        if not live.exists():
            import pytest as _pytest
            _pytest.skip("live xlsm fixture not available")

        # Realistic VBA source > 4096 bytes after encoding.
        body_lines = [
            f"    Cells({i + 1}, 1).Value = \"row {i + 1} -- pyOpenVBA test\"\r\n"
            for i in range(150)
        ]
        src = (
            "Attribute VB_Name = \"LongMod\"\r\n"
            "Sub LongRoutine()\r\n"
            + "".join(body_lines)
            + "End Sub\r\n"
        )
        assert len(src.encode("cp1252")) > 4096

        out = tmp_path / "long_module.xlsm"
        with ExcelFile(live) as wb:
            wb.vba_project().add_module(
                "LongMod", src, kind=VBAModuleKind.standard
            )
            wb.save(out)

        with ExcelFile(out) as wb2:
            assert wb2.get_module("LongMod") == src
            # Inspect the on-disk stream: every chunk must be flag=1.
            from pyopenvba.cfb import CFB as _CFB
            from pyopenvba.vba import parse_vba_project as _parse
            cfb = _CFB.from_bytes(wb2.vba_project_bytes())
            proj = _parse(cfb)
            module = next(m for m in proj.modules if m.name == "LongMod")
            raw = cfb.get_stream_in_storage("VBA", module.stream_name)
            cs = raw[module.text_offset:]
            pos = 1
            flags: list[int] = []
            while pos < len(cs):
                header = struct.unpack_from("<H", cs, pos)[0]
                size = (header & 0x0FFF) + 1
                flags.append((header >> 15) & 0x1)
                pos += 2 + size
            assert len(flags) >= 2, f"expected multi-chunk stream, got {flags}"
            assert all(f == 1 for f in flags), (
                f"every CompressedSourceCode chunk must be token-compressed; "
                f"got flags={flags}"
            )


# ---------------------------------------------------------------------------
# Attribute header preservation (VBE-style body-only edits)
# ---------------------------------------------------------------------------

class TestAttributeHeaderPreservation:
    """Body-only source edits must preserve the leading Attribute VB_* block.

    Document modules (ThisWorkbook, Sheet1, ...) carry host-binding
    attributes in their source.  Stripping them on a source replacement
    breaks the workbook in Excel.
    """

    def test_split_attribute_header_standard(self) -> None:
        from pyopenvba.vba import split_attribute_header
        src = 'Attribute VB_Name = "Module1"\r\nSub Foo()\r\nEnd Sub\r\n'
        header, body = split_attribute_header(src)
        assert header == 'Attribute VB_Name = "Module1"\r\n'
        assert body == "Sub Foo()\r\nEnd Sub\r\n"

    def test_split_attribute_header_document_with_blank_separator(self) -> None:
        from pyopenvba.vba import split_attribute_header
        src = (
            'Attribute VB_Name = "ThisWorkbook"\r\n'
            'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\n'
            'Attribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\n'
            'Attribute VB_Customizable = True\r\n'
            '\r\n'
            'Sub Hello()\r\nEnd Sub\r\n'
        )
        header, body = split_attribute_header(src)
        assert header.endswith('VB_Customizable = True\r\n\r\n')
        assert body == "Sub Hello()\r\nEnd Sub\r\n"

    def test_split_attribute_header_class_with_version_block(self) -> None:
        from pyopenvba.vba import split_attribute_header
        src = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            'Attribute VB_Name = "Class1"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\n'
            '\r\n'
            "Sub Bar()\r\nEnd Sub\r\n"
        )
        header, body = split_attribute_header(src)
        assert header.startswith("VERSION 1.0 CLASS\r\n")
        assert 'Attribute VB_Name = "Class1"' in header
        assert body == "Sub Bar()\r\nEnd Sub\r\n"

    def test_split_attribute_header_no_header(self) -> None:
        from pyopenvba.vba import split_attribute_header
        header, body = split_attribute_header("Sub Foo()\r\nEnd Sub\r\n")
        assert header == ""
        assert body == "Sub Foo()\r\nEnd Sub\r\n"

    def test_set_module_body_only_preserves_document_header(
        self, tmp_path: Path
    ) -> None:
        """`set_module('ThisWorkbook', body)` must keep the Attribute header."""
        from pyopenvba.excel import ExcelFile

        src_path = Path("tests/live_excel_testing/workbook_only_module_test.xlsm")
        if not src_path.exists():
            pytest.skip("requires workbook_only_module_test.xlsm fixture")
        out = tmp_path / "wb_body_only.xlsm"
        import shutil
        shutil.copy(src_path, out)

        with ExcelFile(out) as wb:
            original = wb.get_module("ThisWorkbook")
            assert original.startswith('Attribute VB_Name = "ThisWorkbook"')
            wb.set_module(
                "ThisWorkbook",
                'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n',
            )
            wb.save()

        with ExcelFile(out) as wb2:
            new_src = wb2.get_module("ThisWorkbook")
            # Header preserved verbatim.
            assert new_src.startswith('Attribute VB_Name = "ThisWorkbook"')
            assert 'Attribute VB_PredeclaredId = True' in new_src
            assert 'Attribute VB_Customizable = True' in new_src
            # New body appears after the header.
            assert 'Sub Hello()' in new_src
            assert 'MsgBox "hi"' in new_src
            # Old body is gone.
            assert 'ThisWorkbookTest' not in new_src

    def test_set_module_full_source_replacement_still_works(
        self, tmp_path: Path
    ) -> None:
        """An explicit full-source replacement (with Attribute header) is honored as-is."""
        from pyopenvba.excel import ExcelFile

        src_path = Path("tests/live_excel_testing/workbook_only_module_test.xlsm")
        if not src_path.exists():
            pytest.skip("requires workbook_only_module_test.xlsm fixture")
        out = tmp_path / "wb_full_replace.xlsm"
        import shutil
        shutil.copy(src_path, out)

        full = (
            'Attribute VB_Name = "ThisWorkbook"\r\n'
            'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\n'
            'Attribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\n'
            'Attribute VB_Customizable = True\r\n'
            '\r\n'
            'Sub Custom()\r\nEnd Sub\r\n'
        )
        with ExcelFile(out) as wb:
            wb.set_module("ThisWorkbook", full)
            wb.save()

        with ExcelFile(out) as wb2:
            assert wb2.get_module("ThisWorkbook") == full

    def test_add_module_standard_synthesizes_header(self, tmp_path: Path) -> None:
        """add_module(kind=standard) prepends Attribute VB_Name when missing."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import VBAModuleKind

        src_path = Path("tests/live_excel_testing/test_macro_workbook.xlsm")
        if not src_path.exists():
            pytest.skip("requires test_macro_workbook.xlsm fixture")
        out = tmp_path / "synth_header.xlsm"
        import shutil
        shutil.copy(src_path, out)

        with ExcelFile(out) as wb:
            proj = wb.vba_project()
            proj.add_module(
                "BodyOnly",
                "Sub Foo()\r\n    MsgBox \"foo\"\r\nEnd Sub\r\n",
                kind=VBAModuleKind.standard,
            )
            wb.save()

        with ExcelFile(out) as wb2:
            src = wb2.get_module("BodyOnly")
            assert src.startswith('Attribute VB_Name = "BodyOnly"\r\n')
            assert 'Sub Foo()' in src

    def test_add_module_standard_honors_supplied_header(self, tmp_path: Path) -> None:
        """If caller supplies an Attribute header, add_module uses it as-is."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import VBAModuleKind

        src_path = Path("tests/live_excel_testing/test_macro_workbook.xlsm")
        if not src_path.exists():
            pytest.skip("requires test_macro_workbook.xlsm fixture")
        out = tmp_path / "supplied_header.xlsm"
        import shutil
        shutil.copy(src_path, out)

        full = (
            'Attribute VB_Name = "FullSrc"\r\n'
            'Sub Bar()\r\nEnd Sub\r\n'
        )
        with ExcelFile(out) as wb:
            wb.vba_project().add_module(
                "FullSrc", full, kind=VBAModuleKind.standard
            )
            wb.save()

        with ExcelFile(out) as wb2:
            assert wb2.get_module("FullSrc") == full

    def test_add_module_other_synthesizes_class_header(self, tmp_path: Path) -> None:
        """add_module(kind=other) synthesizes a standard class header when none is supplied."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import CLASS_MODULE_CLSID, VBAModuleKind

        src_path = Path("tests/live_excel_testing/test_macro_workbook.xlsm")
        if not src_path.exists():
            pytest.skip("requires test_macro_workbook.xlsm fixture")
        import shutil
        out = tmp_path / "class_synth.xlsm"
        shutil.copy(src_path, out)

        with ExcelFile(out) as wb:
            proj = wb.vba_project()
            proj.add_module(
                "MyClass",
                "Option Explicit\r\n",
                kind=VBAModuleKind.other,
            )
            wb.save()

        with ExcelFile(out) as wb2:
            src = wb2.get_module("MyClass")
            assert src.startswith('Attribute VB_Name = "MyClass"\r\n')
            assert f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"' in src
            assert 'Attribute VB_GlobalNameSpace = False' in src
            assert 'Attribute VB_PredeclaredId = False' in src
            assert 'Option Explicit' in src

    def test_add_module_other_honors_supplied_header(self, tmp_path: Path) -> None:
        """add_module(kind=other) uses a caller-supplied header as-is."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import CLASS_MODULE_CLSID, VBAModuleKind

        src_path = Path("tests/live_excel_testing/test_macro_workbook.xlsm")
        if not src_path.exists():
            pytest.skip("requires test_macro_workbook.xlsm fixture")
        import shutil
        out = tmp_path / "class_explicit.xlsm"
        shutil.copy(src_path, out)

        full = (
            f'Attribute VB_Name = "ExplicitClass"\r\n'
            f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = False\r\n'
            'Attribute VB_Exposed = False\r\n'
            'Attribute VB_TemplateDerived = False\r\n'
            'Attribute VB_Customizable = False\r\n'
            'Option Explicit\r\n'
        )
        with ExcelFile(out) as wb:
            wb.vba_project().add_module(
                "ExplicitClass", full, kind=VBAModuleKind.other
            )
            wb.save()

        with ExcelFile(out) as wb2:
            assert wb2.get_module("ExplicitClass") == full

    def test_module_body_property_round_trip(self, tmp_path: Path) -> None:
        """``VBAModule.body`` exposes the source minus its attribute header."""
        from pyopenvba.excel import ExcelFile

        src_path = Path("tests/live_excel_testing/workbook_only_module_test.xlsm")
        if not src_path.exists():
            pytest.skip("requires workbook_only_module_test.xlsm fixture")
        out = tmp_path / "body_prop.xlsm"
        import shutil
        shutil.copy(src_path, out)

        with ExcelFile(out) as wb:
            proj = wb.vba_project()
            mod = proj.get_module("ThisWorkbook")
            assert "Attribute VB_Name" not in mod.body
            assert mod.body.startswith("Sub ThisWorkbookTest()")
            mod.body = "Sub Replaced()\r\nEnd Sub\r\n"
            assert mod.source.startswith('Attribute VB_Name = "ThisWorkbook"')
            assert "Sub Replaced()" in mod.source
            wb.save()

        with ExcelFile(out) as wb2:
            src = wb2.get_module("ThisWorkbook")
            assert src.startswith('Attribute VB_Name = "ThisWorkbook"')
            assert "Sub Replaced()" in src
            assert "ThisWorkbookTest" not in src


def test_compress_short_input_emits_lz_not_literal_only():
    """Regression for Phase 5g: short inputs (<= 3640 bytes) must
    be processed by the LZ encoder, not a literal-only fast path.

    Access byte-validates the OVBA cache blob and rejects modules
    whose compressed bytes don't match its own compressor's output,
    even when the decompressed plaintext is identical. The literal-
    only path is MS-OVBA-spec-compliant and round-trips correctly
    but produces all-zero flag bytes, which Access never emits for
    inputs containing 3+ byte repeats.
    """
    # Plaintext with an obvious back-referenceable repeat.
    plain = b"\r\n\r\n" * 4
    out = compress(plain)
    assert decompress(out) == plain
    assert out[0] == 0x01
    header = int.from_bytes(out[1:3], "little")
    payload_len = (header & 0x0FFF) + 1
    payload = out[3:3 + payload_len]
    # Scan flag bytes for any non-zero (copy-token) bit.
    pos = 0
    saw_copy_token = False
    while pos < len(payload):
        flag = payload[pos]
        pos += 1
        if flag != 0:
            saw_copy_token = True
            break
        pos += 8  # 8 literal bytes
    assert saw_copy_token, (
        f"compress() emitted literal-only output for repeatable input; "
        f"Access UI will reject this. payload={payload.hex()}"
    )


def test_compress_byte_exact_against_access_sample_040():
    """Pin the byte-exact MS-OVBA compressor parity with Access.

    Decompresses sample 040 module M's OVBA cache blob, recompresses
    the resulting plaintext through pyopenvba.vba.compress, and
    requires byte-for-byte identity with the original blob. This is
    the invariant that gates Access UI acceptance of rewritten
    modules (Phase 5g).
    """
    from pyopenvba.access_read import AccessReader

    sample = (
        Path(__file__).resolve().parent
        / "live_access_test"
        / "re_corpus"
        / "samples"
        / "040__sub_msgbox_hello.accdb"
    )
    if not sample.exists():
        pytest.skip(f"RE corpus sample missing: {sample}")
    db = AccessReader(sample)
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        for off in db._scan_ovba_signatures(row):  # pyright: ignore[reportPrivateUsage]
            blob = bytes(row)[off:]
            try:
                plain = decompress(blob)
            except Exception:
                continue
            if not plain.startswith(b'Attribute VB_Name = "M"'):
                continue
            recomp = compress(plain)
            assert recomp == blob, (
                f"compressor diverges from Access at (page={page}, "
                f"slot={slot}, off={off}). orig={len(blob)}B "
                f"ours={len(recomp)}B"
            )
            return
    pytest.skip("module M not located in sample 040")


# ---------------------------------------------------------------------------
# Class-source normalization (GitHub issue #1)
# ---------------------------------------------------------------------------
#
# VBE-exported .cls files are in *file-export form*: a leading
# ``VERSION 1.0 CLASS`` / ``BEGIN`` / ``END`` preamble and NO
# ``Attribute VB_Base`` line.  Module streams need *stream form*: no
# VERSION preamble, WITH VB_Base.  Both differences are independently
# fatal in live Excel (missing VB_Base: "Invalid procedure call or
# argument" at the first ``New`` site; VERSION preamble in the stream:
# "Compile error: Expected: end of statement" on line 1 of the class).

_VB_BASE_LINE = f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"'

# The reporter's fixture shape from issue #1: own attribute header,
# no VB_Base, LF line endings.
_ISSUE1_HEADERED_CLS_LF = (
    'Attribute VB_Name = "Class1"\n'
    "Attribute VB_GlobalNameSpace = False\n"
    "Attribute VB_Creatable = False\n"
    "Attribute VB_PredeclaredId = False\n"
    "Attribute VB_Exposed = False\n"
    "\n"
    "Private mMessage As String\n"
    "\n"
    "Public Function Greet() As String\n"
    '    Greet = "Class1.Greet: " & mMessage\n'
    "End Function\n"
)

# A genuine VBE export: VERSION preamble + the same header, CRLF.
_VBE_EXPORT_CLS_CRLF = (
    "VERSION 1.0 CLASS\r\n"
    "BEGIN\r\n"
    "  MultiUse = -1  'True\r\n"
    "END\r\n"
    'Attribute VB_Name = "Class1"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = False\r\n"
    "Attribute VB_Exposed = False\r\n"
    "\r\n"
    "Private mMessage As String\r\n"
)


class TestNormalizeClassSource:
    def test_inserts_vb_base_after_vb_name_preserving_lf(self) -> None:
        out = normalize_class_source(_ISSUE1_HEADERED_CLS_LF)
        lines = out.split("\n")
        assert lines[0] == 'Attribute VB_Name = "Class1"'
        assert lines[1] == _VB_BASE_LINE
        assert "\r\n" not in out
        # Body untouched.
        assert out.endswith("End Function\n")

    def test_strips_version_preamble_and_inserts_vb_base_crlf(self) -> None:
        out = normalize_class_source(_VBE_EXPORT_CLS_CRLF)
        assert not out.startswith("VERSION")
        assert "MultiUse" not in out
        lines = out.split("\r\n")
        assert lines[0] == 'Attribute VB_Name = "Class1"'
        assert lines[1] == _VB_BASE_LINE
        assert out.endswith("Private mMessage As String\r\n")

    def test_stream_form_input_is_unchanged(self) -> None:
        stream_form = synthesize_class_header("Class1") + "\r\nPrivate x As Long\r\n"
        assert normalize_class_source(stream_form) == stream_form

    def test_idempotent(self) -> None:
        once = normalize_class_source(_VBE_EXPORT_CLS_CRLF)
        assert normalize_class_source(once) == once

    def test_prior_header_vb_base_wins_over_universal_clsid(self) -> None:
        host_base = 'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"'
        prior = (
            'Attribute VB_Name = "ThisWorkbook"\r\n'
            f"{host_base}\r\n"
            "Attribute VB_Exposed = True\r\n"
        )
        supplied = (
            'Attribute VB_Name = "ThisWorkbook"\r\n'
            "Attribute VB_Exposed = True\r\n"
            "\r\n"
            "Private Sub Workbook_Open()\r\nEnd Sub\r\n"
        )
        out = normalize_class_source(supplied, prior_header=prior)
        assert host_base in out
        assert CLASS_MODULE_CLSID not in out

    def test_supplied_vb_base_is_kept(self) -> None:
        custom = 'Attribute VB_Base = "0{DEADBEEF-0000-0000-C000-000000000046}"'
        src = (
            'Attribute VB_Name = "C"\r\n'
            f"{custom}\r\n"
            "\r\n"
            "Private x As Long\r\n"
        )
        assert normalize_class_source(src) == src

    def test_version_preamble_with_bare_body_yields_bare_body(self) -> None:
        src = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            "Private x As Long\r\n"
        )
        assert normalize_class_source(src) == "Private x As Long\r\n"

    def test_header_without_vb_name_anchor_is_left_alone(self) -> None:
        src = (
            "Attribute VB_GlobalNameSpace = False\r\n"
            "\r\n"
            "Private x As Long\r\n"
        )
        assert normalize_class_source(src) == src

    def test_empty_source(self) -> None:
        assert normalize_class_source("") == ""


class TestAddModuleClassNormalization:
    def test_headered_cls_without_vb_base_is_normalized(self) -> None:
        project = VBAProject()
        module = project.add_module(
            "Class1", _ISSUE1_HEADERED_CLS_LF, kind=VBAModuleKind.other
        )
        header, _ = split_attribute_header(module.source)
        assert _VB_BASE_LINE in header
        assert module.attribute_header == header

    def test_vbe_export_form_is_normalized(self) -> None:
        project = VBAProject()
        module = project.add_module(
            "Class1", _VBE_EXPORT_CLS_CRLF, kind=VBAModuleKind.other
        )
        assert not module.source.startswith("VERSION")
        header, _ = split_attribute_header(module.source)
        assert _VB_BASE_LINE in header

    def test_version_preamble_with_bare_body_gets_synthesized_header(self) -> None:
        src = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            "Private x As Long\r\n"
        )
        project = VBAProject()
        module = project.add_module("Class1", src, kind=VBAModuleKind.other)
        assert module.attribute_header == synthesize_class_header("Class1")
        assert module.source.endswith("Private x As Long\r\n")

    def test_standard_module_with_header_is_untouched(self) -> None:
        src = (
            'Attribute VB_Name = "Module1"\r\n'
            "\r\n"
            "Sub Hello()\r\nEnd Sub\r\n"
        )
        project = VBAProject()
        module = project.add_module("Module1", src, kind=VBAModuleKind.standard)
        assert module.source == src
        assert CLASS_MODULE_CLSID not in module.source


class TestEncodeLzOracleEquivalence:
    """The optimized 3-gram-indexed LZ encoder must be byte-for-byte
    equivalent to the naive full-window scan it replaced.  Access
    byte-validates OVBA cache blobs against its own compressor, so any
    output drift (including tie-break order) is a correctness bug, not
    a quality regression."""

    @staticmethod
    def _naive_encode_lz(chunk: bytes) -> bytes:
        """Reference implementation: the original O(window) scan."""
        out = bytearray()
        pos = 0
        chunk_len = len(chunk)
        while pos < chunk_len:
            flag_bits = 0
            tokens: list[bytes] = []
            for bit in range(8):
                if pos >= chunk_len:
                    break
                length_mask, offset_mask, bit_count = copy_token_help(pos, 0)
                max_length = length_mask + 3
                max_offset = (offset_mask >> (16 - bit_count)) + 1
                start = max(0, pos - max_offset)
                best_len = 0
                best_offset = 0
                for candidate in range(start, pos):
                    match_len = 0
                    while (pos + match_len < chunk_len
                           and chunk[candidate + match_len] == chunk[pos + match_len]
                           and match_len < max_length):
                        match_len += 1
                    if match_len > best_len:
                        best_len = match_len
                        best_offset = pos - candidate
                if best_len >= 3:
                    flag_bits |= (1 << bit)
                    offset_bits = ((best_offset - 1) << (16 - bit_count)) & offset_mask
                    length_bits = (best_len - 3) & length_mask
                    tokens.append(struct.pack("<H", offset_bits | length_bits))
                    pos += best_len
                else:
                    tokens.append(bytes([chunk[pos]]))
                    pos += 1
            out.append(flag_bits)
            for tok in tokens:
                out.extend(tok)
        return bytes(out)

    def _assert_equivalent(self, chunk: bytes) -> None:
        from pyopenvba import vba as _vba

        encode = getattr(_vba, "_encode_lz")
        assert encode(chunk) == self._naive_encode_lz(chunk), (
            f"encoder diverges from naive oracle for input of "
            f"{len(chunk)} bytes"
        )

    def test_boundary_sizes(self) -> None:
        for chunk in (b"", b"A", b"AB", b"ABC", b"ABCA", b"AAAA"):
            self._assert_equivalent(chunk)

    def test_random_bytes(self) -> None:
        import random

        rng = random.Random(20260722)
        self._assert_equivalent(bytes(rng.randrange(256) for _ in range(4096)))
        rng = random.Random(1)
        self._assert_equivalent(bytes(rng.randrange(64) for _ in range(2048)))

    def test_highly_repetitive(self) -> None:
        self._assert_equivalent(b"A" * 512)
        self._assert_equivalent(b"AB" * 256)
        self._assert_equivalent(b"abc" * 170 + b"ab")
        self._assert_equivalent(b"\x00" * 300 + b"\x00\x01" * 100)

    def test_vba_like_text(self) -> None:
        src = (
            'Attribute VB_Name = "Module1"\r\n'
            "Option Explicit\r\n\r\n"
            + "".join(
                f"Sub Proc{i}()\r\n"
                f'    MsgBox "value {i}"\r\n'
                "End Sub\r\n\r\n"
                for i in range(30)
            )
        ).encode("cp1252")
        self._assert_equivalent(src[:4096])

    def test_repeating_grams_with_capped_matches(self) -> None:
        # Many identical 3-grams whose matches hit the per-position
        # length cap: exercises the early-exit tie-break path.
        self._assert_equivalent(bytes(range(16)) * 128)
