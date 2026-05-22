"""Tests for the MS-OVBA compressor and decompressor."""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from pyopenvba.vba import compress, decompress, copy_token_help
from pyopenvba.exceptions import VBAProjectError


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
        with pytest.raises(VBAProjectError, match="[Rr]aw chunk"):
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

    def test_add_module_other_without_header_rejected(self) -> None:
        """Class/document modules cannot be invented without an explicit header."""
        from pyopenvba.excel import ExcelFile
        from pyopenvba.vba import VBAModuleKind

        src_path = Path("tests/live_excel_testing/test_macro_workbook.xlsm")
        if not src_path.exists():
            pytest.skip("requires test_macro_workbook.xlsm fixture")
        with ExcelFile(src_path) as wb:
            proj = wb.vba_project()
            with pytest.raises(ValueError, match="attribute header"):
                proj.add_module(
                    "InventedClass",
                    "Sub Baz()\r\nEnd Sub\r\n",
                    kind=VBAModuleKind.other,
                )

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




