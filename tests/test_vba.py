"""Tests for the MS-OVBA compressor and decompressor."""

from __future__ import annotations

import struct

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

