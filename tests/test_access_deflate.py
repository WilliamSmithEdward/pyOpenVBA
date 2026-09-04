"""The deflate Access compresses attachments with: classic zlib at level 5,
memLevel 7 and a 32 KB window, carried over into Python."""

from __future__ import annotations

import random
import zlib
from pathlib import Path

from pyopenvba.access import AccessDatabase
from pyopenvba.access._complex import decode_file_data, encode_file_data
from pyopenvba.access._deflate import compress

FIXTURE = Path(__file__).parent / "live_access_test" / "complex_columns.accdb"


def test_the_header_names_a_fast_level_and_the_trailer_is_adler32() -> None:
    stream = compress(b"hello world")
    assert stream[:2] == bytes.fromhex("785e")
    assert (stream[0] * 256 + stream[1]) % 31 == 0
    assert stream[-4:] == (zlib.adler32(b"hello world") & 0xFFFFFFFF).to_bytes(4, "big")
    assert zlib.decompress(stream) == b"hello world"


def test_the_fixture_attachments_access_wrote_are_reproduced_byte_for_byte() -> None:
    """The fixture's attachments were stored by Access itself; encoding
    their decoded contents again gives the stored container exactly,
    deflate stream included."""
    db = AccessDatabase(FIXTURE)
    column = next(c for c in db.complex_columns() if c.column == "Files")
    checked = 0
    for row in db.table(column.flat_table).rows():
        stored = row["FileData"]
        assert isinstance(stored, bytes)
        extension, data = decode_file_data(stored)
        assert encode_file_data(extension, data) == stored
        checked += 1
    assert checked >= 3


def test_every_block_kind_and_a_sliding_window_round_trip() -> None:
    """Stored blocks for random bytes, a fixed tree for a few bytes, dynamic
    trees for text, several blocks for an input past the 8191-symbol block
    and past the 32 KB window: each inflates to what went in."""
    random.seed(5)
    cases = [
        b"",
        b"a",
        b"abc",
        bytes(random.randrange(256) for _ in range(40000)),
        b"the same line again and again\n" * 3000,
        b"".join(f"{i},{random.random():.6f}\n".encode() for i in range(20000)),
    ]
    for data in cases:
        stream = compress(data)
        assert zlib.decompress(stream) == data
        # Whatever the block kinds, the stream is at most a stored copy plus
        # the framing: a header, a trailer, and five bytes per block, where
        # a block holds at most 8191 symbols.
        assert len(stream) <= len(data) + 6 + 5 * (len(data) // 8191 + 2)


def test_the_block_boundary_falls_after_8191_symbols() -> None:
    """memLevel 7 gives an 8192-symbol buffer, flushed one short of full;
    a run of distinct literals crosses that boundary and comes out as two
    blocks that still inflate whole."""
    random.seed(9)
    data = bytes(random.randrange(256) for _ in range(8191)) + b"tail" * 50
    stream = compress(data)
    assert zlib.decompress(stream) == data
