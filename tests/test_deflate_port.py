"""The Python deflate port's bytes, held where they are.

Every saved part goes through pyopenvba._deflate, so a change to its loop must not change one byte. The digests below
were taken from the port after its bytes were checked against classic C zlib 1.3.2, the version it reproduces, at the
same settings. Where the running Python's zlib is that version, the test compares against it too. Elsewhere,
zlib-ng or another version, the digests alone hold the bytes.
"""
from __future__ import annotations

import hashlib
import zlib

import pytest

from pyopenvba._deflate import compress, raw_compress


def _noise(count: int, seed: int) -> bytes:
    """Bytes from a fixed linear congruential generator, the same on every Python."""
    out = bytearray()
    state = seed
    for _ in range(count):
        state = (state * 1103515245 + 12345) & 0x7FFFFFFF
        out.append(state >> 16 & 0xFF)
    return bytes(out)


_WORDS = [bytes(b"abcdefghij<>=\" /"[byte % 16] for byte in _noise(1 + seed % 11, seed)) for seed in range(300)]

#: Stored, fixed and dynamic blocks, long and short matches, the lazy matcher's TOO_FAR rule, runs past the 32 KB window
#: and the 8191-symbol block, and the rows a sheet formatted to its last row is made of.
CASES = {
    "empty": b"",
    "one": b"a",
    "noise": _noise(40000, 1),
    "repeat": b"the same line again and again\n" * 3000,
    "numbers": b"".join(f"{i},{i * 7919 % 100003}\n".encode() for i in range(20000)),
    "words": b" ".join(_WORDS[byte % 300] for byte in _noise(30000, 2)),
    "zeros": bytes(100000),
    "rows": b"".join(f'<row r="{i}" spans="1:1" s="{1 + i % 3}" customFormat="1"><c r="A{i}" s="2"/></row>'.encode()
                     for i in range(1, 6001)),
    "mixed": _noise(5000, 3) + b"abc" * 9000 + bytes(range(256)) * 50,
}
#: Excel's Power Query parts (level 6, memLevel 8), Access's attachments (level 5, memLevel 7), and the far settings.
SETTINGS = [(6, 15, 8), (5, 15, 7), (9, 15, 9), (4, 12, 4), (7, 9, 1)]

DIGESTS: dict[tuple[str, int, int, int], str] = {
    ("empty", 6, 15, 8): "aca503def43dc086",
    ("empty", 5, 15, 7): "4124a9b34347a1f8",
    ("empty", 9, 15, 9): "b171e283c6145acf",
    ("empty", 4, 12, 4): "ea5b548bd880c761",
    ("empty", 7, 9, 1): "1b582e8849445020",
    ("one", 6, 15, 8): "8ecacf17e3a7f86d",
    ("one", 5, 15, 7): "713ac2c2c11b299a",
    ("one", 9, 15, 9): "324b3f61040f09ab",
    ("one", 4, 12, 4): "49b893f08432e7c3",
    ("one", 7, 9, 1): "3d50deb35b880254",
    ("noise", 6, 15, 8): "1244122f05cc299f",
    ("noise", 5, 15, 7): "6c051e89ed5cb1e3",
    ("noise", 9, 15, 9): "68b43899df8c86ee",
    ("noise", 4, 12, 4): "7091b87d0b825ccb",
    ("noise", 7, 9, 1): "05c813401424aba2",
    ("repeat", 6, 15, 8): "563decbf7a186b5e",
    ("repeat", 5, 15, 7): "afcf7ecfe156f71d",
    ("repeat", 9, 15, 9): "eaf0955630b48533",
    ("repeat", 4, 12, 4): "e334eddefac8e5c0",
    ("repeat", 7, 9, 1): "3b0b148acc70ee8a",
    ("numbers", 6, 15, 8): "321d552bd728b787",
    ("numbers", 5, 15, 7): "4cf2c193808adca3",
    ("numbers", 9, 15, 9): "39c188eddad8767a",
    ("numbers", 4, 12, 4): "64dc38198abff2ee",
    ("numbers", 7, 9, 1): "89db42f52278a0a0",
    ("words", 6, 15, 8): "88ea5952b328feaf",
    ("words", 5, 15, 7): "88c15b7dcacd5b65",
    ("words", 9, 15, 9): "3a402692a4d09b4f",
    ("words", 4, 12, 4): "7d1fe54e7ca7f739",
    ("words", 7, 9, 1): "64276a4f2be9ee78",
    ("zeros", 6, 15, 8): "a2ea4b930bf1fb9b",
    ("zeros", 5, 15, 7): "29063f0e6b702e9d",
    ("zeros", 9, 15, 9): "345c550fa604b669",
    ("zeros", 4, 12, 4): "7589a1e3adfe4ac5",
    ("zeros", 7, 9, 1): "ae4f3e9f2cfe50b5",
    ("rows", 6, 15, 8): "54d57f378e083c9c",
    ("rows", 5, 15, 7): "bf9dc6edf30cd58c",
    ("rows", 9, 15, 9): "cc33b2da19c379df",
    ("rows", 4, 12, 4): "5b1f572cbe3712a5",
    ("rows", 7, 9, 1): "98af2d671be4c36e",
    ("mixed", 6, 15, 8): "15cccc5275f38797",
    ("mixed", 5, 15, 7): "686f15cd3a5ead5a",
    ("mixed", 9, 15, 9): "0d2afc7becf8c5b2",
    ("mixed", 4, 12, 4): "4f6d73a9e4a37535",
    ("mixed", 7, 9, 1): "7d1c3ebc37a796b3",
}


@pytest.mark.parametrize("name", list(CASES))
def test_the_port_writes_the_bytes_it_always_has(name: str) -> None:
    data = CASES[name]
    for level, wbits, memlevel in SETTINGS:
        stream = compress(data, level, wbits, memlevel)
        assert zlib.decompress(stream, wbits) == data
        assert hashlib.sha256(stream).hexdigest()[:16] == DIGESTS[(name, level, wbits, memlevel)], (level, wbits, memlevel)
    assert raw_compress(data) == compress(data, 6, 15, 8)[2:-4]


@pytest.mark.skipif(zlib.ZLIB_RUNTIME_VERSION != "1.3.2", reason="the port reproduces classic zlib 1.3.2")
@pytest.mark.parametrize("name", list(CASES))
def test_the_port_writes_what_classic_zlib_writes(name: str) -> None:
    data = CASES[name]
    for level, wbits, memlevel in SETTINGS:
        z = zlib.compressobj(level, zlib.DEFLATED, wbits, memlevel)
        assert compress(data, level, wbits, memlevel) == z.compress(data) + z.flush(), (level, wbits, memlevel)
