"""A version 4 compound file: 4096-byte sectors.

The header is still 512 bytes but takes the whole first sector, so sector n
starts at (n + 1) * 4096 ([MS-CFB] 2.2), not 512 + n * 4096 as version 3's
512-byte sectors do.

The fixture was written by Windows itself, StgCreateStorageEx with
STGOPTIONS SectorSize 4096 (through pywin32), so it is the reference
implementation's layout rather than a reading of the specification. It holds
a stream below the mini-stream cutoff, one above it, and a storage with two
more; each stream holds ``_pattern(seed, length)``.
"""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import CFBError

_FIXTURE = Path(__file__).parent / "fixtures" / "version4_storage.cfb"


def _pattern(seed: int, length: int) -> bytes:
    """The bytes the fixture's streams were written with; they differ sector to sector."""
    return bytes((seed * 31 + i * 7 + (i >> 9)) & 0xFF for i in range(length))


_TOP = {"Small": _pattern(1, 100), "Big": _pattern(2, 10000)}
_IN_VBA = {"dir": _pattern(3, 5000), "Module1": _pattern(4, 300)}


def _check(cfb: CFB) -> None:
    for name, data in _TOP.items():
        assert cfb.get_stream(name) == data, name
    for name, data in _IN_VBA.items():
        assert cfb.get_stream_in_storage("VBA", name) == data, f"VBA/{name}"


def test_fixture_is_version_4() -> None:
    data = _FIXTURE.read_bytes()
    assert struct.unpack_from("<HH", data, 26)[0] == 4
    assert struct.unpack_from("<H", data, 30)[0] == 12


def test_reads_every_stream_in_the_mini_stream_and_in_regular_sectors() -> None:
    _check(CFB.from_bytes(_FIXTURE.read_bytes()))


def test_is_written_back_as_version_3_with_the_same_streams() -> None:
    written = CFB.from_bytes(_FIXTURE.read_bytes()).to_bytes()
    assert struct.unpack_from("<H", written, 26)[0] == 3
    _check(CFB.from_bytes(written))


@pytest.mark.parametrize(("major", "shift"), [(3, 12), (4, 9)])
def test_refuses_a_sector_size_the_version_does_not_have(major: int, shift: int) -> None:
    data = bytearray(_FIXTURE.read_bytes())
    struct.pack_into("<H", data, 26, major)
    struct.pack_into("<H", data, 30, shift)
    with pytest.raises(CFBError, match="sector sizes"):
        CFB.from_bytes(bytes(data))
