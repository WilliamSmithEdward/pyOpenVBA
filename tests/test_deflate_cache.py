"""What deflate keeps for a part saved again: only pictures and other media, themes and styles, within a cap on the bytes it
holds, each found again only if it inflates to the same bytes. No other part is kept, as one may hold a password or a
connection string."""

from __future__ import annotations

import io
import random
import sys
import zipfile
import zlib

import pytest

from pyopenvba import _deflate
from pyopenvba._deflate import cached_bytes, compress, kept_raw_compress, raw_compress
from pyopenvba.powerquery._opc import OpcFile, kept_part


@pytest.fixture(autouse=True)
def an_empty_cache(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "_CACHE", type(_deflate._CACHE)())  # pyright: ignore[reportPrivateUsage]


def _part(seed: int, size: int) -> bytes:
    """Bytes that barely compress, so what is kept is about as large as what went in."""
    return random.Random(seed).randbytes(size)


def test_a_kept_part_saved_again_is_deflated_once_to_the_same_bytes() -> None:
    data = _part(1, 3000)
    first = kept_raw_compress(data)
    assert kept_raw_compress(bytes(data)) is first
    assert first == compress(data, level=6, wbits=15, memlevel=8)[2:-4] == raw_compress(data)
    assert zlib.decompress(first, -15) == data


def test_raw_compress_keeps_nothing() -> None:
    raw_compress(_part(2, 3000))
    assert cached_bytes() == 0


def test_the_cache_stays_under_its_cap_after_large_parts_are_saved_again(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "CACHE_LIMIT", 10_000)
    for round_ in range(3):
        for seed in range(6):
            kept_raw_compress(_part(seed, 3000))
            assert cached_bytes() <= 10_000, (round_, seed)
    assert len(_deflate._CACHE) == 3  # pyright: ignore[reportPrivateUsage]


def test_a_part_larger_than_the_cap_is_not_kept(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "CACHE_LIMIT", 1000)
    kept_raw_compress(_part(7, 3000))
    assert cached_bytes() == 0


def test_a_kept_part_whose_length_and_crc_match_other_bytes_is_not_taken() -> None:
    data = _part(8, 3000)
    _deflate._CACHE[(len(data), zlib.crc32(data))] = raw_compress(_part(9, 3000))  # pyright: ignore[reportPrivateUsage]
    assert zlib.decompress(kept_raw_compress(data), -15) == data


def test_the_cache_holds_no_input() -> None:
    data = _part(10, 3000)
    before = sys.getrefcount(data)
    body = kept_raw_compress(data)
    assert sys.getrefcount(data) == before
    assert list(_deflate._CACHE.values()) == [body]  # pyright: ignore[reportPrivateUsage]


def _package() -> OpcFile:
    held = io.BytesIO()
    with zipfile.ZipFile(held, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", "<Types/>")
    return OpcFile.parse(held.getvalue())


@pytest.mark.parametrize("name", ["xl/connections.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/vbaProject.bin",
                                  "customXml/item1.xml", "docProps/custom.xml"])
def test_a_part_that_may_hold_a_secret_is_never_kept(name: str) -> None:
    package = _package()
    package.write(name, b'<connection connection="Provider=x;Password=hunter2"/>' + _part(11, 2000))
    assert not kept_part(name)
    assert cached_bytes() == 0
    assert package.read(name).startswith(b'<connection connection="Provider=x;Password=hunter2"/>')


@pytest.mark.parametrize("name", ["xl/media/image1.png", "xl/theme/theme1.xml", "xl/styles.xml", "word/media/image2.jpeg"])
def test_media_themes_and_styles_are_kept(name: str) -> None:
    package = _package()
    data = _part(12, 2000)
    package.write(name, data)
    assert kept_part(name)
    assert cached_bytes() > 0
    assert package.read(name) == data
