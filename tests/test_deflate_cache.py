"""raw_compress keeps what it deflated, by the input's digest, within a cap on the bytes it holds."""

from __future__ import annotations

import gc
import hashlib
import sys
import zlib

import pytest

from pyopenvba import _deflate
from pyopenvba._deflate import cached_bytes, compress, raw_compress


@pytest.fixture(autouse=True)
def an_empty_cache(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "_CACHE", type(_deflate._CACHE)())  # pyright: ignore[reportPrivateUsage]


def _part(seed: int, size: int) -> bytes:
    """Bytes that barely compress, so what is kept is about as large as what went in."""
    return b"".join(hashlib.sha256(f"{seed}:{n}".encode()).digest() for n in range(size // 32 + 1))[:size]


def test_a_part_saved_again_is_deflated_once_to_the_same_bytes() -> None:
    data = _part(1, 3000)
    first = raw_compress(data)
    assert raw_compress(bytes(data)) is first
    assert first == compress(data, level=6, wbits=15, memlevel=8)[2:-4]
    assert zlib.decompress(first, -15) == data


def test_the_cache_stays_under_its_cap_after_large_parts_are_saved_again(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "CACHE_LIMIT", 10_000)
    for round_ in range(3):
        for seed in range(6):
            raw_compress(_part(seed, 3000))
            assert cached_bytes() <= 10_000, (round_, seed)
    assert len(_deflate._CACHE) == 3  # pyright: ignore[reportPrivateUsage]


def test_a_part_larger_than_the_cap_is_not_kept(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(_deflate, "CACHE_LIMIT", 1000)
    raw_compress(_part(7, 3000))
    assert cached_bytes() == 0


def test_the_cache_holds_no_input() -> None:
    data = _part(9, 3000)
    before = sys.getrefcount(data)
    body = raw_compress(data)
    gc.collect()
    assert sys.getrefcount(data) == before
    assert list(_deflate._CACHE.values()) == [body]  # pyright: ignore[reportPrivateUsage]
    assert all(len(key[0]) == 32 for key in _deflate._CACHE)  # pyright: ignore[reportPrivateUsage]
