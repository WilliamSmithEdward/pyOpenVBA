"""Hunt for the `_VBA_PROJECT`-equivalent identifier table in Access.

The canonical Office `_VBA_PROJECT` stream begins with magic bytes
`CC 61` (the version is at offset 1). Search every LVAL row -- both
raw and OVBA-decompressed -- for that signature, then dump the prefix
of any hits.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile
from pyopenvba.vba import decompress

CORPUS = (
    Path(__file__).resolve().parents[1]
    / "tests" / "live_access_test" / "re_corpus" / "samples"
)

MAGIC_VBA_PROJECT = b"\xcc\x61"


def try_decompress(raw: bytes) -> bytes | None:
    """Attempt OVBA decompression. Return None on failure."""
    # OVBA blob starts with 0x01 sig byte. Try every plausible
    # starting offset.
    for start in range(min(len(raw), 64)):
        if raw[start] != 0x01:
            continue
        try:
            return decompress(raw, start)
        except Exception:
            continue
    return None


for sample in sorted(CORPUS.glob("*.accdb")):
    db = AccessFile(sample)
    rows = list(db._iter_lval_rows())  # type: ignore[attr-defined]
    hits_raw: list[tuple[int, int, int, int]] = []
    hits_decompressed: list[tuple[int, int, int, int]] = []
    for page, slot, row in rows:
        raw = bytes(row)
        pos = raw.find(MAGIC_VBA_PROJECT)
        if pos != -1:
            hits_raw.append((page, slot, pos, len(raw)))
        decomp = try_decompress(raw)
        if decomp is not None:
            pos = decomp.find(MAGIC_VBA_PROJECT)
            if pos != -1:
                hits_decompressed.append((page, slot, pos, len(decomp)))
    if hits_raw or hits_decompressed:
        print(f"\n=== {sample.name} ===")
        for h in hits_raw:
            page, slot, pos, n = h
            print(f"  RAW   page={page} slot={slot} pos={pos} row_len={n}")
            for page2, slot2, row in rows:
                if (page2, slot2) == (page, slot):
                    s = max(0, pos - 4)
                    e = min(len(row), pos + 96)
                    print(f"        bytes[{s}:{e}] = {bytes(row[s:e]).hex(' ')}")
        for h in hits_decompressed:
            page, slot, pos, n = h
            print(f"  OVBA  page={page} slot={slot} pos={pos} decomp_len={n}")
            for page2, slot2, row in rows:
                if (page2, slot2) == (page, slot):
                    decomp = try_decompress(bytes(row))
                    if decomp is not None:
                        s = max(0, pos - 4)
                        e = min(len(decomp), pos + 96)
                        print(f"        bytes[{s}:{e}] = {decomp[s:e].hex(' ')}")
