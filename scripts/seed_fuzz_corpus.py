"""Seed the persistent fuzz corpus under ``tests/fuzz_corpus/``.

Idempotent and additive: never deletes existing files; only writes a
seed if its target filename does not yet exist. Hand-curated regression
seeds are therefore safe across re-runs.

Usage:
    python scripts/seed_fuzz_corpus.py
"""

from __future__ import annotations

import random
import struct
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "src"))

from pyopenvba.cfb import CFB  # noqa: E402
from pyopenvba.vba import compress  # noqa: E402

CORPUS = REPO / "tests" / "fuzz_corpus"
LIVE_VBA_BIN = REPO / "tests" / "live_excel_testing" / "test_macro_workbook.xlsm"


def _write(target: Path, data: bytes) -> bool:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        return False
    target.write_bytes(data)
    return True


def _load_live_vba_bin() -> bytes | None:
    """Return the raw vbaProject.bin from the live xlsm fixture, or None."""
    if not LIVE_VBA_BIN.exists():
        return None
    import zipfile
    with zipfile.ZipFile(LIVE_VBA_BIN) as z:
        return z.read("xl/vbaProject.bin")


def _bit_flipped_variants(
    base: bytes, count: int, seed: int, min_flips: int = 1, max_flips: int = 8
) -> list[bytes]:
    rng = random.Random(seed)
    out: list[bytes] = []
    for _ in range(count):
        blob = bytearray(base)
        if not blob:
            continue
        for _i in range(rng.randint(min_flips, max_flips)):
            idx = rng.randrange(len(blob))
            blob[idx] ^= rng.randint(1, 255)
        out.append(bytes(blob))
    return out


def seed_cfb() -> int:
    target_dir = CORPUS / "cfb"
    written = 0
    seeds: dict[str, bytes] = {
        "empty.bin": b"",
        "one_byte.bin": b"\xd0",
        "partial_signature.bin": b"\xd0\xcf\x11\xe0",
        "signature_plus_zeros.bin": b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 100,
        "truncated_header.bin": b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 24,
        "random_2k_seed_1.bin": random.Random(1).randbytes(2048),
        "random_2k_seed_2.bin": random.Random(2).randbytes(2048),
        "random_512_seed_3.bin": random.Random(3).randbytes(512),
    }
    # Add a few bit-flipped variants of the live vba_project bin if available.
    live = _load_live_vba_bin()
    if live is not None:
        for i, blob in enumerate(_bit_flipped_variants(live, 4, seed=0xC0FFEE)):
            seeds[f"live_bitflip_{i:02d}.bin"] = blob

    for name, data in seeds.items():
        if _write(target_dir / name, data):
            written += 1
    return written


def seed_decompress() -> int:
    target_dir = CORPUS / "decompress"
    written = 0
    # decompress() expects a 1-byte signature + chunked stream.
    seeds: dict[str, bytes] = {
        "empty.bin": b"",
        "one_zero.bin": b"\x00",
        "wrong_signature.bin": b"\x02\x00\xb0\xff",
        "signature_only.bin": b"\x01",
        "signature_short_header.bin": b"\x01\x00",
        "header_announces_overlong_chunk.bin": b"\x01" + struct.pack("<H", 0xB000 | 99) + b"\x00",
        "compressed_then_garbage.bin": compress(b"Hello world!") + b"\xff\xff\xff\xff",
    }
    for name, data in seeds.items():
        if _write(target_dir / name, data):
            written += 1
    return written


def _live_dir_raw() -> bytes | None:
    live = _load_live_vba_bin()
    if live is None:
        return None
    cfb = CFB.from_bytes(live)
    from pyopenvba.vba import decompress
    return decompress(cfb.get_stream_in_storage("VBA", "dir"))


def seed_dir() -> int:
    target_dir = CORPUS / "dir"
    written = 0
    seeds: dict[str, bytes] = {
        "empty.bin": b"",
        "one_byte.bin": b"\x01",
        "truncated_record_header.bin": b"\x01\x00\x04\x00\x00\x00",
        "version_record_truncated.bin": b"\x09\x00\x02\x00\x00\x00\xff",
    }
    raw = _live_dir_raw()
    if raw is not None:
        for i, blob in enumerate(
            _bit_flipped_variants(raw, 6, seed=0xBADF00D, min_flips=1, max_flips=6)
        ):
            seeds[f"live_bitflip_{i:02d}.bin"] = blob
    for name, data in seeds.items():
        if _write(target_dir / name, data):
            written += 1
    return written


def _live_project_raw() -> bytes | None:
    live = _load_live_vba_bin()
    if live is None:
        return None
    cfb = CFB.from_bytes(live)
    return cfb.get_stream("PROJECT")


def seed_project() -> int:
    target_dir = CORPUS / "project"
    written = 0
    seeds: dict[str, bytes] = {
        "empty.bin": b"",
        "garbage_text.bin": b"<<not a project stream>>\r\n",
        "unterminated_section_header.bin": b"[Host Extender Info\r\n",
        "only_workspace.bin": b"[Workspace]\r\nSheet1=0, 0, 100, 100, C\r\n",
        "binary_in_text.bin": b"ID=\"foo\"\r\n\x00\x01\x02\x03\r\n",
    }
    raw = _live_project_raw()
    if raw is not None:
        for i, blob in enumerate(
            _bit_flipped_variants(raw, 6, seed=0xFEEDFACE, min_flips=1, max_flips=10)
        ):
            seeds[f"live_bitflip_{i:02d}.bin"] = blob
    for name, data in seeds.items():
        if _write(target_dir / name, data):
            written += 1
    return written


def _live_projectwm_raw() -> bytes | None:
    live = _load_live_vba_bin()
    if live is None:
        return None
    cfb = CFB.from_bytes(live)
    try:
        return cfb.get_stream_in_storage("VBA", "PROJECTwm")
    except KeyError:
        return None


def seed_projectwm() -> int:
    target_dir = CORPUS / "projectwm"
    written = 0
    seeds: dict[str, bytes] = {
        "empty.bin": b"",
        "single_null.bin": b"\x00",
        "mbcs_only_no_terminator.bin": b"Module1",
        "mbcs_terminator_no_unicode.bin": b"Module1\x00",
        "mbcs_terminator_truncated_unicode.bin": b"Module1\x00M",
    }
    raw = _live_projectwm_raw()
    if raw is not None:
        for i, blob in enumerate(
            _bit_flipped_variants(raw, 6, seed=0xABCDEF, min_flips=1, max_flips=6)
        ):
            seeds[f"live_bitflip_{i:02d}.bin"] = blob
    for name, data in seeds.items():
        if _write(target_dir / name, data):
            written += 1
    return written


def main() -> int:
    counts = {
        "cfb": seed_cfb(),
        "decompress": seed_decompress(),
        "dir": seed_dir(),
        "project": seed_project(),
        "projectwm": seed_projectwm(),
    }
    total_new = sum(counts.values())
    print(f"Wrote {total_new} new seed file(s) under {CORPUS}:")
    for k, v in counts.items():
        existing = len(list((CORPUS / k).glob("*.bin"))) if (CORPUS / k).exists() else 0
        print(f"  {k:11s}  new={v:3d}  total={existing}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
