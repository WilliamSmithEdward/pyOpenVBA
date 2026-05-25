"""Crack the ProcTrailer cookie (`7B 02 <u32 LE>`) field.

For each corpus sample, extract the cookie value and compare against
candidate hash/length functions of the surrounding bytecode. Goal: find
a deterministic algorithm that matches every sample.
"""
from __future__ import annotations

import binascii
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"


def cookie_of(raw: bytes) -> tuple[int, int]:
    # Find the LAST 7B 02 sequence (the procedure trailer, distinct
    # from any earlier inline use).
    idx = raw.rfind(b"\x7b\x02")
    if idx < 0:
        return -1, -1
    return idx, int.from_bytes(raw[idx + 2:idx + 6], "little")


def all_cookies() -> list[tuple[str, int, int, int, bytes]]:
    rows: list[tuple[str, int, int, int, bytes]] = []
    for path in sorted(CORPUS.glob("*.accdb")):
        try:
            stream = AccessFile(path).read_module_pcode_stream()
        except Exception as e:
            print(f"  SKIP {path.name}: {e}")
            continue
        pos, cookie = cookie_of(stream.raw)
        rows.append((path.name, len(stream.raw), pos, cookie, stream.raw))
    return rows


def main() -> None:
    rows = all_cookies()
    print(f"{'sample':45s}  {'len':>4s}  {'pos':>4s}  {'cookie':>10s}")
    print("-" * 80)
    for name, length, pos, cookie, _ in rows:
        print(f"{name:45s}  {length:>4d}  0x{pos:02x}  0x{cookie:08x}")
    print()

    # Candidate 1: CRC32 of raw[:pos] (i.e. bytecode up to but not
    # including the trailer).
    print("--- candidate: CRC32(raw[:pos]) low 32 bits ---")
    for name, _, pos, cookie, raw in rows:
        crc = binascii.crc32(raw[:pos]) & 0xFFFFFFFF
        ok = "MATCH" if crc == cookie else "    "
        print(f"  {name:45s}  cookie=0x{cookie:08x}  crc=0x{crc:08x}  {ok}")
    print()

    # Candidate 2: CRC32 of raw[:pos+2] (include the 7B 02 opcode).
    print("--- candidate: CRC32(raw[:pos+2]) ---")
    for name, _, pos, cookie, raw in rows:
        crc = binascii.crc32(raw[:pos + 2]) & 0xFFFFFFFF
        ok = "MATCH" if crc == cookie else "    "
        print(f"  {name:45s}  cookie=0x{cookie:08x}  crc=0x{crc:08x}  {ok}")
    print()

    # Candidate 3: low 16 bits of CRC32 (since some cookies look small).
    print("--- candidate: CRC32(raw[:pos]) & 0xFFFF ---")
    for name, _, pos, cookie, raw in rows:
        crc = binascii.crc32(raw[:pos]) & 0xFFFF
        ok = "MATCH" if crc == (cookie & 0xFFFF) else "    "
        print(f"  {name:45s}  cookie={cookie & 0xFFFF:#06x}  crc={crc:#06x}  {ok}")
    print()

    # Candidate 4: cookie may be unrelated to bytecode -- maybe it's a
    # body-region byte offset. Try: offset of the LAST ProcEnd
    # (67 02 00 00 00 00) from the start.
    print("--- candidate: offset of last `67 02 00 00 00 00` ProcEnd ---")
    for name, _, _, cookie, raw in rows:
        pe = raw.rfind(b"\x67\x02\x00\x00\x00\x00")
        ok = "MATCH" if pe == cookie else "    "
        print(f"  {name:45s}  cookie={cookie:#010x}  procend@={pe:#010x}  {ok}")
    print()

    # Candidate 5: maybe cookie is the byte count of the Sub body
    # (between the opening `67 02 06 00 00 00` and the closing
    # `67 02 00 00 00 00`).
    print("--- candidate: body byte-length (between ProcOpen and ProcEnd) ---")
    for name, _, _, cookie, raw in rows:
        po = raw.find(b"\x67\x02\x06\x00\x00\x00")
        pe = raw.rfind(b"\x67\x02\x00\x00\x00\x00")
        body_len = (pe - (po + 6)) if (po >= 0 and pe > po) else -1
        ok = "MATCH" if body_len == cookie else "    "
        print(f"  {name:45s}  cookie={cookie:#010x}  body_len={body_len:#010x}  {ok}")


if __name__ == "__main__":
    main()
