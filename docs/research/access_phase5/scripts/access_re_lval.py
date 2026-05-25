"""Phase-2 RE: decode LVAL row format.

A 4 KiB LVAL page header is 28 bytes:
    [0]      page_type 0x01
    [1]      0x01 (subtype?)
    [2:4]    checksum or version
    [4:8]    'LVAL'
    [8:12]   reserved (zero on all observed pages)
    [12:14]  u16 LE: slot count N (used+tombstoned)
    [14:14+2N]   slot table: u16 LE per slot. Top nibble flags:
                  0xD = tombstone (no row), low 12 bits = ?
                  others = byte offset of the row in the page

Rows grow downward from the top of the page; their end byte is the
start byte of the next slot (descending offsets), or PAGE_SIZE for the
top-most slot.

This script lists every LVAL page in an .accdb, decodes the slot table,
and prints each row's offset/length plus a small hex/ASCII preview of
both the row head AND the row tail (where the continuation pointer
lives in chained LVAL records).
"""
from __future__ import annotations
import sys
from pathlib import Path

PAGE = 4096
HDR = 28


def _ascii(b: bytes) -> str:
    return "".join(chr(c) if 32 <= c < 127 else "." for c in b)


def decode_page(data: bytes, page_num: int) -> None:
    base = page_num * PAGE
    if data[base] != 0x01 or data[base + 4 : base + 8] != b"LVAL":
        return
    n = int.from_bytes(data[base + 12 : base + 14], "little")
    print(f"\n== page {page_num} @0x{base:06X}, slots={n} ==")
    print(f"   hdr bytes[8:14]: {data[base+8:base+14].hex(' ')}")
    slots: list[int] = []
    for i in range(n):
        off = int.from_bytes(data[base + 14 + 2 * i : base + 16 + 2 * i], "little")
        slots.append(off)
    print(f"   slot table: {[f'0x{s:04X}' for s in slots]}")

    # Sort slots descending by offset to compute row spans.
    # Tombstones have top nibble 0xD; their "offset" is a flag value.
    real = [(i, s) for i, s in enumerate(slots) if (s & 0xF000) != 0xD000]
    real.sort(key=lambda t: -t[1])  # descending by offset
    prev_end = PAGE  # absolute within page
    for slot_idx, off in real:
        start = off & 0x0FFF
        end = prev_end
        length = end - start
        row = data[base + start : base + end]
        print(f"   slot {slot_idx}  off=0x{start:04X}  len={length:5d}")
        head = row[: min(48, length)]
        tail = row[max(0, length - 24) :]
        print(f"     head: {head.hex(' ')}")
        print(f"     ascii: {_ascii(head)}")
        if length > 48:
            print(f"     tail: {tail.hex(' ')}")
            print(f"     tasc:  {_ascii(tail)}")
        prev_end = start


def main() -> None:
    if len(sys.argv) < 2:
        print("usage: decode_lval.py <accdb> [page ...]")
        sys.exit(2)
    data = Path(sys.argv[1]).read_bytes()
    if len(sys.argv) > 2:
        pages = [int(x) for x in sys.argv[2:]]
    else:
        pages = []
        for p in range(len(data) // PAGE):
            base = p * PAGE
            if data[base] == 0x01 and data[base + 4 : base + 8] == b"LVAL":
                pages.append(p)
    for p in pages:
        decode_page(data, p)


if __name__ == "__main__":
    main()
