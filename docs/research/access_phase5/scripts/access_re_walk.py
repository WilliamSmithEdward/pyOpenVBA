"""Walk a (page, slot) LVAL chain to validate the row-prefix hypothesis."""
from __future__ import annotations
import sys
from pathlib import Path

PAGE = 4096
HDR = 28


def read_slot(data: bytes, page: int, slot: int) -> tuple[int, int, bytes]:
    base = page * PAGE
    if data[base] != 0x01 or data[base + 4 : base + 8] != b"LVAL":
        raise ValueError(f"page {page} is not LVAL")
    n = int.from_bytes(data[base + 12 : base + 14], "little")
    if slot >= n:
        raise ValueError(f"slot {slot} out of range (n={n})")
    off = int.from_bytes(data[base + 14 + 2 * slot : base + 16 + 2 * slot], "little")
    if (off & 0xF000) == 0xD000:
        raise ValueError(f"slot {slot} on page {page} is a tombstone")
    start = off & 0x0FFF
    # Find row END: smallest offset > start among non-tombstone slots, else PAGE.
    end = PAGE
    for i in range(n):
        o = int.from_bytes(data[base + 14 + 2 * i : base + 16 + 2 * i], "little")
        if (o & 0xF000) == 0xD000:
            continue
        o2 = o & 0x0FFF
        if o2 > start and o2 < end:
            end = o2
    return start, end, bytes(data[base + start : base + end])


def walk(data: bytes, start_page: int, start_slot: int, max_steps: int = 32) -> None:
    cur_p, cur_s = start_page, start_slot
    step = 0
    total = 0
    while step < max_steps:
        try:
            off, end, row = read_slot(data, cur_p, cur_s)
        except ValueError as e:
            print(f"  step {step}: ERROR {e}")
            return
        nxt_s = row[0]
        nxt_p = int.from_bytes(row[1:4], "little")
        print(f"  step {step}: page={cur_p:4d} slot={cur_s} off=0x{off:04X} len={end-off:5d}  "
              f"next=({nxt_p}, {nxt_s})  head={row[:16].hex(' ')}")
        total += len(row) - 4
        if nxt_p == 0 and nxt_s == 0:
            print(f"  END after {step+1} chunks, total payload (excluding 4-byte prefixes) = {total} bytes")
            return
        if nxt_p >= len(data) // PAGE:
            print(f"  step {step}: next_page {nxt_p} out of range")
            return
        cur_p, cur_s = nxt_p, nxt_s
        step += 1
    print("  hit max_steps")


def main() -> None:
    path = Path(sys.argv[1])
    start_page = int(sys.argv[2])
    start_slot = int(sys.argv[3])
    data = path.read_bytes()
    print(f"Walking chain from ({start_page}, {start_slot}) in {path.name}:")
    walk(data, start_page, start_slot)


if __name__ == "__main__":
    main()
