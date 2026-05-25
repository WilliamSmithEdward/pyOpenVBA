"""Dump the bytes BEFORE CAFE for selected corpus samples to look
for an identifier table (UTF-16LE / ASCII identifier names).
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"


def asciiify(b: bytes) -> str:
    return "".join(chr(c) if 32 <= c < 127 else "." for c in b)


def hex_lines(raw: bytes, max_bytes: int = 0x200) -> None:
    for i in range(0, min(len(raw), max_bytes), 16):
        chunk = raw[i:i + 16]
        print(f"  {i:04x}  {chunk.hex(' '):<47s}  {asciiify(chunk)}")


for name in ("040__sub_msgbox_hello.accdb", "044__sub_dim_int.accdb", "047__sub_dim_x_eq_7.accdb"):
    path = CORPUS / name
    if not path.exists():
        print(f"SKIP {name}")
        continue
    db = AccessFile(path)
    streams = db.find_module_streams()
    if not streams:
        continue
    s = streams[0]
    print(f"\n=== {name} (cafe_offset={s.cafe_offset:#x}, row_len={len(s.raw)}) ===")
    print("--- last 256 bytes BEFORE cafe ---")
    pre_start = max(0, s.cafe_offset - 256)
    raw_pre = s.raw[pre_start:s.cafe_offset]
    for i in range(0, len(raw_pre), 16):
        chunk = raw_pre[i:i + 16]
        print(f"  {pre_start + i:04x}  {chunk.hex(' '):<47s}  {asciiify(chunk)}")
    print("--- first 256 bytes of row ---")
    hex_lines(s.raw, 0x100)
