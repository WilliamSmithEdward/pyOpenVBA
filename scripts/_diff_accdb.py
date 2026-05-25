"""Byte-level page-aware diff between two .accdb files.

Outputs a human-readable summary of which 4KiB pages changed and the
ranges of changed bytes within each page. Useful for inferring where
Access writes its updates when a single VBA mutation is performed.
"""

from __future__ import annotations

import sys
from pathlib import Path

PAGE = 4096


def diff(a: Path, b: Path, *, max_pages: int = 50, max_runs_per_page: int = 8) -> None:
    da = a.read_bytes()
    db = b.read_bytes()
    print(f"A: {a.name}  {len(da)} bytes  {len(da)//PAGE} pages")
    print(f"B: {b.name}  {len(db)} bytes  {len(db)//PAGE} pages")
    if len(da) != len(db):
        print(f"!! length delta: {len(db) - len(da):+d}")

    pages_changed: list[tuple[int, list[tuple[int, int]]]] = []
    common = min(len(da), len(db))
    np = common // PAGE
    for p in range(np):
        base = p * PAGE
        sa = da[base : base + PAGE]
        sb = db[base : base + PAGE]
        if sa == sb:
            continue
        runs: list[tuple[int, int]] = []
        i = 0
        while i < PAGE:
            if sa[i] != sb[i]:
                j = i
                while j < PAGE and sa[j] != sb[j]:
                    j += 1
                runs.append((i, j - i))
                i = j
            else:
                i += 1
        pages_changed.append((p, runs))

    extra = len(db) - len(da)
    if extra > 0:
        extra_pages = extra // PAGE
        print(f"!! {extra_pages} new page(s) appended after page {np - 1}")

    print(f"== {len(pages_changed)} page(s) changed ==")
    for p, runs in pages_changed[:max_pages]:
        page_type_a = da[p * PAGE]
        tag_a = da[p * PAGE + 4 : p * PAGE + 8]
        marker = ""
        if tag_a == b"LVAL":
            marker = " [LVAL]"
        elif page_type_a == 0x02:
            marker = " [TABLE_DEF]"
        elif page_type_a == 0x01:
            marker = " [DATA]"
        elif page_type_a == 0x00:
            marker = " [DB_DEF]"
        head = f"page {p:5d} @0x{p*PAGE:06X} type=0x{page_type_a:02X}{marker} -> {len(runs)} run(s):"
        print(head)
        for off, ln in runs[:max_runs_per_page]:
            abs_off = p * PAGE + off
            a_bytes = da[abs_off : abs_off + min(ln, 32)]
            b_bytes = db[abs_off : abs_off + min(ln, 32)]
            print(f"    +{off:4d} (0x{abs_off:06X}) len={ln:5d}")
            print(f"        A: {a_bytes.hex(' ')}{'...' if ln > 32 else ''}")
            print(f"        B: {b_bytes.hex(' ')}{'...' if ln > 32 else ''}")
            try:
                ascii_a = a_bytes.decode("ascii", errors="replace")
                ascii_b = b_bytes.decode("ascii", errors="replace")
                if any(c.isprintable() and not c.isspace() for c in ascii_a + ascii_b):
                    print(f"        A: {ascii_a!r}")
                    print(f"        B: {ascii_b!r}")
            except Exception:
                pass
        if len(runs) > max_runs_per_page:
            print(f"    ... and {len(runs) - max_runs_per_page} more runs")
    if len(pages_changed) > max_pages:
        print(f"... and {len(pages_changed) - max_pages} more pages")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: diff_accdb.py <a.accdb> <b.accdb>")
    diff(Path(sys.argv[1]), Path(sys.argv[2]))
