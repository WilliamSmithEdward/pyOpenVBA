"""Diff `_VBA_PROJECT` rows across samples to localize the identifier
table within the stream.

Strategy: take pairs of samples that differ ONLY in a user identifier
name and diff the cc 61 row. The diff region IS the identifier table.

Pairs:
  010 ("M")            vs 011 ("AB")            -- module name diff
  010 ("M")            vs 015 ("LongName")      -- module name diff
  030 (Sub A())        vs 048 (Sub A() x = 42)  -- + variable "x"
  040 (MsgBox "hello") vs 044 (Dim x As Integer)-- + variable "x"
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"


def get_cc61_row(p: Path) -> bytes:
    db = AccessFile(p)
    for _page, _slot, row in db._iter_lval_rows():  # type: ignore[attr-defined]
        b = bytes(row)
        if b.startswith(b"\xcc\x61"):
            return b
    raise SystemExit(f"no cc61 row in {p.name}")


def show_diff(a_path: str, b_path: str) -> None:
    a = get_cc61_row(CORPUS / a_path)
    b = get_cc61_row(CORPUS / b_path)
    n = min(len(a), len(b))
    print(f"\n=== {a_path} ({len(a)} B) vs {b_path} ({len(b)} B) ===")
    # Find first/last differing byte within the common prefix.
    first = next((i for i in range(n) if a[i] != b[i]), -1)
    last = next((i for i in range(n - 1, -1, -1) if a[i] != b[i]), -1)
    print(f"  first_diff={first}  last_diff={last}  common={n}  da-db={len(a)-len(b)}")
    if first < 0:
        return
    # Show diff windows.
    for label, buf in (("A", a), ("B", b)):
        s = max(0, first - 16)
        e = min(len(buf), last + 16)
        print(f"  {label}[{s:#06x}:{e:#06x}] = {buf[s:e].hex(' ')}")

    # Find all contiguous diff regions (>=4 bytes apart).
    regions: list[tuple[int, int]] = []
    cur: tuple[int, int] | None = None
    for i in range(n):
        if a[i] != b[i]:
            if cur is None:
                cur = (i, i)
            else:
                cur = (cur[0], i)
        else:
            if cur is not None and i - cur[1] > 4:
                regions.append(cur)
                cur = None
    if cur is not None:
        regions.append(cur)
    print(f"  diff regions: {regions}")


show_diff("010__empty_StdModule_M.accdb", "011__empty_StdModule_AB.accdb")
show_diff("010__empty_StdModule_M.accdb", "015__empty_StdModule_LongName.accdb")
show_diff("030__sub_A_empty.accdb", "048__sub_let_int_42.accdb")
show_diff("040__sub_msgbox_hello.accdb", "044__sub_dim_int.accdb")
show_diff("030__sub_A_empty.accdb", "031__sub_B_empty.accdb")
show_diff("044__sub_dim_int.accdb", "045__sub_dim_long.accdb")
