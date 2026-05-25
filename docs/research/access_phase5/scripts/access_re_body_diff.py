"""Phase 4 — focus on body-varying samples to isolate authoritative p-code.

Plan:
1. Baseline = sample 030__sub_a_empty (an empty Sub A()).
2. Targets = 040..049 (single-statement bodies) and 050..051 (If/For).
3. For each LVAL row position present in BOTH baseline and target, report
   length delta + first-byte/last-byte diff.
4. Rows where the length changes proportionally to body size are
   p-code candidates; rows whose content changes but length stays
   constant are likely fixed-size symbol records.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile  # noqa: E402
from pyopenvba.vba import decompress as _ovba_decompress  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"

_DIR_MAGIC = b"\x01\x00\x04\x00\x00\x00"


def is_ovba_cache(row: bytes) -> bool:
    """True if row contains an embedded OVBA stream at any offset that
    decompresses to a module source cache."""
    if not row:
        return False
    i = 0
    n = len(row)
    while i + 3 <= n:
        j = row.find(b"\x01", i)
        if j < 0 or j + 3 > n:
            break
        hdr = int.from_bytes(row[j + 1 : j + 3], "little")
        if ((hdr >> 12) & 0x7) == 0b011:
            try:
                raw = _ovba_decompress(bytes(row[j:]), stream_name="probe")
            except Exception:
                pass
            else:
                if raw.startswith(b"Attribute VB_Name = "):
                    return True
        i = j + 1
    return False


def classify(row: bytes) -> str:
    if not row:
        return "empty"
    if row[0] == 0x01 and len(row) >= 3:
        hdr = int.from_bytes(row[1:3], "little")
        if ((hdr >> 12) & 0x7) == 0b011:
            try:
                raw = _ovba_decompress(bytes(row), stream_name="probe")
            except Exception:
                pass
            else:
                if raw.startswith(_DIR_MAGIC):
                    return "catalog_dir"
                if raw.startswith(b"Attribute VB_Name = "):
                    return "ovba_cache_root"
    if is_ovba_cache(row):
        return "ovba_cache_wrapped"
    if b'ID="{' in row[:64]:
        return "project_plaintext"
    return "unknown_binary"


def collect(path: Path) -> dict[tuple[int, int], tuple[str, bytes]]:
    db = AccessFile(path)
    out: dict[tuple[int, int], tuple[str, bytes]] = {}
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        # Skip page 42 system rows entirely (constant across all VBA samples).
        if page == 42:
            continue
        out[(page, slot)] = (classify(row), bytes(row))
    return out


def diff_pair(baseline_name: str, target_name: str) -> None:
    base = collect(CORPUS / baseline_name)
    targ = collect(CORPUS / target_name)
    print(f"\n=== {baseline_name}  vs  {target_name} ===")
    keys = sorted(set(base) | set(targ))
    for k in keys:
        b = base.get(k)
        t = targ.get(k)
        if b is None:
            print(f"  +{k}  ADDED in target: {t[0]} len={len(t[1])}")  # type: ignore[index]
            continue
        if t is None:
            print(f"  -{k}  REMOVED in baseline: {b[0]} len={len(b[1])}")
            continue
        b_lbl, b_data = b
        t_lbl, t_data = t
        if b_data == t_data:
            continue
        dlen = len(t_data) - len(b_data)
        # Find first differing byte
        first_diff = next(
            (i for i in range(min(len(b_data), len(t_data))) if b_data[i] != t_data[i]),
            None,
        )
        # Last common prefix
        suffix_match = 0
        for i in range(1, min(len(b_data), len(t_data)) + 1):
            if b_data[-i] != t_data[-i]:
                break
            suffix_match = i
        print(
            f"  *{k}  {b_lbl} len {len(b_data)} -> {len(t_data)} ({dlen:+d})"
            f"  first_diff@{first_diff}  suffix_match={suffix_match}"
        )


def main() -> None:
    baseline = "030__sub_A_empty.accdb"
    if not (CORPUS / baseline).exists():
        print(f"no baseline {baseline}", file=sys.stderr)
        return
    targets = [
        "031__sub_B_empty.accdb",
        "032__sub_AB_empty.accdb",
        "033__sub_LongName_empty.accdb",
        "040__sub_msgbox_hello.accdb",
        "041__sub_msgbox_world.accdb",
        "042__sub_msgbox_long.accdb",
        "043__sub_msgbox_two.accdb",
        "044__sub_dim_int.accdb",
        "045__sub_dim_long.accdb",
        "046__sub_dim_string.accdb",
        "047__sub_let_int.accdb",
        "048__sub_let_int_42.accdb",
        "049__sub_comment_only.accdb",
        "050__sub_if_true.accdb",
        "051__sub_for_1_to_3.accdb",
    ]
    for t in targets:
        if (CORPUS / t).exists():
            diff_pair(baseline, t)


if __name__ == "__main__":
    main()
