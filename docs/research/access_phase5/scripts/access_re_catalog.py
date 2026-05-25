"""
Phase 3 RE: catalog/symbol-table diff across corpus samples.

For each .accdb, classify every non-tombstone LVAL row as one of:

* ``project``  -- PROJECT plaintext (starts with ``ID="{``).
* ``ovba``     -- MS-OVBA module source stream (decompresses to
                  ``Attribute VB_Name = "..."``).
* ``catalog``  -- binary section between PROJECT plaintext and the
                  first OVBA stream; the symbol/name table we want
                  to reverse-engineer.
* ``other``    -- anything else (B9 literal tables, E3 comment
                  tables, ACE-internal opaque blobs, ...).

Then byte-diff the ``catalog`` rows across paired samples to isolate
which bytes encode which corpus axis (module name, module kind,
procedure name, source body).
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterator

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessFile, _ovba_decompress  # noqa: E402


def classify_rows(db: AccessFile) -> list[dict[str, object]]:
    out: list[dict[str, object]] = []
    for page, slot, row in db._iter_lval_rows():
        kind = "other"
        meta: dict[str, object] = {}
        if row.startswith(b'ID="{'):
            kind = "project"
        else:
            # try every OVBA candidate
            for label, blob in db._candidate_blobs(page, slot, row):
                try:
                    raw = _ovba_decompress(blob, stream_name=f"({page},{slot})")
                except Exception:
                    continue
                if raw.startswith(b"Attribute VB_Name = "):
                    kind = "ovba"
                    name = raw.decode("latin-1").split('"', 2)[1]
                    meta["module"] = name
                    meta["candidate"] = label
                    meta["decompressed_size"] = len(raw)
                    break
        out.append(
            {
                "page": page,
                "slot": slot,
                "len": len(row),
                "kind": kind,
                "head": row[:16].hex(),
                **meta,
            }
        )
    return out


def find_catalog_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """The 'catalog' rows are everything that is not PROJECT plaintext
    and not a successfully decoded OVBA module."""
    return [r for r in rows if r["kind"] == "other"]


def diff_bytes(a: bytes, b: bytes) -> list[tuple[int, int, int]]:
    """Return runs of (offset, length, kind) where kind = 0 same,
    1 differ. Only differ runs are emitted."""
    out: list[tuple[int, int, int]] = []
    n = min(len(a), len(b))
    i = 0
    while i < n:
        if a[i] != b[i]:
            j = i
            while j < n and a[j] != b[j]:
                j += 1
            out.append((i, j - i, 1))
            i = j
        else:
            i += 1
    if len(a) != len(b):
        out.append((n, abs(len(a) - len(b)), 1))
    return out


def hex_with_context(buf: bytes, off: int, length: int, ctx: int = 4) -> str:
    lo = max(0, off - ctx)
    hi = min(len(buf), off + length + ctx)
    pre = buf[lo:off].hex()
    mid = buf[off : off + length].hex()
    post = buf[off + length : hi].hex()
    return f"{pre}[{mid}]{post}"


def summarize(path: Path) -> dict[str, object]:
    db = AccessFile(path)
    rows = classify_rows(db)
    return {
        "path": str(path.name),
        "rows": rows,
        "lval_pages": sorted({r["page"] for r in rows}),
    }


def main() -> None:
    corpus = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"
    if not corpus.exists():
        sys.exit(f"corpus missing: {corpus}")

    # Phase 3 axis 1: module name (010..016).
    names = [
        "010__empty_StdModule_M.accdb",
        "011__empty_StdModule_AB.accdb",
        "012__empty_StdModule_ABC.accdb",
        "013__empty_StdModule_Mod1.accdb",
        "014__empty_StdModule_Module1.accdb",
        "015__empty_StdModule_LongName.accdb",
        "016__empty_StdModule_UnicodeA.accdb",
    ]

    # The "catalog" row of interest is (68, 1) -- it varies in length
    # proportionally with the module name length (see inventory below).
    # The (67, 6) row is the OVBA module source; (69, 1) is the PROJECT
    # plaintext; (68, 2), (69, 0) carry VBA references / type-info.
    #
    # Dump (68, 1) bytes side by side so name-table fields are visible.

    print("=" * 76)
    print("Catalog row (68, 1) across module-name axis (010..016)")
    print("=" * 76)
    rows: list[tuple[str, str, bytes]] = []
    for fn in names:
        path = corpus / fn
        if not path.exists():
            continue
        # The module name lives in the filename suffix; extract for the label.
        label = fn.split("StdModule_", 1)[1].removesuffix(".accdb")
        db = AccessFile(path)
        row = db._lval_row_bytes(68, 1)
        rows.append((fn, label, row))

    # Show length progression and full hex+ascii for each.
    print(f"\n{'label':<24} len  first-20-bytes")
    for fn, label, row in rows:
        head = row[:20].hex()
        print(f"  {label:<22} {len(row):>5}  {head}")

    # Find positions where the module name (latin-1) appears in each row.
    print("\nModule name positions inside (68, 1):")
    for fn, label, row in rows:
        name = label.encode("latin-1")
        positions = []
        i = 0
        while True:
            j = row.find(name, i)
            if j < 0:
                break
            positions.append(j)
            i = j + 1
        # also search for UTF-16-LE form
        u16 = name.decode("latin-1").encode("utf-16-le")
        u16_positions = []
        i = 0
        while True:
            j = row.find(u16, i)
            if j < 0:
                break
            u16_positions.append(j)
            i = j + 1
        print(
            f"  {label:<22}  ascii@{positions}  u16@{u16_positions}"
        )

    # Pairwise byte diff against baseline (010 / 'M').
    if rows:
        print("\nPairwise diff against 010 (module 'M'):")
        _, base_label, base = rows[0]
        for fn, label, row in rows[1:]:
            diffs = diff_bytes(base, row)
            tot = sum(d[1] for d in diffs)
            print(
                f"\n  {label}  ({len(row) - len(base):+d} bytes,"
                f" {len(diffs)} runs, {tot} differing bytes):"
            )
            for off, length, _ in diffs[:30]:
                a_chunk = base[off : off + length] if off < len(base) else b""
                b_chunk = row[off : off + length] if off < len(row) else b""
                print(
                    f"    +{off:>4} len={length:>3}  "
                    f"010={a_chunk.hex():<32} {label}={b_chunk.hex()}"
                )

    # Dump baseline catalog row in annotated chunks of 16.
    print("\nBaseline 010 catalog row (68, 1) full dump:")
    _, _, base = rows[0]
    for i in range(0, len(base), 16):
        chunk = base[i : i + 16]
        hexs = chunk.hex(" ")
        ascii_s = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        print(f"  +{i:>4}  {hexs:<48}  {ascii_s}")


if __name__ == "__main__":
    main()
