"""Byte-level diff of orig vs rename-only dir-stream and OVBA cache row."""
from __future__ import annotations
from pathlib import Path
from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]


def main() -> None:
    orig_p = REPO / "tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb"
    ren_p = REPO / "demo/output/access_phase5f/diag_E_rename_only.accdb"
    db_o = AccessFile(orig_p)
    db_r = AccessFile(ren_p)
    a = db_o._find_catalog_row()  # type: ignore[attr-defined]
    b = db_r._find_catalog_row()  # type: ignore[attr-defined]
    assert a and b
    ao = a[2]
    br = b[2]
    print(f"orig dir-stream len: {len(ao)}   ren dir-stream len: {len(br)}")
    # Find first diff
    n = min(len(ao), len(br))
    first = None
    for i in range(n):
        if ao[i] != br[i]:
            first = i
            break
    print(f"first diff byte: {first}")
    if first is not None:
        ctx = 80
        s = max(0, first - 16)
        e = min(n, first + ctx)
        print()
        print(f"context [{s}..{e}]:")
        print(f"  orig: {ao[s:e].hex()}")
        print(f"  ren : {br[s:e].hex()}")
        print(f"  orig ascii: {ao[s:e].decode('latin-1', errors='replace')!r}")
        print(f"  ren  ascii: {br[s:e].decode('latin-1', errors='replace')!r}")
    # Check tails
    if len(ao) != len(br):
        print()
        print(f"length delta: {len(br) - len(ao)}")
    # Also dump the M..end region from both, hex
    print()
    print("--- search for 'M\\0' (utf-16) and 'Renamed_M' ---")
    print("orig: occurrences of b'M\\x00' :", [i for i in range(len(ao)) if ao[i:i+2]==b'M\x00'])
    print("ren : occurrences of b'Renamed_M\\x00\\x00...':",
          ao.find(b'Renamed_M'), br.find(b'Renamed_M'))
    # Search for the opaque stream name
    sn = b'KOZOJLJLCGFVSDDQCZFLEWAYIKHL'
    print(f"stream-name {sn!r} in orig at: {ao.find(sn)}, in ren at: {br.find(sn)}")


if __name__ == "__main__":
    main()
