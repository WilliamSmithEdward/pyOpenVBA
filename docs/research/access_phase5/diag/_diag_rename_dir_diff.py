"""Diff dir-stream records in original sample 040 vs diag_E (rename only)."""
from __future__ import annotations
import struct
from pathlib import Path
from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]


def dump_dir(path: Path):
    db = AccessFile(path)
    found = db._find_catalog_row()  # type: ignore[attr-defined]
    assert found is not None
    _, _, dir_raw = found
    out = []
    i = 0
    while i + 6 <= len(dir_raw):
        rid = struct.unpack_from("<H", dir_raw, i)[0]
        size = struct.unpack_from("<I", dir_raw, i + 2)[0]
        if i + 6 + size > len(dir_raw):
            out.append((i, rid, size, b"<truncated>"))
            break
        data = bytes(dir_raw[i + 6 : i + 6 + size])
        out.append((i, rid, size, data))
        i += 6 + size
    return out


def main() -> None:
    orig_p = REPO / "tests/live_access_test/re_corpus/samples/040__sub_msgbox_hello.accdb"
    ren_p = REPO / "demo/output/access_phase5f/diag_E_rename_only.accdb"
    a = dump_dir(orig_p)
    b = dump_dir(ren_p)
    print(f"orig records: {len(a)}   renamed records: {len(b)}")
    print()
    n = max(len(a), len(b))
    for i in range(n):
        ra = a[i] if i < len(a) else (-1, -1, -1, b"")
        rb = b[i] if i < len(b) else (-1, -1, -1, b"")
        same = ra[1] == rb[1] and ra[2] == rb[2] and ra[3] == rb[3]
        mark = "   " if same else "** "
        adisp = ra[3][:40].hex() + ("..." if ra[2] > 40 else "")
        bdisp = rb[3][:40].hex() + ("..." if rb[2] > 40 else "")
        print(
            f"{mark}{i:3d} ORIG id=0x{ra[1]:04X} sz={ra[2]:5d} "
            f"REN id=0x{rb[1]:04X} sz={rb[2]:5d}"
        )
        if not same:
            print(f"      ORIG: {adisp}")
            print(f"      REN : {bdisp}")


if __name__ == "__main__":
    main()
