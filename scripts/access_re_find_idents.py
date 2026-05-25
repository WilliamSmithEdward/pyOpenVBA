"""Locate the LVAL row carrying VBA identifier strings ("MsgBox",
"Integer", "x", etc.). Search every row for known identifier names
in both UTF-16-LE and ASCII forms.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"

# Known identifiers that MUST appear by name in *some* row if Access
# stores identifier names like Office does.
NEEDLES_TEXT = [
    b"M\x00s\x00g\x00B\x00o\x00x\x00",  # UTF-16-LE
    b"MsgBox",                            # ASCII
    b"I\x00n\x00t\x00e\x00g\x00e\x00r\x00",
    b"Integer",
    b"v\x00b\x00O\x00K\x00",
    b"vbOKOnly",
]

for sample in ("040__sub_msgbox_hello.accdb", "044__sub_dim_int.accdb"):
    path = CORPUS / sample
    if not path.exists():
        continue
    db = AccessFile(path)
    print(f"\n=== {sample} ===")
    rows = list(db._iter_lval_rows())
    for needle in NEEDLES_TEXT:
        hits = []
        for page, slot, row in rows:
            raw = bytes(row)
            if needle in raw:
                hits.append((page, slot, raw.find(needle), len(raw)))
        flag = "UTF16" if b"\x00" in needle else "ASCII"
        print(f"  {flag:5s} {needle!r:40s} -> {hits if hits else 'NOT FOUND'}")
