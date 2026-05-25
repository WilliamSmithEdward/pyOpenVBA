"""Compare all rU@ rows side-by-side to find a deterministic discriminator."""
from __future__ import annotations
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"

samples = [
    "010__empty_StdModule_M.accdb",
    "020__empty_ClassModule_C.accdb",
    "030__sub_A_empty.accdb",
    "040__sub_msgbox_hello.accdb",
    "044__sub_dim_int.accdb",
    "048__sub_let_int_42.accdb",
    "051__sub_for_1_to_3.accdb",
]

for name in samples:
    path = CORPUS / name
    if not path.exists():
        continue
    db = AccessFile(path)
    print(f"\n=== {name} ===")
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        if not row.startswith(b"\x72\x55\x40"):
            continue
        hdr = bytes(row[:32])
        # decode some likely fields
        # 'rU@' at 0, then often 0x40 markers, then a counter/u16 maybe.
        print(f"  ({page:>3},{slot}) len={len(row):>4}  hex[0:32]={hdr.hex(' ')}")
        # also print bytes 32..64 to spot per-row distinction
        h2 = bytes(row[32:64])
        print(f"                          hex[32:64]={h2.hex(' ')}")
