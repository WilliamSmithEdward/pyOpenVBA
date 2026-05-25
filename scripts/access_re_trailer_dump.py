"""Dump the bytes around the ProcTrailer for samples 040, 041, 042."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"

for name in ["040__sub_msgbox_hello.accdb", "041__sub_msgbox_world.accdb",
             "042__sub_msgbox_long.accdb", "043__sub_msgbox_two.accdb",
             "030__sub_A_empty.accdb", "044__sub_dim_int.accdb"]:
    raw = AccessFile(CORPUS / name).read_module_pcode_stream().raw
    idx = raw.rfind(b"\x7b\x02")
    print(f"\n{name}  rfind(7B 02) = 0x{idx:x}")
    # dump 32 bytes from the trailer
    chunk = raw[idx:idx + 32]
    print("  bytes:", " ".join(f"{b:02x}" for b in chunk))
