"""Search every LVAL row for the standard VBA module-stream CAFE magic.

pcodedmp expects: ``[binary metadata] ... 0xCAFE <numLines> <line records>
<pcode bytes> ... <compressed source>``. If Access stores standard
VBA p-code (not just execodes) we should find the 0xCAFE word.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"

for path in sorted(CORPUS.glob("04*.accdb")):
    db = AccessFile(path)
    print(f"\n=== {path.name} ===")
    hits: list[tuple[int, int, int]] = []
    for page, slot, row in db._iter_lval_rows():
        buf = bytes(row)
        for needle in (b"\xfe\xca", b"\xca\xfe"):
            start = 0
            while True:
                i = buf.find(needle, start)
                if i < 0:
                    break
                hits.append((page, slot, i))
                start = i + 1
    print(f"  CAFE/FECA hits: {len(hits)}")
    for page, slot, off in hits[:10]:
        print(f"    ({page},{slot}) offset={off:#x}")
