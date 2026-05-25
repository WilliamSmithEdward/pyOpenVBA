"""List ALL rows starting with the 'rU@' p-code magic across samples."""
from __future__ import annotations
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
from pyopenvba.access import AccessFile  # noqa: E402

CORPUS = ROOT / "tests" / "live_access_test" / "re_corpus" / "samples"

for path in sorted(CORPUS.glob("*.accdb")):
    name = path.name
    if not (name.startswith("030") or name.startswith("04") or name.startswith("05") or name.startswith("020") or name.startswith("010")):
        continue
    db = AccessFile(path)
    hits: list[tuple[int, int, int]] = []
    for page, slot, row in db._iter_lval_rows():  # pyright: ignore[reportPrivateUsage]
        if row.startswith(b"\x72\x55\x40"):
            hits.append((page, slot, len(row)))
    print(f"{name}: {hits}")
