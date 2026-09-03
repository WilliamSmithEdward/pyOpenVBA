"""The module's own stream: its `Attribute VB_Name` line, and the
MODULEOFFSET in the dir stream that says where the source starts.

The stream is found with the library's reader, rebuilt with the research
module's `Perf`, and written back through the engine's row writer -- which
is the part that did not exist before, since renaming to a name of a
different length resizes the row.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from accdb_write import Perf, find_moduleoffset_pos  # noqa: E402
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access_read import AccessReader  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402

STORAGE = "MSysAccessStorage"


def dir_row(db: AccessDatabase):
    for rid, row in db.table(STORAGE).rows_with_ids():
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes):
            return rid, decompress(row["Lv"])
    raise LookupError("the project has no dir stream")


def rename_module_stream(db: AccessDatabase, path: Path, old: str, new: str) -> str:
    """Rewrite the module's `Attribute VB_Name` and the dir stream's
    MODULEOFFSET for it, which moves when the name's length changes."""
    reader = AccessReader(path)
    streams = reader.find_module_streams()
    stream = next((s for s in streams if s.name.lower() == old.lower()), None)
    if stream is None:
        raise LookupError(f"no module stream for {old!r}; have {[s.name for s in streams]}")

    rid_dir, decompressed = dir_row(db)
    at = find_moduleoffset_pos(decompressed, old)
    modoff = int.from_bytes(decompressed[at : at + 4], "little")

    perf = Perf(bytes(stream.raw), modoff)
    attributes = [line.replace(f'"{old}"', f'"{new}"') for line in perf.attribute_lines()]
    body = "\r\n".join(attributes + perf.source_lines()).encode("latin-1")
    rebuilt, new_modoff = perf.build(new_source=body)

    # The reader's page and slot name the long value's own row, not the
    # catalog row that points at it, so the row is found by its payload.
    storage = db.table(STORAGE)
    payload = bytes(stream.raw)
    target = next(
        (rid for rid, row in storage.rows_with_ids() if row.get("Lv") == payload), None
    )
    if target is None:
        raise LookupError(f"no storage row holds the module stream for {old!r}")
    storage.update_row(target, {"Lv": rebuilt})

    fixed = bytearray(decompressed)
    fixed[at : at + 4] = new_modoff.to_bytes(4, "little")
    storage.update_row(rid_dir, {"Lv": compress(bytes(fixed))})
    return f"module stream (offset {modoff} -> {new_modoff})"
