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
from module_delete import stream_name_of  # noqa: E402
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402

STORAGE = "MSysAccessStorage"
VBA_STORAGE = 17
QUOTE = chr(34)
CRLF = chr(13) + chr(10)


def dir_row(db: AccessDatabase):
    for rid, row in db.table(STORAGE).rows_with_ids():
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes):
            return rid, decompress(row["Lv"])
    raise LookupError("the project has no dir stream")


def module_row(db: AccessDatabase, dir_stream: bytes, name: str):
    """The storage row holding a module's stream, found through the dir
    stream's MODULESTREAMNAME rather than by scanning the file: a module
    this library created carries no p-code to scan for."""
    row_name = stream_name_of(dir_stream, name)
    for rid, row in db.table(STORAGE).rows_with_ids():
        if row["ParentId"] == VBA_STORAGE and row["Name"] == row_name:
            return rid, bytes(row["Lv"])
    raise LookupError(f"no storage row named {row_name!r} for module {name!r}")


def rename_module_stream(db: AccessDatabase, path: Path, old: str, new: str) -> str:
    """Rewrite the module's `Attribute VB_Name` and the dir stream's
    MODULEOFFSET for it, which moves when the name's length changes."""
    rid_dir, decompressed = dir_row(db)
    target, payload = module_row(db, decompressed, old)
    at = find_moduleoffset_pos(decompressed, old)
    modoff = int.from_bytes(decompressed[at : at + 4], "little")

    quoted, replacement = QUOTE + old + QUOTE, QUOTE + new + QUOTE
    if modoff == 0:
        # A source-only stream, which is what create now writes: there is
        # no compiled region to rebuild, only the attribute to rewrite.
        lines = decompress(payload).decode("latin-1").split(CRLF)
        rebuilt = compress(CRLF.join(line.replace(quoted, replacement) for line in lines).encode("latin-1"))
        new_modoff = 0
    else:
        perf = Perf(payload, modoff)
        attributes = [line.replace(quoted, replacement) for line in perf.attribute_lines()]
        rebuilt, new_modoff = perf.build(
            new_source=CRLF.join(attributes + perf.source_lines()).encode("latin-1")
        )

    db.table(STORAGE).update_row(target, {"Lv": rebuilt})
    fixed = bytearray(decompressed)
    fixed[at : at + 4] = new_modoff.to_bytes(4, "little")
    db.table(STORAGE).update_row(rid_dir, {"Lv": compress(bytes(fixed))})
    return f"module stream (offset {modoff} -> {new_modoff})"


def set_source(db: AccessDatabase, name: str, code: str, kind: str = "module") -> str:
    """Replace a module's source outright.

    The stream becomes the compressed source alone, MODULEOFFSET goes to
    zero and the compiled cache is marked stale, so VBA compiles the new
    text on the next open.  That is what makes the forms the p-code
    compiler refuses -- `Const`, arrays, `Static`, fixed-length strings,
    a whole new procedure -- reachable anyway.
    """
    from module_create import invalidate_cache, module_source

    rid_dir, dir_stream = dir_row(db)
    target, _payload = module_row(db, dir_stream, name)
    storage = db.table(STORAGE)
    storage.update_row(target, {"Lv": module_source(name, code, kind)})

    fixed = bytearray(dir_stream)
    at = find_moduleoffset_pos(bytes(fixed), name)
    was = int.from_bytes(fixed[at : at + 4], "little")
    fixed[at : at + 4] = bytes(4)
    storage.update_row(rid_dir, {"Lv": compress(bytes(fixed))})

    for rid, row in list(storage.rows_with_ids()):
        if row["Name"] == "_VBA_PROJECT" and isinstance(row.get("Lv"), bytes):
            storage.update_row(rid, {"Lv": invalidate_cache(row["Lv"])})
    return f"source of {name!r} replaced (offset {was} -> 0, cache stale)"


if __name__ == "__main__":
    if len(sys.argv) < 5:
        raise SystemExit("usage: module_stream.py SOURCE TARGET NAME CODE [module|class]")
    _db = AccessDatabase(Path(sys.argv[1]))
    print(
        set_source(
            _db,
            sys.argv[3],
            sys.argv[4].replace("|", chr(10)),
            sys.argv[5] if len(sys.argv) > 5 else "module",
        )
    )
    _db.save(Path(sys.argv[2]))
