"""Delete a module: a command line over the shipped writer.

The structures a module occupies were worked out here and now live in
`pyopenvba.access._vba` and `AccessDatabase.delete_module`.  The piece
that stayed behind is `remove_module` in `vba_module_table.py`: it takes
a module's entry out of the compiled `_VBA_PROJECT` cache byte-exactly,
which the shipped writer does not need because it marks the cache stale
instead.

    python module_delete.py SOURCE TARGET NAME
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._vba import module_blocks, records  # noqa: E402
from pyopenvba.vba import decompress  # noqa: E402

PROJECTCOOKIE = 0x0013


def stream_name_of(stream: bytes, name: str) -> str:
    """The storage row a module's code lives in, from its dir block."""
    for module, row_name, _kind in module_blocks(stream):
        if module.lower() == name.lower():
            return row_name
    raise LookupError(f"the dir stream names no stream for {name!r}")


def project_cookie(stream: bytes) -> bytes:
    for _at, ident, _size, payload in records(stream):
        if ident == PROJECTCOOKIE:
            return payload
    raise LookupError("the dir stream has no project cookie")


def dir_stream(db: AccessDatabase) -> bytes:
    for _rid, row in db.table("MSysAccessStorage").rows_with_ids():
        value = row.get("Lv")
        if row["Name"] == "dir" and isinstance(value, bytes):
            return decompress(value)
    raise LookupError("the database has no dir stream")


def delete(source: Path, target: Path, name: str) -> str:
    db = AccessDatabase(source)
    db.delete_module(name)
    db.save(target)
    return f"deleted {name!r}; {[m.name for m in db.modules()]} left"


if __name__ == "__main__":
    print(delete(Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3]))
