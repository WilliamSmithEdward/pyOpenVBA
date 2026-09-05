"""Make the engine skeleton Compact and Repair starts from.

DAO's ``CompactDatabase`` copies the source into a bare database -- what
``DBEngine.CreateDatabase`` makes -- before that routine writes its
permission rows.  Measured: the pages after the system tables of a
compacted file equal the bare database's, the bare file's permission page
is the first page the copy takes, and the copied permission table starts
empty.

Usage, with the DAO oracle's ``create-blank`` (ACE) or ``create-blank-mdb``
(Jet 4) command having produced the bare file::

    powershell -File tests/live_access_test/dao_oracle.ps1 -Command create-blank -Path bare.accdb
    python docs/research/access_compaction/make_skeleton.py bare.accdb

writes ``src/pyopenvba/_templates/blank_files/engine_skeleton.accdb`` (or
``.mdb`` for a Jet 4 bare file): the bare file with its permission rows
taken back out as if never written -- their data page dropped, the index
root fresh, the counters and the map bits reset.
"""

from __future__ import annotations

import struct
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._alloc import GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW, read_usage_map, remove_from_map, set_usage_bit  # noqa: E402
from pyopenvba.access._pages import read_usage_map_ref  # noqa: E402
from pyopenvba.access._schema import empty_index_root  # noqa: E402
from pyopenvba.access._tdef import OFFSET_INDEX_HEADERS, OFFSET_ROW_COUNT  # noqa: E402

TEMPLATES = Path(__file__).resolve().parents[3] / "src" / "pyopenvba" / "_templates" / "blank_files"


def main(bare_path: Path) -> None:
    bare = AccessDatabase(bare_path)
    store = bare.store
    aces = bare.table("MSysACEs")
    definition = aces.definition
    owned = sorted(read_usage_map_ref(store, definition.owned_pages_ref).pages())
    last = store.page_count - 1
    if owned != [last] or not definition.row_count:
        raise SystemExit(f"not the bare database expected: {store.page_count} pages, ACE pages {owned}, {definition.row_count} rows")
    remove_from_map(store, definition.owned_pages_ref, last)
    remove_from_map(store, definition.free_space_pages_ref, last)
    definition.row_count = 0
    bare.patch_definition(definition, OFFSET_ROW_COUNT, struct.pack("<I", 0))
    real = definition.real_indexes[0]
    real.entry_count = real.row_count = 0
    bare.patch_definition(definition, OFFSET_INDEX_HEADERS, struct.pack("<II", 0, 0))
    store.write(real.root_page, empty_index_root(definition.page))
    set_usage_bit(store, read_usage_map(store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW), last, True)
    store.truncate(last)
    out = TEMPLATES / ("engine_skeleton" + bare_path.suffix.lower())
    out.write_bytes(store.to_bytes())
    check = AccessDatabase(out)
    print(f"wrote {out}: {check.store.page_count} pages, {check.table('MSysACEs').row_count} permission rows, {len(check.catalog())} catalog rows")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
