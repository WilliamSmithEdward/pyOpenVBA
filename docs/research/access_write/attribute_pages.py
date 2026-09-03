"""Say which table owns each page two databases differ on, and how its
rows changed.  The old research could only see raw pages; the engine can
now name the table, decode the row and diff it field by field.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access._alloc import read_usage_map_ref  # noqa: E402

PAGE = 4096


def owners(db: AccessDatabase) -> dict[int, str]:
    """Page number to the table that owns it."""
    out: dict[int, str] = {}
    for table in db.tables(include_system=True):
        d = table.definition
        out[d.page] = f"{table.name} (definition)"
        for extra in d.pages[1:]:
            out[extra] = f"{table.name} (definition, continued)"
        for page in read_usage_map_ref(db.store, d.owned_pages_ref).pages():
            if page < db.store.page_count:
                out.setdefault(page, f"{table.name} (data)")
        for number, (owned_ref, _free) in d.column_usage_maps.items():
            column = next((c.name for c in d.columns if c.number == number), str(number))
            for page in read_usage_map_ref(db.store, owned_ref).pages():
                if page < db.store.page_count:
                    out.setdefault(page, f"{table.name}.{column} (long values)")
        for i, real in enumerate(d.real_indexes):
            out.setdefault(real.root_page, f"{table.name} index {i} (root)")
            if i < len(d.index_usage_map_refs) if hasattr(d, "index_usage_map_refs") else False:
                continue
    return out


def rows_of(db: AccessDatabase, name: str) -> dict[tuple[int, int], dict[str, object]]:
    try:
        table = db.table(name)
    except Exception:  # noqa: BLE001
        return {}
    return {(rid.page, rid.slot): row for rid, row in table.rows_with_ids()}


def short(value: object, width: int = 60) -> str:
    text = repr(value)
    return text if len(text) <= width else text[: width - 3] + "..."


def compare(before: Path, after: Path, tables: list[str] | None = None) -> None:
    a, b = AccessDatabase(before), AccessDatabase(after)
    raw_a, raw_b = before.read_bytes(), after.read_bytes()
    owner_a, owner_b = owners(a), owners(b)
    pages = [
        n
        for n in range(1, max(len(raw_a), len(raw_b)) // PAGE)
        if raw_a[n * PAGE : (n + 1) * PAGE] != raw_b[n * PAGE : (n + 1) * PAGE]
    ]
    print(f"{len(raw_a) // PAGE} -> {len(raw_b) // PAGE} pages; {len(pages)} differ")
    for n in pages:
        who = owner_b.get(n) or owner_a.get(n) or "unowned"
        kind = raw_b[n * PAGE] if (n + 1) * PAGE <= len(raw_b) else -1
        count = sum(
            1
            for i in range(PAGE)
            if (n * PAGE + i) < min(len(raw_a), len(raw_b)) and raw_a[n * PAGE + i] != raw_b[n * PAGE + i]
        )
        print(f"  page {n:5}  type {kind:#04x}  {count:5} bytes  {who}")

    names = tables if tables is not None else sorted({o.split(" ")[0] for o in owner_b.values()})
    for name in names:
        old, new = rows_of(a, name), rows_of(b, name)
        gone = [k for k in old if k not in new]
        added = [k for k in new if k not in old]
        changed = [k for k in old if k in new and old[k] != new[k]]
        if not (gone or added or changed):
            continue
        print(f"\n### {name}")
        for key in gone:
            print(f"  - row {key}: " + ", ".join(f"{c}={short(v, 40)}" for c, v in old[key].items() if v not in (None, b"")))
        for key in added:
            print(f"  + row {key}: " + ", ".join(f"{c}={short(v, 40)}" for c, v in new[key].items() if v not in (None, b"")))
        for key in changed:
            for column in old[key]:
                if old[key][column] != new[key][column]:
                    print(f"  ~ row {key} {column}: {short(old[key][column])} -> {short(new[key][column])}")


if __name__ == "__main__":
    compare(Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3:] or None)
