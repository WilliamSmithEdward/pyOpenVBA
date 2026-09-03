"""Live ACE gate for the storage engine (opt-in).

The ACE engine itself, driven through DAO from PowerShell with no Access
window and no VBA, builds a table holding every column type and fills it.
The engine reads the file back with no Office involved.  The two views
must agree field for field, so what is compared is exactly what the real
engine would hand to any client.

Opt-in: set ``RUN_LIVE_ACCESS=1`` on a Windows machine with the Access
database engine installed (it ships with Office).  Skipped everywhere
else, including CI.  DAO is a test-time oracle only; pyOpenVBA never uses
COM.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import subprocess
import sys
import uuid
from decimal import Decimal
from collections.abc import Callable
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, CatalogEntry, RowId, Table
from pyopenvba.access._index import leaf_pages
from pyopenvba.access_read import AccessError
from test_access_write import check_indexes

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS") != "1" or sys.platform != "win32",
    reason="live ACE gate: set RUN_LIVE_ACCESS=1 on Windows with Access installed",
)

HERE = Path(__file__).parent
ORACLE = HERE / "live_access_test" / "dao_oracle.ps1"
TEMPLATE = (
    HERE.parents[0]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
ROWS = 120
_TIMEOUT = 300


def oracle(*args: str) -> str:
    """Run the DAO oracle script; its stdout is the result."""
    done = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ORACLE),
            *args,
        ],
        capture_output=True,
        timeout=_TIMEOUT,
        check=False,
    )
    stdout = done.stdout.decode("utf-8", errors="replace")
    if done.returncode != 0:
        raise AssertionError(
            f"dao_oracle.ps1 {' '.join(args)} failed ({done.returncode}):\n"
            f"{stdout}\n{done.stderr.decode('utf-8', errors='replace')}"
        )
    return stdout


def _format(value: object) -> str:
    """Mirror Format-Cell in dao_oracle.ps1 for the engine's decoded values."""
    if value is None:
        return "<null>"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, Decimal):
        return f"{value:.4f}"
    if isinstance(value, float):
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text if text not in ("", "-0") else "0"
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, bytes):
        return value.hex()
    if isinstance(value, uuid.UUID):
        return "{" + str(value).upper() + "}"
    if isinstance(value, str):
        return value.replace("\t", "\\t").replace("\n", "\\n")
    return str(value)


def _engine_dump(db: AccessDatabase, name: str) -> list[str]:
    table = db.table(name)
    rows = sorted(table.rows(), key=lambda r: int(str(r["Id"])))
    return [
        "\t".join(f"{col}={_format(row[col])}" for col in table.column_names)
        for row in rows
    ]


def _diff(expected: str, mine: list[str]) -> list[str]:
    expected_lines = expected.split("\n")
    out = [
        f"row {i + 1}:\n  engine(ACE): {a}\n  pyopenvba:   {m}"
        for i, (a, m) in enumerate(zip(expected_lines, mine))
        if a != m
    ]
    if len(expected_lines) != len(mine):
        out.append(f"ACE returned {len(expected_lines)} rows, pyopenvba {len(mine)}")
    return out


def _check_indexes(db: AccessDatabase, oracle_built: bool = True) -> None:
    """Every index of AllTypes must decode to the row values it points at,
    in key order, with the node pages naming their children's last keys.
    ``oracle_built`` also expects the oracle's index set -- one per
    indexable column type, a descending one, a two-column one, a unique
    ignore-nulls one."""
    from pyopenvba.access._index import TextKey, node_pages, parse_index_page
    from pyopenvba.access._rows import split_row

    table = db.table("AllTypes")
    rows: dict[tuple[int, int], dict[str, object]] = {}
    for page, slot, data in table.raw_rows():
        rows[(page, slot)] = table.decode(split_row(table.definition, data))
    names = {i.name for i in table.indexes}
    if oracle_built:
        assert {"IX_Txt", "IX_BigDesc", "IX_FlagTiny", "IX_UniqueBig", "IX_Frac", "IX_Uid", "IX_Bin"} <= names
    for index in table.indexes:
        expected = len(rows)
        if index.ignores_nulls:
            expected -= sum(
                1 for r in rows.values() if all(r[c.name] is None for c, _ in index.columns)
            )
        count = 0
        previous: list[object] | None = None
        for values, page, row in index.entries():
            count += 1
            actual = rows[(page, row)]
            for (column, _asc), value in zip(index.columns, values):
                if value is None:
                    assert actual[column.name] is None, (index.name, column.name)
                elif isinstance(value, TextKey):
                    assert isinstance(actual[column.name], str)
                elif isinstance(value, float):
                    assert isinstance(actual[column.name], float)
                    assert abs(value - float(str(actual[column.name]))) < 1e-9
                else:
                    assert value == actual[column.name], (index.name, column.name, value)
            previous = values
        assert previous is not None
        assert count == expected, (index.name, count, expected)
        for node in node_pages(db.store, index.real.root_page):
            for entry in node.entries:
                assert entry.child is not None
                last = parse_index_page(db.store, entry.child).entries[-1]
                assert (last.key, last.page, last.row) == (entry.key, entry.page, entry.row)
    if oracle_built:
        assert table.index("IX_UniqueBig").unique and table.index("IX_UniqueBig").ignores_nulls
        assert table.index("IX_BigDesc").columns[0][1] is False

    # And the inverse: the row's own values, run through the key codec,
    # must produce exactly the bytes the engine stored -- text included.
    from pyopenvba.access._index import encode_key, leaf_entries

    exact: dict[tuple[int, int], dict[str, object]] = {}
    for page, slot, data in table.raw_rows():
        exact[(page, slot)] = table._exact_values(split_row(table.definition, data))  # pyright: ignore[reportPrivateUsage]
    for index in table.indexes:
        for entry in leaf_entries(db.store, index.real.root_page):
            row = exact[(entry.page, entry.row)]
            values = [row[c.name] for c, _ in index.columns]
            assert encode_key(values, index.columns) == entry.key, (index.name, values)


def test_engine_reads_rows_pyopenvba_wrote(tmp_path: Path) -> None:
    """pyOpenVBA inserts, updates and deletes rows of every scalar type in
    a table the engine built; the engine must then read exactly what we
    meant, keep working with the table (its own inserts afterwards), and
    compact the file without complaint."""
    import datetime as dt
    import uuid
    from decimal import Decimal

    target = tmp_path / "written.accdb"
    shutil.copy(TEMPLATE, target)
    assert oracle("-Command", "build-alltypes", "-Path", str(target), "-Rows", "40") == "ok"

    db = AccessDatabase(target)
    table = db.table("AllTypes")
    inserted: list[dict[str, object]] = []
    for i in range(1, 61):
        row: dict[str, object] = {
            "Flag": i % 2 == 0,
            "Tiny": i % 256,
            "Small": -i * 100,
            "Big": i * 100003,
            "Cash": Decimal(i) / 8,
            "Sgl": i / 4,
            "Dbl": i * 1.5 + 0.25,
            "Stamp": dt.datetime(2000 + i % 20, 1 + i % 12, 1 + i % 28, i % 24, i % 60, (i * 7) % 60),
            "Bin": bytes((k * 7 + i) % 256 for k in range(1 + i % 50)),
            "Txt": f"pyopenvba {i} Привет 日本" if i % 3 else f"ascii {i}",
            "Uid": uuid.UUID(f"{i:08x}-abcd-ef01-2345-6789abcdef01"),
            "Frac": Decimal(i) / 16 * (-1 if i % 5 == 0 else 1),
            "Huge": i * 10_000_000_000,
        }
        if i % 7 == 0:
            row = {"Flag": False}
        row_id = table.insert_row(row)
        inserted.append(dict(row, Id=table.definition.next_autonumber, _rid=row_id))
    # Edit five and delete five of the engine's own rows.
    engine_rows = {r["Id"]: rid for rid, r in table.rows_with_ids() if isinstance(r["Id"], int) and r["Id"] <= 40}
    for k in (3, 8, 13, 18, 23):
        table.update_row(engine_rows[k], {"Txt": f"edited {k}", "Big": -k, "Stamp": dt.datetime(1999, 12, 31, 23, 59, k)})
    for k in (5, 10, 15, 20, 25):
        table.delete_row(engine_rows[k])
    db.save()

    expected = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "AllTypes"))
    assert isinstance(expected, str)
    mine = _engine_dump(AccessDatabase(target), "AllTypes")
    mismatches = _diff(expected, mine)
    assert not mismatches, f"{len(mismatches)} rows differ\n" + "\n".join(mismatches[:5])
    lines = {line.split("\t")[0]: line for line in expected.split("\n")}
    assert len(lines) == 40 - 5 + 60
    for row in inserted:
        cells = dict(cell.split("=", 1) for cell in lines[f"Id={row['Id']}"].split("\t"))
        for name, value in row.items():
            if name in ("Id", "_rid"):
                continue
            if name == "Bin" and isinstance(value, bytes):
                value = value.ljust(50, b"\x00")  # fixed BINARY(50) comes back padded
            assert cells[name] == _format(value), (row["Id"], name, cells[name], value)
    assert "Txt=edited 3" in lines["Id=3"] and "Big=-3" in lines["Id=3"]
    assert "Id=5" not in lines
    # The engine keeps working with what we wrote.
    assert oracle("-Command", "insert-alltypes-more", "-Path", str(target), "-Rows", "10") == "ok"
    after = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "AllTypes"))
    assert isinstance(after, str) and len(after.split("\n")) == 105
    assert oracle("-Command", "compact", "-Path", str(target)) == "ok"
    compacted = AccessDatabase(Path(str(target) + ".compact.accdb"))
    assert compacted.table("AllTypes").row_count == 105
    _check_indexes(compacted)


def test_engine_reads_long_values_and_moved_rows_pyopenvba_wrote(tmp_path: Path) -> None:
    """Memo and OLE values of every storage kind, and a row that outgrew
    its page, written by pyOpenVBA and read back by the engine."""
    target = tmp_path / "memos.accdb"
    shutil.copy(TEMPLATE, target)
    assert oracle("-Command", "build-memos", "-Path", str(target)) == "ok"
    db = AccessDatabase(target)
    table = db.table("Memos")
    memos: dict[str, tuple[str | None, bytes | None]] = {
        "inline": ("short and sweet", bytes(range(50))),
        "inline unicode": ("Привет 日本 мир", None),
        "single": ("s" * 1200 + " end", bytes(range(256)) * 8),
        "chained": ("c" * 6000 + " end", bytes(range(256)) * 30),
        "chained long": ("L" * 30000 + " end", None),
        "nulls": (None, None),
    }
    for label, (memo, blob) in memos.items():
        table.insert_row({"T": label, "M": memo, "O": blob})
    ids = {str(r["T"]): rid for rid, r in table.rows_with_ids()}
    table.update_row(ids["single"], {"M": "replaced with a chained value " * 400})
    table.update_row(ids["chained"], {"M": "now inline"})
    table.delete_row(ids["chained long"])
    db.save()

    expected = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "Memos"))
    assert isinstance(expected, str)
    mine = _engine_dump(AccessDatabase(target), "Memos")
    mismatches = _diff(expected, mine)
    assert not mismatches, f"{len(mismatches)} rows differ\n" + "\n".join(m[:400] for m in mismatches[:3])
    lines = {line.split("\t")[1]: line for line in expected.split("\n")}
    assert lines["T=inline"].split("\t")[2] == "M=short and sweet"
    assert lines["T=inline unicode"].split("\t")[2] == "M=" + "Привет 日本 мир"
    assert lines["T=single"].split("\t")[2] == "M=" + "replaced with a chained value " * 400
    assert lines["T=chained"].split("\t")[2] == "M=now inline"
    assert lines["T=chained"].split("\t")[3] == "O=" + (bytes(range(256)) * 30).hex()
    assert "T=chained long" not in lines
    assert lines["T=nulls"].split("\t")[2:] == ["M=<null>", "O=<null>"]

    # A row grown past its page's free space, then the engine's view of it.
    simple = tmp_path / "grown.accdb"
    shutil.copy(TEMPLATE, simple)
    assert oracle("-Command", "build-simple", "-Path", str(simple)) == "ok"
    db = AccessDatabase(simple)
    table = db.table("Simple")
    home = table.data_pages()[0]
    while table.insert_row({"N": 1, "T": "filler " * 6}).page == home:
        pass
    rid = next(rid for rid, r in table.rows_with_ids() if r["Id"] == 2)
    table.update_row(rid, {"T": "this row grew and had to move elsewhere"})
    db.save()
    dumped = json.loads(oracle("-Command", "dump", "-Path", str(simple), "-Table", "Simple"))
    assert isinstance(dumped, str)
    assert "Id=2\tN=20\tT=this row grew and had to move elsewhere" in dumped
    assert _diff(dumped, _engine_dump(AccessDatabase(simple), "Simple")) == []
    assert oracle("-Command", "compact", "-Path", str(simple)) == "ok"


def test_engine_uses_a_table_pyopenvba_created(tmp_path: Path) -> None:
    """A table of every column type created from nothing by pyOpenVBA: the
    engine inserts into it, reads it back, and compacts the file."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    target = tmp_path / "created.accdb"
    shutil.copy(TEMPLATE, target)
    db = AccessDatabase(target)
    table = db.create_table(
        "AllTypes",
        [
            ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Flag", "Boolean"), ColumnSpec("Tiny", "Byte"),
            ColumnSpec("Small", "Integer"), ColumnSpec("Big", "Long"), ColumnSpec("Cash", "Currency"),
            ColumnSpec("Sgl", "Single"), ColumnSpec("Dbl", "Double"), ColumnSpec("Stamp", "DateTime"),
            ColumnSpec("Bin", "Binary", size=50), ColumnSpec("Txt", "Text", size=100, compressed=False),
            ColumnSpec("Blob", "OLE"), ColumnSpec("Story", "Memo", compressed=False), ColumnSpec("Uid", "GUID"),
            ColumnSpec("Frac", "Decimal", size=(18, 4)), ColumnSpec("Huge", "BigInt"),
        ],
        [IndexSpec("PrimaryKey", ("Id",), primary=True), IndexSpec("IX_Txt", ("Txt",)), IndexSpec("IX_FlagTiny", ("Flag", ("Tiny", False)))],
    )
    table.insert_row({"Flag": True, "Tiny": 3, "Txt": "made by pyopenvba", "Story": "m" * 5000, "Blob": bytes(range(100))})
    db.save()
    # The engine adds rows of its own through SQL and reads the whole table.
    assert oracle("-Command", "insert-alltypes-more", "-Path", str(target), "-Rows", "25") == "ok"
    dumped = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "AllTypes"))
    assert isinstance(dumped, str)
    lines = dumped.split("\n")
    assert len(lines) == 26
    assert lines[0].startswith("Id=1\tFlag=True\tTiny=3\t") and "Txt=made by pyopenvba" in lines[0]
    assert "Story=" + "m" * 5000 in lines[0] and "Blob=" + bytes(range(100)).hex() in lines[0]
    assert _diff(dumped, _engine_dump(AccessDatabase(target), "AllTypes")) == []
    assert oracle("-Command", "compact", "-Path", str(target)) == "ok"
    compacted = AccessDatabase(Path(str(target) + ".compact.accdb"))
    assert compacted.table("AllTypes").row_count == 26
    _check_indexes(compacted, oracle_built=False)


def _differing_pages(ours: bytes, engine: bytes) -> list[int]:
    """Pages after page 0 (whose statement counter is not reproduced) that
    differ between two databases of the same size."""
    assert len(ours) == len(engine), (len(ours), len(engine))
    return [
        n for n in range(1, len(ours) // 4096)
        if ours[n * 4096 : (n + 1) * 4096] != engine[n * 4096 : (n + 1) * 4096]
    ]


def _describe_pages(ours: bytes, engine: bytes, pages: list[int]) -> str:
    """The first differing bytes of each listed page, for the assertion."""
    lines: list[str] = []
    for n in pages[:3]:
        a, b = ours[n * 4096 : (n + 1) * 4096], engine[n * 4096 : (n + 1) * 4096]
        offsets = [i for i in range(4096) if a[i] != b[i]]
        lo = offsets[0]
        lines.append(f"page {n} (type {b[0]:#04x}, owner {int.from_bytes(b[4:8], 'little')}): {len(offsets)} bytes from {lo}; ours {a[lo:lo + 16].hex(' ')} / engine {b[lo:lo + 16].hex(' ')}")
    return "; ".join(lines)


def _mask_lval_stamps(blob: bytearray, chain_lengths: tuple[int, ...]) -> None:
    """Zero the per-session stamp the engine puts at offset 8 of every LVAL
    page and in the 12-byte definition of a chained value of each given
    length -- the one thing about a chain that is not reproducible."""
    for n in range(1, len(blob) // 4096):
        start = n * 4096
        page = blob[start : start + 4096]
        if page[0] == 1 and bytes(page[4:8]) == b"LVAL":
            blob[start + 8 : start + 12] = b"\0\0\0\0"
        for length in chain_lengths:
            marker = length.to_bytes(3, "little") + b"\x00"  # length, kind chained
            pos = page.find(marker)
            while pos >= 0:
                blob[start + pos + 8 : start + pos + 12] = b"\0\0\0\0"
                pos = page.find(marker, pos + 1)


def _stamp_between(blob: bytes, first: float, last: float) -> float:
    """The DateUpdate DAO wrote at an intermediate step, recovered from the
    row version its bytes outlived: the first stored double strictly
    between the two catalog stamps, else the last stamp."""
    import struct

    lo, hi = min(first, last), max(first, last)
    for offset in range(4096, len(blob) - 8):
        value = struct.unpack_from("<d", blob, offset)[0]
        if lo < value < hi:
            return value
    return last


def _catalog_entry(path: Path, name: str) -> CatalogEntry:
    return next(e for e in AccessDatabase(path).catalog() if e.name == name)


def test_create_and_drop_table_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """CREATE TABLE with a named primary key plus CREATE INDEX, then DROP
    TABLE, by the engine and by pyOpenVBA on copies of the blank database.
    Only page 0 and the two catalog timestamps may differ."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    differing = _differing_pages
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    assert oracle("-Command", "create-keyed", "-Path", str(theirs)) == "ok"
    entry = _catalog_entry(theirs, "Simple")
    assert isinstance(entry.date_create, dt.datetime)
    db = AccessDatabase(TEMPLATE)
    db.create_table(
        "Simple",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=50, compressed=False)],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
        created=entry.date_create_serial,
        updated=entry.date_update_serial,
    )
    assert not (d := differing(db.to_bytes(), theirs.read_bytes())), f"create: pages differ from the engine's: {d}"

    # CREATE INDEX: the engine also re-stamps the catalog row, which moves
    # it when its page cannot hold a fresh copy.
    assert oracle("-Command", "index-simple", "-Path", str(theirs)) == "ok"
    entry = _catalog_entry(theirs, "Simple")
    assert isinstance(entry.date_update, dt.datetime)
    db.create_index("Simple", IndexSpec("IX_N", ("N",)), updated=entry.date_update_serial)
    assert not (d := differing(db.to_bytes(), theirs.read_bytes())), f"index: pages differ from the engine's: {d}"

    # DROP TABLE, exact to the byte: the engine's order of releasing maps
    # and killing rows decides which stale bytes remain, and that order is
    # reproduced.
    assert oracle("-Command", "drop-simple", "-Path", str(theirs)) == "ok"
    db.drop_table("Simple")
    assert not (d := differing(db.to_bytes(), theirs.read_bytes())), f"drop: pages differ from the engine's: {d}"


def test_definitions_over_one_page_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """Definitions that run past one page: a 150-column table (two pages)
    with rows, a four-page and a three-page table, then CREATE INDEX on
    both, which rewrites each definition onto a fresh continuation chain.
    Every page but page 0 must match the engine's."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same(step: str, db: AccessDatabase) -> None:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"

    def text_columns(names: list[str]) -> list[ColumnSpec]:
        return [ColumnSpec("Id", "Long", autonumber=True), *(ColumnSpec(n, "Text", size=20, compressed=False) for n in names)]

    def ddl(table: str, names: list[str]) -> str:
        return f"CREATE TABLE {table} (Id AUTOINCREMENT CONSTRAINT PrimaryKey PRIMARY KEY, " + ", ".join(f"{n} TEXT(20)" for n in names) + ")"

    def long_names(count: int) -> list[str]:
        return [f"Column{k:03d}".ljust(30, "_") for k in range(1, count)]

    def stamps(name: str) -> dict[str, object]:
        entry = _catalog_entry(theirs, name)
        return {"created": entry.date_create_serial, "updated": entry.date_update_serial}

    key = [IndexSpec("PrimaryKey", ("Id",), primary=True)]
    db = AccessDatabase(TEMPLATE)

    wide_names = [f"Col{k:03d}" for k in range(1, 151)]
    engine_runs(ddl("Wide", wide_names))
    wide = db.create_table("Wide", text_columns(wide_names), key, **stamps("Wide"))
    assert len(wide.definition.pages) == 2
    same("wide table", db)

    rows = [{n: f"r{r}c{k}" for k, n in enumerate(wide_names, start=1) if (k + r) % 4} for r in (1, 2, 3)]
    engine_runs(*(f"INSERT INTO Wide ({', '.join(row)}) VALUES ({', '.join(repr(v) for v in row.values())})" for row in rows))
    for row in rows:
        wide.insert_row(row)
    same("wide rows", db)

    engine_runs(ddl("Four", long_names(151)), ddl("Three", long_names(92)))
    four = db.create_table("Four", text_columns(long_names(151)), key, **stamps("Four"))
    three = db.create_table("Three", text_columns(long_names(92)), key, **stamps("Three"))
    assert len(four.definition.pages) == 4 and len(three.definition.pages) == 2
    same("four- and two-page tables", db)

    engine_runs("CREATE INDEX IX ON Four (Column001_____________________)", "CREATE INDEX IX ON Three (Column001_____________________)")
    db.create_index("Four", IndexSpec("IX", ("Column001_____________________",)), updated=_catalog_entry(theirs, "Four").date_update)
    db.create_index("Three", IndexSpec("IX", ("Column001_____________________",)), updated=_catalog_entry(theirs, "Three").date_update)
    assert len(db.table("Three").definition.pages) == 3
    same("indexes rewriting the definitions", db)

    from test_access_write import check_indexes

    for name in ("Wide", "Four", "Three"):
        check_indexes(db.table(name))


def test_deletes_retire_and_truncate_releases_pages_as_the_engine_does(tmp_path: Path) -> None:
    """Filtered deletes retire each page they empty (type 0x09, released,
    out of the maps) except a table's first data page, and rejoin pages
    that lost rows to the free-space map; DELETE FROM without a filter
    releases every page untouched and resets the index.  Byte for byte
    against DAO on the same sequence of tables.  Every oracle call is its
    own DAO session, so the database is reopened between steps: pages a
    session released come back into use only in the next one."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        return AccessDatabase(ours)

    def stamps(name: str) -> dict[str, object]:
        entry = _catalog_entry(theirs, name)
        return {"created": entry.date_create_serial, "updated": entry.date_update_serial}

    def delete_where(table: Table, keep: Callable[[int], bool]) -> None:
        for row_id, row in list(table.rows_with_ids()):
            if not keep(int(row["Id"])):  # pyright: ignore[reportArgumentType]
                table.delete_row(row_id)

    db = AccessDatabase(TEMPLATE)
    key = [IndexSpec("PK", ("Id",), primary=True)]
    text_memo = [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("T", "Text", size=255, compressed=False), ColumnSpec("M", "Memo", compressed=False)]

    # R: four data pages and three single-row long values.
    memo = {i: f"m{i:02d}" * 500 for i in (3, 10, 20)}
    engine_runs(
        "CREATE TABLE R (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, T TEXT(255), M MEMO)",
        *(f"INSERT INTO R (T, M) VALUES ('{f'{chr(116)}{i:02d}' * 80}', '{memo.get(i, 'short')}')" for i in range(1, 25)),
    )
    r = db.create_table("R", text_memo, key, **stamps("R"))
    for i in range(1, 25):
        r.insert_row({"T": f"t{i:02d}" * 80, "M": memo.get(i, "short")})
    db = same_then_reopen("R built", db)
    engine_runs("DELETE FROM R WHERE Id > 12")
    delete_where(db.table("R"), lambda i: i <= 12)
    db = same_then_reopen("R filtered delete", db)
    engine_runs("DELETE FROM R")
    db.table("R").truncate()
    db = same_then_reopen("R truncated", db)

    # S: a table whose only page is emptied by a filtered delete stays.
    engine_runs("CREATE TABLE S (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG)", *(f"INSERT INTO S (N) VALUES ({n})" for n in (1, 2, 3)))
    s_table = db.create_table("S", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long")], key, **stamps("S"))
    for n in (1, 2, 3):
        s_table.insert_row({"N": n})
    db = same_then_reopen("S built", db)
    engine_runs("DELETE FROM S WHERE Id > 0")
    delete_where(db.table("S"), lambda i: False)
    db = same_then_reopen("S emptied by filter", db)

    # U: the first data page emptied while the second keeps rows; V: a
    # column's only LVAL page emptied.
    engine_runs(
        "CREATE TABLE U (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, T TEXT(255))",
        *(f"INSERT INTO U (T) VALUES ('{f'{chr(117)}{i:02d}' * 80}')" for i in range(1, 11)),
        "CREATE TABLE V (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)",
        "INSERT INTO V (M) VALUES ('" + "v" * 1500 + "')",
        "INSERT INTO V (M) VALUES ('tiny')",
    )
    u = db.create_table("U", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("T", "Text", size=255, compressed=False)], key, **stamps("U"))
    for i in range(1, 11):
        u.insert_row({"T": f"u{i:02d}" * 80})
    v = db.create_table("V", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("M", "Memo", compressed=False)], key, **stamps("V"))
    v.insert_row({"M": "v" * 1500})
    v.insert_row({"M": "tiny"})
    db = same_then_reopen("U and V built", db)
    engine_runs("DELETE FROM U WHERE Id <= 7", "DELETE FROM V WHERE Id = 1")
    delete_where(db.table("U"), lambda i: i > 7)
    delete_where(db.table("V"), lambda i: i != 1)
    db = same_then_reopen("U partly and V's long value deleted", db)
    engine_runs("DELETE FROM U WHERE Id <= 8")
    delete_where(db.table("U"), lambda i: i > 8)
    db = same_then_reopen("U's first page emptied", db)
    for name in ("R", "S", "U", "V"):
        check_indexes(db.table(name))


def test_long_value_placement_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """Where single-row long values land: the page last written when it
    has room, else the first listed page, else a fresh one; pages stay
    listed while more than 256 bytes are free; an update stores the new
    value before freeing the old; a delete re-lists its page; a freed
    chain's pages come back only when the chain predates the session.
    One DAO session per step, byte for byte but for the chains' per-session
    stamps."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = bytearray(db.to_bytes()), bytearray(theirs.read_bytes())
        _mask_lval_stamps(ours, (10000,))
        _mask_lval_stamps(engine, (10000,))
        assert not (d := _differing_pages(bytes(ours), bytes(engine))), f"{step}: pages differ from the engine's: {_describe_pages(bytes(ours), bytes(engine), d)}"
        return AccessDatabase(db.to_bytes())

    def stamps(name: str) -> dict[str, object]:
        entry = _catalog_entry(theirs, name)
        return {"created": entry.date_create_serial, "updated": entry.date_update_serial}

    def by_id(table: Table, wanted: int) -> RowId:
        return next(rid for rid, row in table.rows_with_ids() if row["Id"] == wanted)

    memo = [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("M", "Memo", compressed=False)]
    key = [IndexSpec("PK", ("Id",), primary=True)]
    db = AccessDatabase(TEMPLATE)

    # One session: two 3000-byte values on two pages, then two 900-byte ones.
    values = ["a" * 1500, "b" * 1500, "c" * 450, "d" * 450]
    engine_runs("CREATE TABLE L3 (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)", *(f"INSERT INTO L3 (M) VALUES ('{v}')" for v in values))
    l3 = db.create_table("L3", memo, key, **stamps("L3"))
    for v in values:
        l3.insert_row({"M": v})
    db = same_then_reopen("cursor placement", db)

    # Pages left with 256 and 258 bytes free, by two values each.
    for free in (256, 258):
        second = "b" * ((2778 - free) // 2)
        engine_runs(f"CREATE TABLE F{free} (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)", f"INSERT INTO F{free} (M) VALUES ('{'a' * 650}')", f"INSERT INTO F{free} (M) VALUES ('{second}')")
        table = db.create_table(f"F{free}", memo, key, **stamps(f"F{free}"))
        table.insert_row({"M": "a" * 650})
        table.insert_row({"M": second})
        db = same_then_reopen(f"listing threshold at {free}", db)

    # Six 1300-byte values, a delete, an insert into the hole, an update
    # of a value on the other page, and an insert that fits nowhere.
    engine_runs("CREATE TABLE L (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)", *(f"INSERT INTO L (M) VALUES ('{chr(97 + i) * 650}')" for i in range(6)))
    l_table = db.create_table("L", memo, key, **stamps("L"))
    for i in range(6):
        l_table.insert_row({"M": chr(97 + i) * 650})
    db = same_then_reopen("L built", db)
    engine_runs("DELETE FROM L WHERE Id = 2")
    db.table("L").delete_row(by_id(db.table("L"), 2))
    db = same_then_reopen("L delete", db)
    engine_runs("INSERT INTO L (M) VALUES ('" + "n" * 250 + "')")
    db.table("L").insert_row({"M": "n" * 250})
    db = same_then_reopen("L insert into the hole", db)
    engine_runs("UPDATE L SET M = '" + "u" * 400 + "' WHERE Id = 4")
    db.table("L").update_row(by_id(db.table("L"), 4), {"M": "u" * 400})
    db = same_then_reopen("L update", db)
    engine_runs("INSERT INTO L (M) VALUES ('" + "p" * 1400 + "')")
    db.table("L").insert_row({"M": "p" * 1400})
    db = same_then_reopen("L oversized insert", db)

    # A 10 KB chain created and freed in one session gets fresh pages for
    # its replacement; one created in an earlier session gives its pages
    # back to the replacement in order.
    ten_k, ten_m = "k" * 5000, "m" * 5000
    engine_runs("CREATE TABLE K1 (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)", f"INSERT INTO K1 (M) VALUES ('{ten_k}')", "DELETE FROM K1 WHERE Id = 1", f"INSERT INTO K1 (M) VALUES ('{ten_m}')")
    k1 = db.create_table("K1", memo, key, **stamps("K1"))
    k1.delete_row(k1.insert_row({"M": ten_k}))
    k1.insert_row({"M": ten_m})
    db = same_then_reopen("K1 same-session chain", db)
    engine_runs("CREATE TABLE K2 (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, M MEMO)", f"INSERT INTO K2 (M) VALUES ('{ten_k}')")
    db.create_table("K2", memo, key, **stamps("K2")).insert_row({"M": ten_k})
    db = same_then_reopen("K2 built", db)
    engine_runs("DELETE FROM K2 WHERE Id = 1", f"INSERT INTO K2 (M) VALUES ('{ten_m}')")
    k2 = db.table("K2")
    k2.delete_row(by_id(k2, 1))
    k2.insert_row({"M": ten_m})
    db = same_then_reopen("K2 older chain reused", db)
    for name in ("L3", "F256", "F258", "L", "K1", "K2"):
        check_indexes(db.table(name))


def test_relationships_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """Two tables, then ALTER TABLE ADD CONSTRAINT ... FOREIGN KEY twice
    (a second child), each step its own DAO session, byte for byte."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        return AccessDatabase(ours)

    def stamps(name: str) -> dict[str, object]:
        entry = _catalog_entry(theirs, name)
        return {"created": entry.date_create_serial, "updated": entry.date_update_serial}

    key = [IndexSpec("PK", ("Id",), primary=True)]
    db = AccessDatabase(TEMPLATE)
    engine_runs(
        "CREATE TABLE Parent (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Name TEXT(50))",
        "CREATE TABLE Child (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, ParentId LONG, Remark TEXT(50))",
    )
    db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50, compressed=False)], key, **stamps("Parent"))
    db.create_table("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long"), ColumnSpec("Remark", "Text", size=50, compressed=False)], key, **stamps("Child"))
    db = same_then_reopen("two tables", db)
    for child in ("Child", "Child2"):
        if child == "Child2":
            engine_runs("CREATE TABLE Child2 (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, ParentId LONG)")
            db.create_table("Child2", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key, **stamps("Child2"))
            db = same_then_reopen("second child", db)
        name = f"FK_{child}_Parent"
        engine_runs(f"ALTER TABLE {child} ADD CONSTRAINT {name} FOREIGN KEY (ParentId) REFERENCES Parent (Id)")
        db.create_relationship(
            name, child, ("ParentId",), "Parent", ("Id",),
            created=_catalog_entry(theirs, name).date_create_serial,
            table_updated=stamps(child)["updated"],
            referenced_updated=stamps("Parent")["updated"],
        )
        db = same_then_reopen(f"constraint on {child}", db)
    assert [r.name for r in db.relationships()][-2:] == ["FK_Child_Parent", "FK_Child2_Parent"]
    engine_runs("ALTER TABLE Child DROP CONSTRAINT FK_Child_Parent")
    db.drop_relationship("FK_Child_Parent", table_updated=stamps("Child")["updated"], referenced_updated=stamps("Parent")["updated"])
    db = same_then_reopen("constraint dropped", db)
    assert [r.name for r in db.relationships()][-1:] == ["FK_Child2_Parent"]
    # Renaming the referenced table follows into the relationship rows.
    script.write_text("Parents" + chr(10), encoding="ascii")
    assert oracle("-Command", "rename-table", "-Path", str(theirs), "-Table", "Parent", "-SqlFile", str(script)) == "ok"
    db.rename_table("Parent", "Parents", updated=stamps("Parents")["updated"])
    db = same_then_reopen("table renamed", db)
    assert db.relationships()[-1].referenced_table == "Parents"
    # Renaming the foreign-key column follows into the relationship row.
    script.write_text("ParentId" + chr(10) + "PId" + chr(10), encoding="ascii")
    assert oracle("-Command", "rename-column", "-Path", str(theirs), "-Table", "Child2", "-SqlFile", str(script)) == "ok"
    db.table("Child2").rename_column("ParentId", "PId", updated=stamps("Child2")["updated"])
    db = same_then_reopen("column renamed", db)
    assert db.relationships()[-1].columns == ("PId",)
    for name in ("Parents", "Child", "Child2", "MSysRelationships", "MSysObjects", "MSysACEs"):
        check_indexes(db.table(name))


def test_properties_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """DAO appends a table Description, then a field Caption and Description;
    three set_properties calls must leave identical pages (each append is
    one rewrite of the LvProp value)."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    script.write_text("CREATE TABLE Parent (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Name TEXT(50))" + chr(10), encoding="ascii")
    assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"
    entry = _catalog_entry(theirs, "Parent")
    db = AccessDatabase(TEMPLATE)
    db.create_table(
        "Parent",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50, compressed=False)],
        [IndexSpec("PK", ("Id",), primary=True)],
        created=entry.date_create_serial,
        updated=entry.date_update_serial,
    )
    assert not _differing_pages(db.to_bytes(), theirs.read_bytes())
    db = AccessDatabase(db.to_bytes())
    assert oracle("-Command", "set-props", "-Path", str(theirs), "-Table", "Parent") == "ok"
    table = db.table("Parent")
    table.set_properties({"Description": "Table described by DAO"})
    table.set_properties({"Caption": "Name shown"}, column="Name")
    table.set_properties({"Description": "Field described by DAO"}, column="Name")
    ours, engine = db.to_bytes(), theirs.read_bytes()
    assert not (d := _differing_pages(ours, engine)), f"properties: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
    assert table.column_properties("Name") == {"Caption": "Name shown", "Description": "Field described by DAO"}
    # Renaming the column carries its property block along.
    db = AccessDatabase(db.to_bytes())
    script.write_text("Name" + chr(10) + "FullName" + chr(10), encoding="ascii")
    assert oracle("-Command", "rename-column", "-Path", str(theirs), "-Table", "Parent", "-SqlFile", str(script)) == "ok"
    db.table("Parent").rename_column("Name", "FullName", updated=_catalog_entry(theirs, "Parent").date_update_serial)
    ours, engine = db.to_bytes(), theirs.read_bytes()
    assert not (d := _differing_pages(ours, engine)), f"column renamed: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
    assert db.table("Parent").column_properties("FullName") == {"Caption": "Name shown", "Description": "Field described by DAO"}


def test_saved_queries_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """Four CreateQueryDef calls -- a plain select, a joined DISTINCT TOP
    GROUP BY HAVING ORDER BY DESC query, a parameter query and a DELETE --
    each its own DAO session, byte for byte."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    script.write_text(
        "CREATE TABLE Parent (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Name TEXT(50))" + chr(10)
        + "CREATE TABLE Child (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, ParentId LONG, Remark TEXT(50))" + chr(10)
        + "CREATE TABLE Sales (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Region TEXT(20), Quarter TEXT(10), Amount CURRENCY)" + chr(10),
        encoding="ascii",
    )
    assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"
    key = [IndexSpec("PK", ("Id",), primary=True)]
    db = AccessDatabase(TEMPLATE)
    for name, columns in (
        ("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50, compressed=False)]),
        ("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long"), ColumnSpec("Remark", "Text", size=50, compressed=False)]),
        ("Sales", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Region", "Text", size=20, compressed=False), ColumnSpec("Quarter", "Text", size=10, compressed=False), ColumnSpec("Amount", "Currency")]),
    ):
        entry = _catalog_entry(theirs, name)
        db.create_table(name, columns, key, created=entry.date_create_serial, updated=entry.date_update_serial)
    assert not _differing_pages(db.to_bytes(), theirs.read_bytes())
    statements = {
        "ParentsAbove1": "SELECT Parent.Id, Parent.Name FROM Parent WHERE Parent.Id > 1 ORDER BY Parent.Name",
        "Q2": "SELECT DISTINCT TOP 5 Child.Id, Child.ParentId AS P FROM Child INNER JOIN Parent ON Child.ParentId = Parent.Id WHERE Child.Id > 2 GROUP BY Child.Id, Child.ParentId HAVING Count(*) > 0 ORDER BY Child.Id DESC",
        "Q3": "PARAMETERS [Which] Long; SELECT * FROM Parent WHERE Id = [Which]",
        "Q4": "DELETE FROM Child WHERE Id < 0",
        "Q5": "UPDATE Parent SET Parent.Name = 'x' WHERE Parent.Id = 3",
        "Q6": "INSERT INTO Child ( ParentId, Remark ) SELECT Parent.Id, Parent.Name FROM Parent",
        "Q7": "SELECT Parent.Id, Parent.Name INTO Copied FROM Parent",
        "Q8": "SELECT Parent.Id FROM Parent UNION SELECT Child.Id FROM Child",
        # Crosstabs: the value column, the row headings, and the pivot last,
        # with and without an IN list, a TOP, a join and a parameter.
        "X1": "TRANSFORM Sum(Amount) AS Total SELECT Region FROM Sales GROUP BY Region PIVOT Quarter",
        "X2": "TRANSFORM Count(*) AS N SELECT Region, Sum(Amount) AS Tot FROM Sales WHERE Amount > 5 GROUP BY Region ORDER BY Region PIVOT Quarter IN ('Q1', 'Q2', 'Q3', 'Q4')",
        "X3": "TRANSFORM Sum(Amount) SELECT TOP 5 Region FROM Sales GROUP BY Region PIVOT Quarter",
        "X4": "TRANSFORM Sum(s.Amount) AS T SELECT c.Remark FROM Sales AS s INNER JOIN Child AS c ON s.Region = c.Remark GROUP BY c.Remark PIVOT s.Quarter",
        "X5": "PARAMETERS [Low] Currency; TRANSFORM Sum(Amount) SELECT Region FROM Sales WHERE Amount > [Low] GROUP BY Region PIVOT Quarter",
    }
    for name, sql in statements.items():
        script.write_text(sql + chr(10), encoding="ascii")
        assert oracle("-Command", "create-query", "-Path", str(theirs), "-Table", name, "-SqlFile", str(script)) == "ok"
        db = AccessDatabase(db.to_bytes())
        entry = _catalog_entry(theirs, name)
        assert entry.date_create_serial is not None and entry.date_update_serial is not None
        db.create_query(
            name, sql,
            created=entry.date_create_serial,
            owner_updated=_stamp_between(theirs.read_bytes(), entry.date_create_serial, entry.date_update_serial),
            updated=entry.date_update_serial,
        )
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{name}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        assert db.query(name).sql == sql
    assert oracle("-Command", "delete-query", "-Path", str(theirs), "-Table", "Q3") == "ok"
    db = AccessDatabase(db.to_bytes())
    db.drop_query("Q3")
    ours, engine = db.to_bytes(), theirs.read_bytes()
    assert not (d := _differing_pages(ours, engine)), f"delete: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
    for name in ("MSysObjects", "MSysQueries", "MSysACEs"):
        check_indexes(db.table(name))


def test_alter_table_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """ADD COLUMN (fixed, variable, and fixed again after a drop) and DROP
    COLUMN (fixed and variable) on a table with rows, each its own DAO
    session, byte for byte."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = bytearray(db.to_bytes()), bytearray(theirs.read_bytes())
        _mask_lval_stamps(ours, (6000,))  # the memo chain's per-session stamp
        _mask_lval_stamps(engine, (6000,))
        assert not (d := _differing_pages(bytes(ours), bytes(engine))), f"{step}: pages differ from the engine's: {_describe_pages(bytes(ours), bytes(engine), d)}"
        return AccessDatabase(db.to_bytes())

    def updated() -> object:
        return _catalog_entry(theirs, "W").date_update_serial

    engine_runs("CREATE TABLE W (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG, T TEXT(40))", *(f"INSERT INTO W (N, T) VALUES ({i}, 'row {i}')" for i in range(1, 11)))
    entry = _catalog_entry(theirs, "W")
    db = AccessDatabase(TEMPLATE)
    table = db.create_table(
        "W",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=40, compressed=False)],
        [IndexSpec("PK", ("Id",), primary=True)],
        created=entry.date_create_serial,
        updated=entry.date_update_serial,
    )
    for i in range(1, 11):
        table.insert_row({"N": i, "T": f"row {i}"})
    db = same_then_reopen("built", db)
    steps: list[tuple[str, Callable[[Table], object]]] = [
        ("ALTER TABLE W ADD COLUMN Extra LONG", lambda t: t.add_column(ColumnSpec("Extra", "Long"), updated=updated())),
        ("ALTER TABLE W ADD COLUMN Remark TEXT(30)", lambda t: t.add_column(ColumnSpec("Remark", "Text", size=30, compressed=False), updated=updated())),
        ("ALTER TABLE W DROP COLUMN Extra", lambda t: t.drop_column("Extra", updated=updated())),
        ("ALTER TABLE W ADD COLUMN Again LONG", lambda t: t.add_column(ColumnSpec("Again", "Long"), updated=updated())),
        ("ALTER TABLE W DROP COLUMN T", lambda t: t.drop_column("T", updated=updated())),
        ("ALTER TABLE W ADD COLUMN Notes MEMO", lambda t: t.add_column(ColumnSpec("Notes", "Memo", compressed=False), updated=updated())),
        ("INSERT INTO W (N, Notes) VALUES (99, '" + "n" * 3000 + "')", lambda t: t.insert_row({"N": 99, "Notes": "n" * 3000})),
        ("ALTER TABLE W DROP COLUMN Notes", lambda t: t.drop_column("Notes", updated=updated())),
        ("ALTER TABLE W ALTER COLUMN Remark TEXT(60)", lambda t: t.alter_column("Remark", ColumnSpec("Remark", "Text", size=60, compressed=False), updated=updated())),
        ("ALTER TABLE W ALTER COLUMN N DOUBLE", lambda t: t.alter_column("N", ColumnSpec("N", "Double"), updated=updated())),
    ]
    for statement, ours in steps:
        engine_runs(statement)
        ours(db.table("W"))
        db = same_then_reopen(statement[:40], db)
    assert db.table("W").column_names == ["Id", "N", "Remark", "Again"] and db.table("W").row_count == 11
    check_indexes(db.table("W"))


def test_a_second_map_page_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """A table whose 58 usage-map rows overflow its map page: 32 indexes
    and 12 Memo columns, created in one DAO session, byte for byte."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    ddl = "CREATE TABLE Many (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, " + ", ".join(f"C{i:02d} LONG" for i in range(1, 33)) + ", " + ", ".join(f"M{i:02d} MEMO" for i in range(1, 13)) + ")"
    statements = [ddl] + [f"CREATE INDEX IX{i:02d} ON Many (C{i:02d})" for i in range(1, 32)]
    script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
    assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"
    entry = _catalog_entry(theirs, "Many")
    db = AccessDatabase(TEMPLATE)
    columns = [ColumnSpec("Id", "Long", autonumber=True), *(ColumnSpec(f"C{i:02d}", "Long") for i in range(1, 33)), *(ColumnSpec(f"M{i:02d}", "Memo", compressed=False) for i in range(1, 13))]
    db.create_table("Many", columns, [IndexSpec("PK", ("Id",), primary=True)], created=entry.date_create_serial, updated=entry.date_create_serial)
    for i in range(1, 32):
        db.create_index("Many", IndexSpec(f"IX{i:02d}", (f"C{i:02d}",)), updated=entry.date_create_serial)
    # The last CREATE INDEX stamps the row; earlier stamps are overwritten.
    rid = next(rid for rid, row in db.table("MSysObjects").rows_with_ids() if row["Name"] == "Many")
    db.table("MSysObjects").update_row(rid, {"DateUpdate": entry.date_update_serial})
    ours, engine = db.to_bytes(), theirs.read_bytes()
    assert not (d := _differing_pages(ours, engine)), f"second map page: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
    check_indexes(db.table("Many"))


def test_growth_past_512_pages_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """450 memo rows carry the file from 121 to about 570 pages.  The global
    usage map, the table's maps and the column's maps all have to grow
    their inline bitmaps and re-base, exactly as the engine does."""
    theirs = tmp_path / "grown_theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    assert oracle("-Command", "build-memos", "-Path", str(theirs)) == "ok"
    ours_path = tmp_path / "grown_ours.accdb"
    shutil.copy(theirs, ours_path)
    assert oracle("-Command", "grow-memos", "-Path", str(theirs), "-Rows", "450") == "ok"

    db = AccessDatabase(ours_path)
    table = db.table("Memos")
    for i in range(1, 451):
        table.insert_row({"T": f"m{i}", "M": "a" * 1600})
    db.save()
    ours = db.to_bytes()
    engine = theirs.read_bytes()
    assert len(ours) == len(engine), (len(ours) // 4096, len(engine) // 4096)
    assert len(ours) // 4096 > 512
    different = [
        n for n in range(1, len(ours) // 4096)
        if ours[n * 4096 : (n + 1) * 4096] != engine[n * 4096 : (n + 1) * 4096]
    ]
    assert not different, f"pages differ from the engine's: {different[:20]}"
    # The engine keeps working with the grown file.
    assert oracle("-Command", "grow-memos", "-Path", str(ours_path), "-Rows", "20") == "ok"
    assert oracle("-Command", "compact", "-Path", str(ours_path)) == "ok"
    assert AccessDatabase(Path(str(ours_path) + ".compact.accdb")).table("Memos").row_count == 3 + 450 + 20


def test_memo_inserts_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """A single-page and a chained memo written by pyOpenVBA against the
    engine's own, page for page.  The chained value carries a per-session
    stamp in its definition and on its first page, which is masked."""
    for chars in (200, 3000):
        base = tmp_path / f"memo{chars}.accdb"
        shutil.copy(TEMPLATE, base)
        assert oracle("-Command", "build-memos", "-Path", str(base)) == "ok"
        theirs = tmp_path / f"theirs{chars}.accdb"
        shutil.copy(base, theirs)
        assert oracle("-Command", "insert-memo", "-Path", str(theirs), "-Rows", str(chars)) == "ok"
        db = AccessDatabase(base)
        db.table("Memos").insert_row({"T": f"memo {chars}", "M": "a" * chars})
        ours = bytearray(db.to_bytes())
        engine = bytearray(theirs.read_bytes())
        assert len(ours) == len(engine)
        if chars == 3000:
            # Mask the stamp: 4 bytes in the row's definition and 4 at
            # offset 8 of the first LVAL page.
            for blob in (ours, engine):
                for n in range(1, len(blob) // 4096):
                    page = blob[n * 4096 : (n + 1) * 4096]
                    if page[0] == 1 and bytes(page[4:8]) == b"LVAL":
                        blob[n * 4096 + 8 : n * 4096 + 12] = b"\0\0\0\0"
                for n in range(1, len(blob) // 4096):
                    page = blob[n * 4096 : (n + 1) * 4096]
                    marker = b"\x70\x17\x00\x00"  # length 6000, kind chained
                    pos = page.find(marker)
                    if pos >= 0:
                        blob[n * 4096 + pos + 8 : n * 4096 + pos + 12] = b"\0\0\0\0"
        different = [
            n for n in range(1, len(ours) // 4096)
            if ours[n * 4096 : (n + 1) * 4096] != engine[n * 4096 : (n + 1) * 4096]
        ]
        assert not different, f"{chars}-char memo: pages differ from the engine's: {different}"


def test_single_insert_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """The same row inserted by the engine and by pyOpenVBA into copies of
    one database must leave identical pages -- page 0 aside, where the
    engine bumps a counter that is left alone."""
    base = tmp_path / "base.accdb"
    shutil.copy(TEMPLATE, base)
    assert oracle("-Command", "build-simple", "-Path", str(base)) == "ok"
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(base, theirs)
    assert oracle("-Command", "insert-simple", "-Path", str(theirs), "-Rows", "60") == "ok"

    db = AccessDatabase(base)
    db.table("Simple").insert_row({"N": 60, "T": "inserted 60"})
    ours = db.to_bytes()
    engine = theirs.read_bytes()
    assert len(ours) == len(engine)
    different = [
        n for n in range(1, len(ours) // 4096)
        if ours[n * 4096 : (n + 1) * 4096] != engine[n * 4096 : (n + 1) * 4096]
    ]
    assert not different, f"pages differ from the engine's: {different}"

    # And a delete.
    theirs2 = tmp_path / "theirs2.accdb"
    shutil.copy(base, theirs2)
    assert oracle("-Command", "delete-simple", "-Path", str(theirs2), "-Rows", "3") == "ok"
    db = AccessDatabase(base)
    table = db.table("Simple")
    rid = next(rid for rid, r in table.rows_with_ids() if r["Id"] == 3)
    table.delete_row(rid)
    ours = db.to_bytes()
    engine = theirs2.read_bytes()
    different = [
        n for n in range(1, len(ours) // 4096)
        if ours[n * 4096 : (n + 1) * 4096] != engine[n * 4096 : (n + 1) * 4096]
    ]
    assert not different, f"pages differ from the engine's after a delete: {different}"


def test_every_code_point_keys_as_the_engine_keys_it(tmp_path: Path) -> None:
    """One indexed row per BMP code point plus composition samples, as the
    generator uses; the encoder must match all of them except keys past the
    engine's 509-byte cap, which it refuses."""
    from pyopenvba.access._collation import MAX_KEY_LENGTH, encode_text_key
    from pyopenvba.access._index import leaf_entries
    from pyopenvba.access._rows import split_row

    target = tmp_path / "chars.accdb"
    shutil.copy(TEMPLATE, target)
    assert oracle("-Command", "build-collation", "-Path", str(target)) == "ok"
    db = AccessDatabase(target)
    table = db.table("Chars")
    index = table.index("IX_Ch")
    checked = capped = 0
    wrong: list[str] = []
    for entry in leaf_entries(db.store, index.real.root_page):
        raw = table.fetch_row(entry.page, entry.row)
        assert raw is not None
        text = table.decode(split_row(table.definition, raw))["Ch"]
        assert isinstance(text, str)
        expected = entry.key[1:]
        if len(expected) >= MAX_KEY_LENGTH:
            capped += 1
            with pytest.raises(AccessError):
                encode_text_key(text)
            continue
        checked += 1
        if encode_text_key(text) != expected:
            wrong.append(f"{text!r}: engine {expected.hex()} encoder {encode_text_key(text).hex()}")
    assert checked > 63000 and capped >= 1
    assert not wrong, f"{len(wrong)} keys differ:\n" + "\n".join(wrong[:10])


def test_every_column_type_reads_back_as_the_engine_shows_it(tmp_path: Path) -> None:
    target = tmp_path / "alltypes.accdb"
    shutil.copy(TEMPLATE, target)
    assert oracle("-Command", "build-alltypes", "-Path", str(target), "-Rows", str(ROWS)) == "ok"
    expected = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "AllTypes"))
    expected_wide = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "Wide"))
    assert isinstance(expected, str) and isinstance(expected_wide, str)

    db = AccessDatabase(target)
    table = db.table("AllTypes")
    assert table.row_count == ROWS
    assert table.column_names == [
        "Id", "Flag", "Tiny", "Small", "Big", "Cash", "Sgl", "Dbl", "Stamp",
        "Bin", "Txt", "Blob", "Story", "Uid", "Frac", "Huge",
    ]
    mismatches = _diff(expected, _engine_dump(db, "AllTypes"))
    assert not mismatches, f"{len(mismatches)} rows differ\n" + "\n".join(mismatches[:5])

    _check_indexes(db)

    # The wide table's definition spans pages.
    wide = db.table("Wide")
    assert len(wide.definition.pages) >= 2, wide.definition.pages
    assert wide.definition.definition_length > 4096
    assert len(wide.columns) == 151
    mismatches = _diff(expected_wide, _engine_dump(db, "Wide"))
    assert not mismatches, "\n".join(mismatches[:3])


SQL_GATE_SELECTS = (
    "SELECT Id, Flag, Tiny, Small, Big, Cash, Dbl, Txt FROM AllTypes WHERE Big > 0 AND Txt LIKE '*1*' ORDER BY Id",
    "SELECT Id, Txt FROM AllTypes WHERE Txt IS NULL OR Small < 0 ORDER BY Id DESC",
    "SELECT TOP 5 Id, Dbl FROM AllTypes ORDER BY Dbl DESC, Id",
    "SELECT Flag, Count(*) AS N, Sum(Big) AS SumBig, Avg(Dbl) AS AvgDbl, Min(Cash) AS MinCash, Max(Stamp) AS LastStamp FROM AllTypes GROUP BY Flag ORDER BY Flag",
    "SELECT DISTINCT Tiny FROM AllTypes ORDER BY Tiny",
    "SELECT Id, Len(Txt) AS L, UCase(Left(Txt, 3)) AS U, IIf(Flag, 'on', 'off') AS F, IIf(Txt IS NULL, '-', Txt) AS N, Big * 2 + 1 AS Calc, Cash + 1 AS CashPlus FROM AllTypes ORDER BY Id",
    "SELECT a.Id, b.Id AS Other, b.Txt FROM AllTypes AS a INNER JOIN AllTypes AS b ON a.Flag = b.Flag WHERE a.Id < b.Id AND a.Id <= 3 ORDER BY a.Id, b.Id",
    "SELECT Count(*) AS N, Sum(Cash) AS Total, Avg(Cash) AS Mean, Count(Txt) AS Named FROM AllTypes WHERE Stamp > #1/1/2000#",
    "SELECT Id FROM AllTypes WHERE Tiny BETWEEN 10 AND 200 AND Big NOT IN (1, 2, 3) ORDER BY Id",
    "SELECT Id, Year(Stamp) AS Y, Month(Stamp) AS M FROM AllTypes WHERE Stamp IS NOT NULL AND Flag = TRUE ORDER BY Id",
    "SELECT Tiny, Count(*) AS N FROM AllTypes GROUP BY Tiny HAVING Count(*) > 1 ORDER BY Tiny",
    "SELECT Id, Txt & '!' AS Cat, Txt + '!' AS Sum, Txt & Story AS Both FROM AllTypes WHERE Id IN (1, 7, 14) ORDER BY Id",
    # Subqueries, derived tables and unions, which DAO runs too.
    "SELECT Id FROM AllTypes WHERE Id IN (SELECT Id FROM AllTypes WHERE Tiny > 15) ORDER BY Id",
    "SELECT Id FROM AllTypes WHERE Id NOT IN (SELECT Id FROM AllTypes WHERE Flag = TRUE) ORDER BY Id",
    "SELECT a.Id FROM AllTypes AS a WHERE EXISTS (SELECT 1 FROM AllTypes AS b WHERE b.Tiny = a.Tiny AND b.Id <> a.Id) ORDER BY a.Id",
    "SELECT a.Id, (SELECT Count(*) FROM AllTypes AS b WHERE b.Flag = a.Flag) AS Same FROM AllTypes AS a WHERE a.Id <= 5 ORDER BY a.Id",
    "SELECT Id FROM AllTypes WHERE Big = (SELECT Max(Big) FROM AllTypes)",
    "SELECT t.Flag, t.N FROM (SELECT Flag, Count(*) AS N FROM AllTypes GROUP BY Flag) AS t ORDER BY t.Flag",
    "SELECT Tiny FROM AllTypes WHERE Id <= 3 UNION SELECT Tiny FROM AllTypes WHERE Id <= 6 ORDER BY Tiny",
    "SELECT Tiny FROM AllTypes WHERE Id <= 3 UNION ALL SELECT Tiny FROM AllTypes WHERE Id <= 3",
    # Crosstabs, run rather than saved.
    "TRANSFORM Sum(Big) AS Total SELECT Flag FROM AllTypes GROUP BY Flag PIVOT Tiny Mod 3",
    "TRANSFORM Count(*) AS N SELECT Flag, Count(Txt) AS Named FROM AllTypes WHERE Id <= 20 GROUP BY Flag ORDER BY Flag PIVOT Tiny Mod 4 IN (0, 1, 2, 3)",
)


def test_sql_executor_matches_the_engine(tmp_path: Path) -> None:
    """``AccessDatabase.execute`` answers SELECTs as DAO does on the same
    database (values and column names, row for row) and writes UPDATE and
    DELETE results byte for byte as DAO's Execute does."""
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    assert oracle("-Command", "build-alltypes", "-Path", str(theirs), "-Rows", "30") == "ok"
    script = tmp_path / "statement.sql"
    db = AccessDatabase(theirs)
    problems: list[str] = []
    for sql in SQL_GATE_SELECTS:
        script.write_text(sql, encoding="ascii")
        expected = json.loads(oracle("-Command", "query-dump", "-Path", str(theirs), "-SqlFile", str(script)))
        rows = db.execute(sql)
        assert isinstance(rows, list)
        mine = [chr(9).join(f"{name}={_format(value)}" for name, value in row.items()) for row in rows]
        if not expected and not mine:
            continue
        for line in _diff(expected, mine):
            problems.append(f"{sql}\n{line}")
    assert not problems, chr(10).join(problems)

    before = theirs.read_bytes()
    statements = (
        "UPDATE AllTypes SET Big = Big + 1, Txt = UCase(Txt) WHERE Id <= 10",
        "UPDATE AllTypes SET Txt = Txt & ' and more' WHERE Id BETWEEN 11 AND 15",
        "DELETE FROM AllTypes WHERE Small < 0 AND Id > 5",
    )
    script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
    assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"
    db = AccessDatabase(before)
    assert [db.execute(sql) for sql in statements] == [10, 5, 21]
    ours, engine = db.to_bytes(), theirs.read_bytes()
    differing = _differing_pages(ours, engine)
    assert not differing, f"UPDATE and DELETE pages differ from the engine's: {_describe_pages(ours, engine, differing)}"


def test_indexes_built_over_rows_count_them_as_the_engine_does(tmp_path: Path) -> None:
    """An index created over existing rows records how many rows it holds
    (nulls left out when it ignores them) beside its distinct-key count;
    a filtered delete takes one off per row and caps the distinct count
    at what is left; a foreign-key index built over rows does the same.
    Byte for byte against DAO's CREATE INDEX, ADD CONSTRAINT and a SQL
    DELETE on the same tables."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        return AccessDatabase(ours)

    def stamps(name: str) -> dict[str, object]:
        entry = _catalog_entry(theirs, name)
        return {"created": entry.date_create_serial, "updated": entry.date_update_serial}

    def counters(db: AccessDatabase, table: str) -> list[tuple[int, int]]:
        return [(real.row_count, real.entry_count) for real in db.table(table).definition.real_indexes]

    key = [IndexSpec("PK", ("Id",), primary=True)]
    rows = [(1, "a"), (1, "b"), (2, None), (None, "c"), (3, "d"), (None, "e")]
    engine_runs(
        "CREATE TABLE Counted (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG, T TEXT(50))",
        *(f"INSERT INTO Counted (N, T) VALUES ({'NULL' if n is None else n}, {'NULL' if t is None else repr(t)})" for n, t in rows),
    )
    db = AccessDatabase(TEMPLATE)
    counted = db.create_table("Counted", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=50, compressed=False)], key, **stamps("Counted"))
    for n, t in rows:
        counted.insert_row({"N": n, "T": t})
    db = same_then_reopen("Counted built", db)

    engine_runs("CREATE INDEX IX_N ON Counted (N)", "CREATE UNIQUE INDEX IX_T ON Counted (T) WITH IGNORE NULL")
    updated = stamps("Counted")["updated"]
    db.create_index("Counted", IndexSpec("IX_N", ("N",)), updated=updated)
    db.create_index("Counted", IndexSpec("IX_T", ("T",), unique=True, ignore_nulls=True), updated=updated)
    db = same_then_reopen("indexes over six rows", db)
    # The primary key predates the rows and counts nothing; IX_N holds all
    # six (two of them null, one distinct key); IX_T skips its null.
    assert counters(db, "Counted") == [(0, 6), (6, 4), (5, 5)]

    engine_runs("DELETE FROM Counted WHERE N = 1")
    assert db.execute("DELETE FROM Counted WHERE N = 1") == 2
    db = same_then_reopen("two rows deleted", db)
    assert counters(db, "Counted") == [(0, 6), (4, 4), (3, 3)]

    # An update that writes an indexed column costs that index one counted
    # row, even when the value does not change; a row whose key is null in
    # an ignore-nulls index costs it nothing.
    engine_runs("UPDATE Counted SET N = N + 100 WHERE Id = 5")
    assert db.execute("UPDATE Counted SET N = N + 100 WHERE Id = 5") == 1
    db = same_then_reopen("indexed column updated", db)
    assert counters(db, "Counted") == [(0, 6), (3, 3), (3, 3)]

    engine_runs("UPDATE Counted SET T = T WHERE Id = 6")
    assert db.execute("UPDATE Counted SET T = T WHERE Id = 6") == 1
    db = same_then_reopen("indexed column rewritten with its own value", db)
    assert counters(db, "Counted") == [(0, 6), (3, 3), (2, 2)]

    engine_runs("UPDATE Counted SET T = 'zz' WHERE Id = 3")
    assert db.execute("UPDATE Counted SET T = 'zz' WHERE Id = 3") == 1
    db = same_then_reopen("null key given a value", db)
    assert counters(db, "Counted") == [(0, 6), (3, 3), (2, 2)]

    # A plain index counts its null-keyed rows, so writing or deleting one
    # costs it a row like any other.
    engine_runs("UPDATE Counted SET N = 7 WHERE Id = 4")
    assert db.execute("UPDATE Counted SET N = 7 WHERE Id = 4") == 1
    db = same_then_reopen("null key of a plain index written", db)
    assert counters(db, "Counted") == [(0, 6), (2, 2), (2, 2)]

    engine_runs("DELETE FROM Counted WHERE Id = 6")
    assert db.execute("DELETE FROM Counted WHERE Id = 6") == 1
    db = same_then_reopen("row with a null plain key deleted", db)
    assert counters(db, "Counted") == [(0, 6), (1, 1), (1, 1)]

    engine_runs(
        "CREATE TABLE Kids (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, CountedId LONG)",
        "INSERT INTO Kids (CountedId) VALUES (3)",
        "INSERT INTO Kids (CountedId) VALUES (5)",
        "INSERT INTO Kids (CountedId) VALUES (NULL)",
        "INSERT INTO Kids (CountedId) VALUES (3)",
    )
    kids = db.create_table("Kids", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("CountedId", "Long")], key, **stamps("Kids"))
    for value in (3, 5, None, 3):
        kids.insert_row({"CountedId": value})
    db = same_then_reopen("Kids built", db)
    engine_runs("ALTER TABLE Kids ADD CONSTRAINT FK_Kids_Counted FOREIGN KEY (CountedId) REFERENCES Counted (Id)")
    db.create_relationship(
        "FK_Kids_Counted", "Kids", ("CountedId",), "Counted", ("Id",),
        created=_catalog_entry(theirs, "FK_Kids_Counted").date_create_serial,
        table_updated=stamps("Kids")["updated"],
        referenced_updated=stamps("Counted")["updated"],
    )
    db = same_then_reopen("constraint over four rows", db)
    assert counters(db, "Kids")[-1][0] == 4
    for name in ("Counted", "Kids"):
        check_indexes(db.table(name))


def test_dropping_and_building_indexes_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """An index built over four hundred rows and then dropped, byte for
    byte against DAO.  The build exercises a B-tree that outgrows one
    page: the engine fills a leaf with uncompressed entries, compresses it
    with their shared prefix when it overflows, and closes it as soon as a
    key arrives without that prefix, so the leaves break where the prefix
    changes rather than when the page is full.  DROP INDEX then releases
    every page the index held with their bytes left alone, deletes its
    usage-map row and rewrites the definition without its records."""
    from test_access_write import check_indexes

    from pyopenvba.access import ColumnSpec, IndexSpec

    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    rows = 400

    def engine_runs(*statements: str) -> None:
        script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        return AccessDatabase(ours)

    def updated(name: str) -> object:
        return _catalog_entry(theirs, name).date_update_serial

    engine_runs(
        "CREATE TABLE Big (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG, T TEXT(60))",
        *(f"INSERT INTO Big (N, T) VALUES ({i * 7 % 991}, 'value number {i} padded out')" for i in range(1, rows + 1)),
    )
    entry = _catalog_entry(theirs, "Big")
    db = AccessDatabase(TEMPLATE)
    big = db.create_table(
        "Big",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=60, compressed=False)],
        [IndexSpec("PK", ("Id",), primary=True)],
        created=entry.date_create_serial,
        updated=entry.date_update_serial,
    )
    for i in range(1, rows + 1):
        big.insert_row({"N": i * 7 % 991, "T": f"value number {i} padded out"})
    db = same_then_reopen("four hundred rows", db)

    engine_runs("CREATE INDEX IX_T ON Big (T)", "CREATE INDEX IX_N ON Big (N)")
    when = updated("Big")
    db.create_index("Big", IndexSpec("IX_T", ("T",)), updated=when)
    db.create_index("Big", IndexSpec("IX_N", ("N",)), updated=when)
    db = same_then_reopen("a four-leaf text index and a one-page one", db)
    leaves = [len(page.entries) for page in leaf_pages(db.store, db.table("Big").definition.real_indexes[1].root_page)]
    assert leaves == [111, 111, 111, 67], leaves

    engine_runs("DROP INDEX IX_T ON Big")
    db.drop_index("Big", "IX_T", updated=updated("Big"))
    db = same_then_reopen("the four-leaf index dropped", db)
    assert sorted(i.name for i in db.table("Big").indexes) == ["IX_N", "PK"]

    # A page freed by a drop comes back only in a later session, so the
    # index made right after it takes one freed earlier instead.
    engine_runs("DROP INDEX IX_N ON Big", "CREATE INDEX IX_N2 ON Big (N)")
    when = updated("Big")
    db.drop_index("Big", "IX_N", updated=when)
    db.create_index("Big", IndexSpec("IX_N2", ("N",)), updated=when)
    db = same_then_reopen("dropped and rebuilt in one session", db)
    assert sorted(i.name for i in db.table("Big").indexes) == ["IX_N2", "PK"]
    check_indexes(db.table("Big"))


DDL_STEPS = (
    ("Parent", "CREATE TABLE Parent (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Name TEXT(50), Remark MEMO)"),
    ("Child", "CREATE TABLE Child (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, ParentId LONG, Amount CURRENCY, Flag BIT)"),
    ("Parent", "CREATE INDEX IX_Name ON Parent (Name)"),
    ("Child", "CREATE UNIQUE INDEX IX_Amount ON Child (Amount) WITH IGNORE NULL"),
    ("Child", "CREATE INDEX IX_Two ON Child (ParentId, Amount DESC)"),
    ("Parent", "ALTER TABLE Parent ADD COLUMN City TEXT(30)"),
    ("Parent", "ALTER TABLE Parent ADD COLUMN Score DOUBLE"),
    ("Parent", "ALTER TABLE Parent ALTER COLUMN Score SINGLE"),
    ("Parent", "ALTER TABLE Parent DROP COLUMN City"),
    ("Child", "DROP INDEX IX_Two ON Child"),
    ("Types", "CREATE TABLE Types (A BYTE, B SHORT, C INTEGER, D SINGLE, E DOUBLE, F CURRENCY, G DATETIME, H BIT, I GUID, J BIGINT, K MEMO, L LONGBINARY, M BINARY(20), N VARCHAR(40))"),
    (None, "DROP TABLE Types"),
)


def test_ddl_through_execute_matches_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """Every DDL statement ``db.execute`` runs leaves the bytes DAO's
    Execute leaves for the same statement: CREATE TABLE with a named
    primary key and every column type, CREATE INDEX (plain, unique
    ignore-nulls, two-column descending), ALTER TABLE ADD / ALTER / DROP
    COLUMN, ADD and DROP CONSTRAINT, DROP INDEX and DROP TABLE.  Each
    statement is its own DAO session, so the database is reopened between
    them and the engine's own timestamps are handed to the writer."""
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    db = AccessDatabase(TEMPLATE)

    def engine_runs(sql: str) -> None:
        script.write_text(sql + chr(10), encoding="ascii")
        assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script)) == "ok"

    def same_then_reopen(step: str, db: AccessDatabase) -> AccessDatabase:
        ours, engine = db.to_bytes(), theirs.read_bytes()
        assert not (d := _differing_pages(ours, engine)), f"{step}: pages differ from the engine's: {_describe_pages(ours, engine, d)}"
        return AccessDatabase(ours)

    for name, sql in DDL_STEPS:
        engine_runs(sql)
        stamps: dict[str, object] = {}
        if name is not None:
            entry = _catalog_entry(theirs, name)
            stamps = {"created": entry.date_create_serial, "updated": entry.date_update_serial}
        assert db.execute(sql, **stamps) == 0  # pyright: ignore[reportArgumentType]
        db = same_then_reopen(sql, db)

    # A foreign key names two tables, and the engine stamps both.
    engine_runs("ALTER TABLE Child ADD CONSTRAINT FK_Child_Parent FOREIGN KEY (ParentId) REFERENCES Parent (Id)")
    child, parent = _catalog_entry(theirs, "Child"), _catalog_entry(theirs, "Parent")
    db.execute(
        "ALTER TABLE Child ADD CONSTRAINT FK_Child_Parent FOREIGN KEY (ParentId) REFERENCES Parent (Id)",
        created=_catalog_entry(theirs, "FK_Child_Parent").date_create_serial,
        updated=child.date_update_serial,
        referenced_updated=parent.date_update_serial,
    )
    db = same_then_reopen("ADD CONSTRAINT", db)
    assert [r.name for r in db.relationships()][-1] == "FK_Child_Parent"

    engine_runs("ALTER TABLE Child DROP CONSTRAINT FK_Child_Parent")
    child, parent = _catalog_entry(theirs, "Child"), _catalog_entry(theirs, "Parent")
    db.execute(
        "ALTER TABLE Child DROP CONSTRAINT FK_Child_Parent",
        updated=child.date_update_serial,
        referenced_updated=parent.date_update_serial,
    )
    db = same_then_reopen("DROP CONSTRAINT", db)

    # The rows a SQL INSERT writes into what the SQL DDL built.
    engine_runs("INSERT INTO Child (ParentId, Amount, Flag) VALUES (1, 9.99, TRUE)")
    assert db.execute("INSERT INTO Child (ParentId, Amount, Flag) VALUES (1, 9.99, TRUE)") == 1
    db = same_then_reopen("INSERT after the DDL", db)
    for table in ("Parent", "Child"):
        check_indexes(db.table(table))


def test_a_transaction_writes_what_the_engine_writes(tmp_path: Path) -> None:
    """DAO wrapping its statements in BeginTrans/CommitTrans leaves the
    same bytes as running them plainly, and so does ours: a transaction
    is a way of undoing work, not a different way of writing it.  A
    rolled-back block leaves the database exactly as it was."""
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    script = tmp_path / "step.sql"
    statements = (
        "CREATE TABLE T (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG, T TEXT(40))",
        *(f"INSERT INTO T (N, T) VALUES ({i}, 'row {i}')" for i in range(1, 21)),
        "UPDATE T SET N = N + 100 WHERE Id <= 5",
        "DELETE FROM T WHERE Id > 15",
    )
    script.write_text(chr(10).join(statements) + chr(10), encoding="ascii")
    assert oracle("-Command", "sql-file", "-Path", str(theirs), "-SqlFile", str(script), "-Transaction") == "ok"

    db = AccessDatabase(TEMPLATE)
    entry_stamps: dict[str, object] = {}
    with db.transaction():
        entry = _catalog_entry(theirs, "T")
        entry_stamps = {"created": entry.date_create_serial, "updated": entry.date_update_serial}
        for sql in statements:
            db.execute(sql, **entry_stamps) if sql.startswith("CREATE") else db.execute(sql)  # pyright: ignore[reportArgumentType]
    ours, engine = db.to_bytes(), theirs.read_bytes()
    assert not (d := _differing_pages(ours, engine)), f"pages differ from the engine's: {_describe_pages(ours, engine, d)}"

    # What a rollback undoes: the same statements, then put back.
    before = db.to_bytes()
    try:
        with db.transaction():
            db.execute("INSERT INTO T (N, T) VALUES (99, 'gone')")
            db.execute("CREATE TABLE Gone (A LONG)")
            db.execute("DELETE FROM T")
            raise RuntimeError("roll it back")
    except RuntimeError:
        pass
    assert db.to_bytes() == before
    assert db.table_names() == ["T"] and db.table("T").row_count == 15
    check_indexes(db.table("T"))


def test_a_usage_map_past_its_page_becomes_a_reference_map(tmp_path: Path) -> None:
    """A hundred and thirty megabytes of long values, byte for byte.
    Past about thirty thousand pages a column's owned-pages map can no
    longer grow inside its row: the engine turns it into the reference
    form -- a row of seventeen chunk pointers, each naming a page whose
    bytes are one 32 736-page bitmap -- and this does the same, on the
    same page, at the same row.  Only the per-session stamps of the
    chained values differ, as always."""
    from pyopenvba.access import ColumnSpec, IndexSpec
    from pyopenvba.access._pages import GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW, read_usage_map, read_usage_map_ref

    rows, size = 130, 1024 * 1024
    theirs = tmp_path / "theirs.accdb"
    shutil.copy(TEMPLATE, theirs)
    assert oracle("-Command", "fill-big", "-Path", str(theirs), "-Rows", str(rows), "-Size", str(size)) == "ok"
    entry = _catalog_entry(theirs, "Bulk")

    value = bytes(((k * 7 + 11) % 256) for k in range(size))
    db = AccessDatabase(TEMPLATE)
    table = db.create_table(
        "Bulk",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("B", "OLE")],
        [IndexSpec("PK", ("Id",), primary=True)],
        created=entry.date_create_serial,
        updated=entry.date_update_serial,
    )
    for _ in range(rows):
        table.insert_row({"B": value})

    ours, engine = bytearray(db.to_bytes()), bytearray(theirs.read_bytes())
    assert len(ours) == len(engine), f"{len(ours) // 4096} pages against the engine's {len(engine) // 4096}"
    _mask_lval_stamps(ours, (size,))
    _mask_lval_stamps(engine, (size,))
    differing = _differing_pages(bytes(ours), bytes(engine))
    assert not differing, f"pages differ from the engine's: {_describe_pages(bytes(ours), bytes(engine), differing)}"

    # The map that grew out of its row, in both files.
    reopened = AccessDatabase(bytes(db.to_bytes()))
    definition = reopened.table("Bulk").definition
    owned = next(iter(definition.column_usage_maps.values()))[0]
    umap = read_usage_map_ref(reopened.store, owned)
    assert umap.kind == 1, "the column's owned-pages map should have become a reference map"
    assert len([p for p in umap.reference_pages if p]) == 2  # past one bitmap page's reach
    assert len(umap.pages()) == rows * (size // 4072 + 1)
    # The global free map runs out of row too, and gains chunks of its own.
    global_map = read_usage_map(reopened.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    assert global_map.kind == 1
    assert [p for p in global_map.reference_pages if p]
    assert reopened.table("Bulk").row_count == rows
    assert len(next(reopened.table("Bulk").rows())["B"]) == size  # pyright: ignore[reportArgumentType]
