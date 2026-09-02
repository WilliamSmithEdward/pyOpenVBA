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
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, CatalogEntry
from pyopenvba.access_read import AccessError

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
    entry = next(e for e in AccessDatabase(theirs).catalog() if e.name == "Simple")
    assert isinstance(entry.date_create, dt.datetime)
    db = AccessDatabase(TEMPLATE)
    db.create_table(
        "Simple",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=50, compressed=False)],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
        created=entry.date_create,
    )
    assert not (d := differing(db.to_bytes(), theirs.read_bytes())), f"create: pages differ from the engine's: {d}"

    # CREATE INDEX: the engine also re-stamps the catalog row, which moves
    # it when its page cannot hold a fresh copy.
    assert oracle("-Command", "index-simple", "-Path", str(theirs)) == "ok"
    entry = next(e for e in AccessDatabase(theirs).catalog() if e.name == "Simple")
    assert isinstance(entry.date_update, dt.datetime)
    db.create_index("Simple", IndexSpec("IX_N", ("N",)), updated=entry.date_update)
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
