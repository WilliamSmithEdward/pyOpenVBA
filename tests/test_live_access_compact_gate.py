"""Compact and Repair against the engine's own, page for page.

DAO's ``CompactDatabase`` is run with the engine's clock frozen at the
source's creation date (see ``live_access_test/frozen_clock.py``), which
keeps the SIDs the engine keys to that date unchanged, and pyOpenVBA's
``compact_and_repair`` is given the same instant.  Every page but page 0
must then match, the per-session stamps on long-value chains excepted.

Opt-in: set ``RUN_LIVE_ACCESS=1`` on a Windows machine with Access and
pywin32 installed.
"""

from __future__ import annotations

import os
import shutil
import struct
import subprocess
import sys
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._rows import LongValueRef, decode_long_value_ref, split_row

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS") != "1" or sys.platform != "win32",
    reason="live ACE gate: set RUN_LIVE_ACCESS=1 on Windows with Access installed",
)

HERE = Path(__file__).parent
ORACLE = HERE / "live_access_test" / "dao_oracle.ps1"
FROZEN_CLOCK = HERE / "live_access_test" / "frozen_clock.py"
TEMPLATE = HERE.parents[0] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
ATTACHMENTS = HERE / "live_access_test" / "complex_columns.accdb"
_TIMEOUT = 900


def oracle(*args: str) -> str:
    done = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ORACLE), *args],
        capture_output=True,
        text=True,
        timeout=_TIMEOUT,
    )
    assert done.returncode == 0, f"DAO oracle failed: {done.stdout[-300:]} {done.stderr[-300:]}"
    return done.stdout.strip()


def frozen_compaction(source: Path, destination: Path) -> float:
    """The engine's compaction of ``source`` under a frozen clock; returns
    the frozen serial, which is the source's creation date."""
    pytest.importorskip("win32com")
    done = subprocess.run(
        [sys.executable, str(FROZEN_CLOCK), str(source), str(destination)],
        capture_output=True,
        text=True,
        timeout=_TIMEOUT,
    )
    assert done.returncode == 0, f"frozen compaction failed: {done.stdout[-300:]} {done.stderr[-500:]}"
    return float(done.stdout.strip())


def _lval_stamp(blob: bytes | bytearray) -> int | None:
    """The per-session stamp the engine wrote on its first long-value page."""
    for n in range(1, len(blob) // 4096):
        start = n * 4096
        if blob[start] == 1 and blob[start + 4 : start + 8] == b"LVAL":
            stamp = struct.unpack_from("<I", blob, start + 8)[0]
            if stamp:
                return stamp
    return None


def _chain_lengths(path: Path) -> set[int]:
    """The lengths of every chained long value, whose 12-byte definitions
    carry the one stamp a chain gets."""
    db = AccessDatabase(path)
    lengths: set[int] = set()
    for name in db.table_names(include_system=True):
        table = db.table(name)
        definition = table.definition
        for _page, _slot, data in table.raw_rows():
            parts = split_row(definition, data)
            for column in definition.columns:
                raw = parts.values.get(column.number)
                if column.is_long_value and raw and decode_long_value_ref(raw).kind == LongValueRef.KIND_CHAINED:
                    lengths.add(decode_long_value_ref(raw).length)
    return lengths


def _mask_stamps(blob: bytearray, chain_lengths: set[int]) -> None:
    """Zero the per-session stamp at offset 8 of every LVAL page and in the
    definition of every chained value -- the one thing about a chain the
    engine does not derive from the data."""
    for n in range(1, len(blob) // 4096):
        start = n * 4096
        page = bytes(blob[start : start + 4096])
        if page[0] == 1 and page[4:8] == b"LVAL":
            blob[start + 8 : start + 12] = b"\0\0\0\0"
        for length in chain_lengths:
            marker = length.to_bytes(3, "little") + b"\x00"
            pos = page.find(marker)
            while pos >= 0:
                blob[start + pos + 8 : start + pos + 12] = b"\0\0\0\0"
                pos = page.find(marker, pos + 1)


def assert_compaction_matches(source: Path, tmp_path: Path) -> None:
    engine_path = tmp_path / "engine.accdb"
    serial = frozen_compaction(source, engine_path)
    engine = bytearray(engine_path.read_bytes())
    ours = bytearray(AccessDatabase(source).compact_and_repair(clock=lambda: serial, lval_stamp=_lval_stamp(engine)).to_bytes())
    assert len(ours) == len(engine), f"ours has {len(ours) // 4096} pages, the engine's {len(engine) // 4096}"
    lengths = _chain_lengths(engine_path)
    _mask_stamps(ours, lengths)
    _mask_stamps(engine, lengths)
    differing: list[str] = []
    for n in range(1, len(ours) // 4096):
        a, b = ours[n * 4096 : (n + 1) * 4096], engine[n * 4096 : (n + 1) * 4096]
        if a != b:
            offsets = [i for i in range(4096) if a[i] != b[i]]
            lo = offsets[0]
            owner = int.from_bytes(b[4:8], "little") if b[0] in (1, 3, 4, 5) else 0
            differing.append(
                f"page {n} (type {b[0]:#04x}, owner {owner}): {len(offsets)} bytes from {lo}; "
                f"ours {a[lo:lo + 12].hex(' ')} / engine {b[lo:lo + 12].hex(' ')}"
            )
    assert not differing, f"{len(differing)} pages differ from the engine's: " + "; ".join(differing[:5])


def _run(path: Path, statements: list[str], tmp_path: Path) -> None:
    script = tmp_path / "statements.sql"
    script.write_text("\n".join(statements), encoding="utf-8")
    oracle("-Command", "sql-file", "-Path", str(path), "-SqlFile", str(script))


def test_deleted_rows_keys_memos_and_a_relationship_compact_as_the_engine_compacts_them(tmp_path: Path) -> None:
    """Two AutoNumber tables with a foreign key between them, a memo column
    whose values span inline, single-page and shared pages, heap and keyed
    tables with rows out of key order, a dropped table, deletes that leave
    dead rows and an AutoNumber past its largest value, and a saved join
    query.  The engine's copy renumbers every page, writes keyed rows in
    key order, resets the counters, rebuilds the permission rows and the
    relationship, and pyOpenVBA's copy lands on the same bytes."""
    source = tmp_path / "source.accdb"
    shutil.copy(TEMPLATE, source)
    setup = [
        "CREATE TABLE A (Id AUTOINCREMENT CONSTRAINT PKA PRIMARY KEY, Name TEXT(40), Notes MEMO, Amount CURRENCY)",
        "CREATE INDEX IxName ON A (Name)",
        "CREATE TABLE B (Id AUTOINCREMENT CONSTRAINT PKB PRIMARY KEY, AId LONG, Tag TEXT(20))",
        "CREATE TABLE C (Id LONG, Filler TEXT(200))",
        "CREATE TABLE D (Id LONG, Filler TEXT(200))",
        "CREATE TABLE Zeta (Id LONG CONSTRAINT PKZ PRIMARY KEY, V TEXT(10))",
        "CREATE TABLE alpha (Id LONG, V TEXT(10))",
    ]
    setup += [f"INSERT INTO Zeta VALUES ({i}, '{v}')" for i, v in ((5, "five"), (3, "three"), (9, "nine"), (1, "one"))]
    setup += [f"INSERT INTO alpha VALUES ({i}, '{v}')" for i, v in ((5, "five"), (3, "three"), (9, "nine"), (1, "one"))]
    setup += [f"INSERT INTO A (Name, Notes, Amount) VALUES ('name {i}', '{'memo text ' * (i % 40)}', {i}.25)" for i in range(1, 301)]
    setup += [f"INSERT INTO B (AId, Tag) VALUES ({i}, 'tag{i}')" for i in range(1, 101)]
    setup += [f"INSERT INTO C VALUES ({i}, '{'c' * 150}')" for i in range(1, 201)]
    setup += [f"INSERT INTO D VALUES ({i}, '{'d' * 150}')" for i in range(1, 201)]
    _run(source, setup, tmp_path)
    query = tmp_path / "query.sql"
    query.write_text("SELECT A.Id, A.Name, B.Tag FROM A INNER JOIN B ON A.Id = B.AId", encoding="utf-8")
    oracle("-Command", "create-query", "-Path", str(source), "-Table", "QJoin", "-SqlFile", str(query))
    for statement in (
        "DELETE FROM B WHERE Id > 50 OR AId Mod 3 = 0",
        "ALTER TABLE B ADD CONSTRAINT FK_B_A FOREIGN KEY (AId) REFERENCES A (Id)",
        "DELETE FROM A WHERE Id Mod 3 = 0",
        "DROP TABLE C",
        "DELETE FROM D WHERE Id <= 190",
        "INSERT INTO A (Name, Notes, Amount) VALUES ('late', NULL, 1)",
    ):
        query.write_text(statement, encoding="utf-8")
        oracle("-Command", "run-sql", "-Path", str(source), "-SqlFile", str(query))
    assert_compaction_matches(source, tmp_path)


def test_counters_indexes_dropped_columns_and_queries_compact_as_the_engine_compacts_them(tmp_path: Path) -> None:
    """An emptied AutoNumber table (its counter goes to zero), one with a
    step of five (kept, the counter at the largest value present), one
    with no index, a unique index that is not the key (rows keep their
    order), a table whose first row needs long-value pages before its
    data page, a dropped column (the copy renumbers the rest but keeps
    bytes 9-10 of each header), and two queries that sort before and
    after the tables."""
    source = tmp_path / "source.accdb"
    shutil.copy(TEMPLATE, source)
    setup = [
        "CREATE TABLE E1 (Id AUTOINCREMENT CONSTRAINT PK1 PRIMARY KEY, V TEXT(5))",
        "CREATE TABLE E2 (Id AUTOINCREMENT(100, 5) CONSTRAINT PK2 PRIMARY KEY, V TEXT(5))",
        "CREATE TABLE E3 (Id AUTOINCREMENT, V TEXT(5))",
        "CREATE TABLE P (Id LONG, V TEXT(20) NOT NULL, W TEXT(10), M MEMO)",
        "CREATE UNIQUE INDEX UxV ON P (V)",
        "CREATE TABLE H (Id LONG, M MEMO)",
        "CREATE TABLE O (A LONG, B LONG)",
        "ALTER TABLE O ADD COLUMN C TEXT(5)",
        "ALTER TABLE O DROP COLUMN B",
        "INSERT INTO O (A, C) VALUES (1, 'one')",
        "INSERT INTO O (A, C) VALUES (2, 'two')",
    ]
    setup += [f"INSERT INTO E1 (V) VALUES ('v{i}')" for i in range(10)] + ["DELETE FROM E1"]
    setup += [f"INSERT INTO E2 (V) VALUES ('v{i}')" for i in range(3)] + ["DELETE FROM E2 WHERE Id = 110"]
    setup += [f"INSERT INTO E3 (V) VALUES ('v{i}')" for i in range(5)] + ["DELETE FROM E3 WHERE Id = 5"]
    setup += [f"INSERT INTO P (Id, V) VALUES ({i}, '{v}')" for i, v in enumerate(["mike", "alpha", "zulu", "bravo", "kilo"], start=1)]
    setup += ["INSERT INTO H VALUES (1, '" + "h" * 6000 + "')"]
    setup += [f"INSERT INTO H VALUES ({i}, 'short {i}')" for i in range(2, 6)]
    setup += ["INSERT INTO H VALUES (6, '" + "k" * 9000 + "')"]
    _run(source, setup, tmp_path)
    query = tmp_path / "query.sql"
    for name, sql in (("Aq", "SELECT * FROM P"), ("Zq", "SELECT * FROM H")):
        query.write_text(sql, encoding="utf-8")
        oracle("-Command", "create-query", "-Path", str(source), "-Table", name, "-SqlFile", str(query))
    assert_compaction_matches(source, tmp_path)


def test_attachments_and_multi_valued_columns_compact_as_the_engine_compacts_them(tmp_path: Path) -> None:
    """The fixture Access wrote with an attachment and a multi-valued
    column: the flat tables sort before their table and are copied with
    their columns renumbered in header order, the MSysComplexColumns
    rows point at the new ids, and the complex counter comes out one
    past the source's."""
    source = tmp_path / "source.accdb"
    shutil.copy(ATTACHMENTS, source)
    assert_compaction_matches(source, tmp_path)
