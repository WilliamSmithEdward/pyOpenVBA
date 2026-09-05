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


def test_a_jet4_database_compacts_as_the_engine_compacts_it(tmp_path: Path) -> None:
    """The same copy in the Jet 4 format, from the bare database the engine
    makes with ``CreateDatabase(..., dbVersion40)``: two AutoNumber tables
    with a foreign key, a memo column, keyed and heap tables, deletes and a
    saved query.  The skeleton is the Jet 4 one; the recipe is the same."""
    source = tmp_path / "source.mdb"
    oracle("-Command", "create-blank-mdb", "-Path", str(source))
    setup = [
        "CREATE TABLE A (Id AUTOINCREMENT CONSTRAINT PKA PRIMARY KEY, Name TEXT(40), Notes MEMO, Amount CURRENCY)",
        "CREATE INDEX IxName ON A (Name)",
        "CREATE TABLE B (Id AUTOINCREMENT CONSTRAINT PKB PRIMARY KEY, AId LONG, Tag TEXT(20))",
        "CREATE TABLE Zeta (Id LONG CONSTRAINT PKZ PRIMARY KEY, V TEXT(10))",
        "CREATE TABLE alpha (Id LONG, V TEXT(10))",
    ]
    setup += [f"INSERT INTO Zeta VALUES ({i}, '{v}')" for i, v in ((5, "five"), (3, "three"), (9, "nine"), (1, "one"))]
    setup += [f"INSERT INTO alpha VALUES ({i}, '{v}')" for i, v in ((5, "five"), (3, "three"), (9, "nine"), (1, "one"))]
    setup += [f"INSERT INTO A (Name, Notes, Amount) VALUES ('name {i}', '{'memo text ' * (i % 40)}', {i}.25)" for i in range(1, 121)]
    setup += [f"INSERT INTO B (AId, Tag) VALUES ({i}, 'tag{i}')" for i in range(1, 41)]
    _run(source, setup, tmp_path)
    query = tmp_path / "query.sql"
    query.write_text("SELECT A.Id, A.Name, B.Tag FROM A INNER JOIN B ON A.Id = B.AId", encoding="utf-8")
    oracle("-Command", "create-query", "-Path", str(source), "-Table", "QJoin", "-SqlFile", str(query))
    for statement in (
        "DELETE FROM B WHERE Id > 20 OR AId Mod 3 = 0",
        "ALTER TABLE B ADD CONSTRAINT FK_B_A FOREIGN KEY (AId) REFERENCES A (Id)",
        "DELETE FROM A WHERE Id Mod 3 = 0",
    ):
        query.write_text(statement, encoding="utf-8")
        oracle("-Command", "run-sql", "-Path", str(source), "-SqlFile", str(query))
    assert_compaction_matches(source, tmp_path)


def test_access_objects_links_and_every_query_shape_compact_as_the_engine_compacts_them(tmp_path: Path) -> None:
    """A database Access itself has written to: two forms, a report, a
    macro and a module besides the template's, a linked table, a
    self-referencing relationship, a two-column cascading one, a foreign
    key riding an existing index, a Decimal and a BigInt column, table
    and column properties, a text key, an AutoNumber heap with explicit
    ids, and seven queries (select, group, parameter, union, action,
    crosstab and pass-through).  Needs desktop Access: set
    ``RUN_LIVE_ACCESS_VBA=1``."""
    if os.environ.get("RUN_LIVE_ACCESS_VBA") != "1":
        pytest.skip("needs desktop Access: set RUN_LIVE_ACCESS_VBA=1")
    harness = pytest.importorskip("pyvbaharness")
    from decimal import Decimal

    from pyopenvba.access import ColumnSpec

    source = tmp_path / "source.accdb"
    shutil.copy(TEMPLATE, source)
    setup = [
        "CREATE TABLE Types1 (Id AUTOINCREMENT CONSTRAINT PK1 PRIMARY KEY, B BYTE, S SHORT, L LONG, Sg SINGLE, D DOUBLE, C CURRENCY, Dt DATETIME, Yn YESNO, T TEXT(50), M MEMO, O LONGBINARY, G GUID, Bn BINARY(10))",
        "INSERT INTO Types1 (B, S, L, Sg, D, C, Dt, Yn, T, M) VALUES (1, 2, 3, 1.5, 2.25, 3.75, #2020-01-02 03:04:05#, TRUE, 'one', 'memo one')",
        "INSERT INTO Types1 (B, S, L, Sg, D, C, Dt, Yn, T, M) VALUES (200, -2, 30000, -1.5, 1E10, -3.75, #1999-12-31#, FALSE, 'two', NULL)",
        "INSERT INTO Types1 (L, T) VALUES (7, 'three')",
        "CREATE TABLE Multi (K1 LONG, K2 TEXT(10), V DOUBLE, CONSTRAINT PKM PRIMARY KEY (K1, K2))",
        "CREATE INDEX IxVDesc ON Multi (V DESC)",
        "CREATE UNIQUE INDEX IxK2 ON Multi (K2) WITH IGNORE NULL",
        "INSERT INTO Multi VALUES (2, 'b', 1.5)",
        "INSERT INTO Multi VALUES (1, 'z', 2.5)",
        "INSERT INTO Multi VALUES (1, 'a', 3.5)",
        "INSERT INTO Multi VALUES (3, 'c', NULL)",
        "CREATE TABLE Self (Id LONG CONSTRAINT PKS PRIMARY KEY, ParentId LONG, Label TEXT(10))",
        "INSERT INTO Self VALUES (1, NULL, 'root')",
        "INSERT INTO Self VALUES (2, 1, 'child')",
        "INSERT INTO Self VALUES (3, 1, 'child2')",
        "ALTER TABLE Self ADD CONSTRAINT FK_Self FOREIGN KEY (ParentId) REFERENCES Self (Id)",
        "CREATE TABLE Kid (Id LONG CONSTRAINT PKK PRIMARY KEY, K1 LONG, K2 TEXT(10), Remark TEXT(10))",
        "INSERT INTO Kid VALUES (1, 2, 'b', 'x')",
        "INSERT INTO Kid VALUES (2, 1, 'a', 'y')",
        "CREATE TABLE Shared (Id LONG CONSTRAINT PKSh PRIMARY KEY, SelfId LONG)",
        "CREATE INDEX IxSelfId ON Shared (SelfId)",
        "INSERT INTO Shared VALUES (1, 2)",
        "INSERT INTO Shared VALUES (2, 2)",
        "ALTER TABLE Shared ADD CONSTRAINT FK_Shared_Self FOREIGN KEY (SelfId) REFERENCES Self (Id)",
        "CREATE TABLE Described (Id LONG, Name TEXT(30) NOT NULL, Remark TEXT(30))",
        "INSERT INTO Described VALUES (1, 'n', 'x')",
        "CREATE TABLE TextKey (Code TEXT(5) CONSTRAINT PKT PRIMARY KEY, V LONG)",
        "INSERT INTO TextKey VALUES ('m', 1)",
        "INSERT INTO TextKey VALUES ('a', 2)",
        "INSERT INTO TextKey VALUES ('Z', 3)",
        "INSERT INTO TextKey VALUES ('b', 4)",
        "CREATE TABLE HeapAuto (Id AUTOINCREMENT, V TEXT(5))",
        "INSERT INTO HeapAuto (Id, V) VALUES (5, 'e')",
        "INSERT INTO HeapAuto (Id, V) VALUES (3, 'c')",
        "INSERT INTO HeapAuto (Id, V) VALUES (9, 'i')",
        "INSERT INTO HeapAuto (Id, V) VALUES (1, 'a')",
    ]
    _run(source, setup, tmp_path)
    # A Decimal and a BigInt column, which DAO's DDL cannot declare, and a
    # cascading two-column relationship, which its DDL will not accept.
    db = AccessDatabase(source)
    dec = db.create_table("Dec", [ColumnSpec("Id", "Long"), ColumnSpec("Dc", "Decimal", size=(18, 4)), ColumnSpec("Bi", "BigInt")])
    dec.insert_row({"Id": 1, "Dc": Decimal("12.3456"), "Bi": 2**40})
    db.create_relationship("FK_Kid_Multi", "Kid", ("K1", "K2"), "Multi", ("K1", "K2"), cascade_updates=True, cascade_deletes=True)
    db.save()
    oracle("-Command", "set-props", "-Path", str(source), "-Table", "Described")
    other = tmp_path / "other.accdb"
    shutil.copy(TEMPLATE, other)
    _run(other, ["CREATE TABLE Remote (Id LONG CONSTRAINT PKR PRIMARY KEY, Name TEXT(20))", "INSERT INTO Remote VALUES (1, 'far')"], tmp_path)
    link = tmp_path / "link.txt"
    link.write_text(";" + chr(9) + str(other) + chr(9) + "Remote", encoding="utf-8")
    oracle("-Command", "link-table", "-Path", str(source), "-Table", "LinkedRemote", "-SqlFile", str(link))
    query = tmp_path / "query.sql"
    queries = {
        "QWhere": "SELECT Types1.Id, Types1.T FROM Types1 WHERE Types1.L > 5 ORDER BY Types1.T DESC",
        "QGroup": "SELECT Multi.K1, Sum(Multi.V) AS Total FROM Multi GROUP BY Multi.K1 HAVING Sum(Multi.V) > 0",
        "QParam": "PARAMETERS pMin Long; SELECT * FROM Types1 WHERE Types1.L >= pMin",
        "QUnion": "SELECT Id FROM Self UNION SELECT Id FROM Kid",
        "QUpdate": "UPDATE Types1 SET Types1.L = 0 WHERE Types1.Id = 1",
        "QCross": "TRANSFORM Sum(Multi.V) AS S SELECT Multi.K1 FROM Multi GROUP BY Multi.K1 PIVOT Multi.K2",
    }
    for name, sql in queries.items():
        query.write_text(sql, encoding="utf-8")
        oracle("-Command", "create-query", "-Path", str(source), "-Table", name, "-SqlFile", str(query))
    query.write_text("SELECT 1 AS One", encoding="utf-8")
    oracle("-Command", "create-passthrough", "-Path", str(source), "-Table", "QPass", "-SqlFile", str(query))
    module_text = tmp_path / "amod.txt"
    module_text.write_text("Option Compare Database" + chr(10) + "Option Explicit" + chr(10) + chr(10) + "Public Sub Hi()" + chr(10) + "    Debug.Print 1" + chr(10) + "End Sub" + chr(10), encoding="utf-8")
    macro_text = tmp_path / "mac1.txt"
    macro_text.write_text("Version =196611" + chr(10) + "ColumnsShown =0" + chr(10) + "Begin" + chr(10) + '    Action ="Beep"' + chr(10) + "End" + chr(10), encoding="utf-8")
    with harness.AccessSession() as access:
        access.open_document(source, read_only=False)
        result = access.run_vba(_BUILD_OBJECTS, proc="Build", args=(str(module_text), str(macro_text)), timeout=300.0)
        assert result.outcome == "passed" and result.value == "ok", f"Access could not build the objects: {result.outcome} {result.value}"
    assert_compaction_matches(source, tmp_path)


def test_wide_tables_compact_as_the_engine_compacts_them(tmp_path: Path) -> None:
    """Two tables whose definitions run past one page: one copied before
    the system tables (its pages come round within 16 and the file grows
    from 64 to 72), one copied last with a memo column, a primary key and
    rows (rewritten after the rows and twice for the index, its pages
    coming back at 136 and 144).  Every stale continuation page the
    churn leaves behind has to hold the engine's bytes."""
    source = tmp_path / "source.accdb"
    shutil.copy(TEMPLATE, source)
    wide = ", ".join(f"C{i:03} TEXT(20)" for i in range(1, 161))
    setup = [
        f"CREATE TABLE AWide (Id LONG, {wide})",
        f"CREATE TABLE ZWide (Id LONG CONSTRAINT PKZ PRIMARY KEY, M MEMO, {wide})",
        "INSERT INTO ZWide (Id, M, C001) VALUES (2, 'second memo', 'two')",
        "INSERT INTO ZWide (Id, M, C001) VALUES (1, 'first memo', 'one')",
    ]
    _run(source, setup, tmp_path)
    assert_compaction_matches(source, tmp_path)


_BUILD_OBJECTS = """
Public Function Build(ByVal modPath As String, ByVal macPath As String) As String
    Dim f As Object, r As Object, c As Object, n As String
    Set f = CreateForm()
    Set c = CreateControl(f.Name, 109, 0, "", "", 500, 500, 2000, 300)
    c.Name = "txtA"
    n = f.Name
    DoCmd.Close acForm, n, acSaveYes
    DoCmd.Rename "Zform", acForm, n
    Set f = CreateForm()
    n = f.Name
    DoCmd.Close acForm, n, acSaveYes
    DoCmd.Rename "Aform", acForm, n
    Set r = CreateReport()
    n = r.Name
    DoCmd.Close acReport, n, acSaveYes
    DoCmd.Rename "Rpt1", acReport, n
    Application.LoadFromText acModule, "AMod", modPath
    Application.LoadFromText acMacro, "Mac1", macPath
    Build = "ok"
End Function
"""
