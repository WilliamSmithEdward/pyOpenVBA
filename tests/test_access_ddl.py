"""Jet DDL through ``AccessDatabase.execute``: CREATE, DROP and ALTER."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._ddl import DDL_TYPES, column_spec
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def _fresh() -> AccessDatabase:
    return AccessDatabase(TEMPLATE)


@pytest.mark.parametrize(
    ("word", "expected"),
    [
        ("BIT", "Boolean"), ("YESNO", "Boolean"), ("LOGICAL1", "Boolean"),
        ("BYTE", "Byte"), ("INTEGER1", "Byte"),
        ("SHORT", "Integer"), ("SMALLINT", "Integer"), ("INTEGER2", "Integer"),
        # The Jet trap: INTEGER is four bytes here, not two.
        ("INTEGER", "Long"), ("INT", "Long"), ("LONG", "Long"), ("INTEGER4", "Long"),
        ("COUNTER", "Long"), ("AUTOINCREMENT", "Long"),
        ("CURRENCY", "Currency"), ("MONEY", "Currency"),
        ("SINGLE", "Single"), ("REAL", "Single"), ("IEEESINGLE", "Single"),
        ("DOUBLE", "Double"), ("FLOAT", "Double"), ("NUMBER", "Double"),
        ("DATETIME", "DateTime"), ("TIMESTAMP", "DateTime"),
        ("TEXT", "Text"), ("VARCHAR", "Text"), ("ALPHANUMERIC", "Text"),
        ("MEMO", "Memo"), ("LONGTEXT", "Memo"), ("NOTE", "Memo"),
        ("LONGBINARY", "OLE"), ("OLEOBJECT", "OLE"), ("GENERAL", "OLE"),
        ("BINARY", "Binary"), ("GUID", "GUID"), ("BIGINT", "BigInt"),
    ],
)
def test_every_type_word_makes_the_column_the_engine_makes(word: str, expected: str) -> None:
    db = _fresh()
    size = "(20)" if word in ("TEXT", "VARCHAR", "ALPHANUMERIC", "BINARY") else ""
    db.execute(f"CREATE TABLE T (C {word}{size})")
    column = db.table("T").columns[0]
    assert column.type_name == expected
    assert column.auto_number is (word in ("COUNTER", "AUTOINCREMENT"))
    assert DDL_TYPES[word.lower()] == column.type_code


def test_create_table_with_keys_and_a_foreign_key() -> None:
    db = _fresh()
    db.execute("CREATE TABLE Parent (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Name TEXT(50) CONSTRAINT U1 UNIQUE)")
    db.execute("CREATE TABLE Child (Id AUTOINCREMENT CONSTRAINT PK2 PRIMARY KEY, ParentId LONG, CONSTRAINT FK FOREIGN KEY (ParentId) REFERENCES Parent (Id))")
    parent = db.table("Parent")
    # ``.rC`` is the entry the relationship put on the parent's key.
    assert [i.name for i in parent.indexes] == [".rC", "PK", "U1"]
    assert parent.primary_key is not None and parent.primary_key.name == "PK"
    assert [r.name for r in db.relationships()][-1] == "FK"
    assert db.relationships()[-1].table == "Child"
    assert db.relationships()[-1].referenced_table == "Parent"


def test_not_null_is_taken_and_changes_nothing() -> None:
    # Measured: the engine writes a NOT NULL column's header exactly as it
    # writes a nullable one's.
    db = _fresh()
    db.execute("CREATE TABLE N (A LONG NOT NULL, B TEXT(20) NOT NULL, C LONG)")
    flags = [c.flags for c in db.table("N").columns]
    assert flags[0] == flags[2]
    assert all(c.nullable for c in db.table("N").columns)


def test_index_statements() -> None:
    db = _fresh()
    db.execute("CREATE TABLE T (Id LONG, N LONG, T TEXT(20))")
    db.execute("CREATE INDEX IX_N ON T (N)")
    db.execute("CREATE UNIQUE INDEX IX_T ON T (T) WITH IGNORE NULL")
    db.execute("CREATE INDEX IX_Two ON T (N, T DESC)")
    table = db.table("T")
    assert [i.name for i in table.indexes] == ["IX_N", "IX_T", "IX_Two"]
    assert table.index("IX_T").real.unique
    assert [asc for _c, asc in table.index("IX_Two").columns] == [True, False]
    db.execute("DROP INDEX IX_Two ON T")
    assert [i.name for i in db.table("T").indexes] == ["IX_N", "IX_T"]


def test_alter_table_statements() -> None:
    db = _fresh()
    db.execute("CREATE TABLE T (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, N LONG)")
    db.execute("INSERT INTO T (N) VALUES (7)")
    db.execute("ALTER TABLE T ADD COLUMN City TEXT(30)")
    db.execute("ALTER TABLE T ADD COLUMN Score DOUBLE")
    db.execute("ALTER TABLE T ALTER COLUMN Score SINGLE")
    db.execute("ALTER TABLE T DROP COLUMN City")
    assert [c.name for c in db.table("T").columns] == ["Id", "N", "Score"]
    assert db.table("T").definition.column("Score").type_name == "Single"
    assert db.execute("SELECT N FROM T") == [{"N": 7}]


def test_drop_table_and_reopen(tmp_path: Path) -> None:
    db = _fresh()
    db.execute("CREATE TABLE Keep (A LONG)")
    db.execute("CREATE TABLE Gone (A LONG, B TEXT(10))")
    db.execute("DROP TABLE Gone")
    db.save(tmp_path / "ddl.accdb")
    again = AccessDatabase(tmp_path / "ddl.accdb")
    assert again.table_names() == ["Keep"]


def test_a_bigint_column_brings_the_version_properties() -> None:
    db = _fresh()
    db.execute("CREATE TABLE Big (A LONG, B BIGINT)")
    db.execute("CREATE TABLE Small (A LONG)")
    assert db.table("Big").properties() == {
        "FCMinReadVer": "16.0.7124.1000",
        "FCMinWriteVer": "16.0.7124.1000",
        "FCMinDesignVer": "16.0.7124.1000",
    }
    assert db.table("Small").properties() == {}


def test_statements_the_engine_refuses_are_refused_here() -> None:
    db = _fresh()
    for sql, message in (
        ("CREATE TABLE X (A CHAR(10))", "fixed-width"),
        ("CREATE TABLE X (A DECIMAL(9,2))", "no DECIMAL"),
        ("CREATE TABLE X (A NUMERIC(9,2))", "no NUMERIC"),
        ("CREATE TABLE X (A TEXT(20) WITH COMPRESSION)", "WITH COMPRESSION"),
        ("CREATE TABLE X (A NOSUCH)", "unknown type"),
        ("DROP INDEX IX ON Nowhere", "no table named"),
        ("ALTER TABLE Nowhere ADD COLUMN A LONG", "no table named"),
    ):
        with pytest.raises(AccessError, match=message):
            db.execute(sql)


def test_column_spec_reads_inline_constraints() -> None:
    spec, constraints = column_spec("Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY")
    assert spec.name == "Id" and spec.autonumber and spec.type == "long"
    assert constraints == [("primary", "PK", ["Id"], None)]
    spec, constraints = column_spec("ParentId LONG CONSTRAINT FK REFERENCES Parent (Id)")
    assert constraints == [("foreign", "FK", ["ParentId"], ("Parent", ["Id"]))]
    spec, constraints = column_spec("Name TEXT(50) NOT NULL")
    assert spec.size == 50 and constraints == []
