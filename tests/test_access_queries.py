"""Saved queries: MSysQueries rows from Jet SQL and back, and the catalog
object DAO's CreateQueryDef makes.  The row encodings were measured on a
plain select, a joined DISTINCT TOP GROUP BY HAVING ORDER BY DESC query, a
parameter query and a DELETE."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec, QueryRow, SavedQuery
from pyopenvba.access._queries import rows_from_sql
from pyopenvba.access.database import OBJECT_QUERY, QUERY_ACTION_FLAG
from pyopenvba.access_read import AccessError
from test_access_write import check_indexes

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"

PLAIN = "SELECT Parent.Id, Parent.Name FROM Parent WHERE Parent.Id > 1 ORDER BY Parent.Name"
JOINED = (
    "SELECT DISTINCT TOP 5 Child.Id, Child.ParentId AS P FROM Child INNER JOIN Parent ON Child.ParentId = Parent.Id "
    "WHERE Child.Id > 2 GROUP BY Child.Id, Child.ParentId HAVING Count(*) > 0 ORDER BY Child.Id DESC"
)
PARAMETERS = "PARAMETERS [Which] Long; SELECT * FROM Parent WHERE Id = [Which]"
DELETE = "DELETE FROM Child WHERE Id < 0"


def test_rows_match_what_dao_wrote() -> None:
    assert rows_from_sql(PLAIN) == [
        QueryRow(0, 1, flag=0),
        QueryRow(255, 1),
        QueryRow(6, 1, expression="Parent.Id", flag=0),
        QueryRow(6, 2, expression="Parent.Name", flag=0),
        QueryRow(5, 1, name1="Parent"),
        QueryRow(8, 1, expression="Parent.Id > 1"),
        QueryRow(11, 1, expression="Parent.Name"),
    ]
    assert rows_from_sql(JOINED) == [
        QueryRow(0, 1, flag=0),
        QueryRow(255, 1),
        QueryRow(6, 1, expression="Child.Id", flag=0),
        QueryRow(6, 2, name1="P", expression="Child.ParentId", flag=0),
        QueryRow(7, 1, name1="Child", name2="Parent", expression="Child.ParentId = Parent.Id", flag=1),
        QueryRow(5, 1, name1="Child"),
        QueryRow(5, 2, name1="Parent"),
        QueryRow(8, 1, expression="Child.Id > 2"),
        QueryRow(9, 1, expression="Child.Id", flag=0),
        QueryRow(9, 2, expression="Child.ParentId", flag=0),
        QueryRow(10, 1, expression="Count(*) > 0"),
        QueryRow(11, 1, name1="d", expression="Child.Id"),
        QueryRow(3, 1, name1="5", flag=0x12),
    ]
    assert rows_from_sql(PARAMETERS) == [
        QueryRow(0, 1, flag=0),
        QueryRow(255, 1),
        QueryRow(2, 1, name1="[Which]", flag=4),
        QueryRow(5, 1, name1="Parent"),
        QueryRow(8, 1, expression="Id = [Which]"),
        QueryRow(3, 1, flag=1),
    ]
    assert rows_from_sql(DELETE) == [
        QueryRow(0, 1, flag=0),
        QueryRow(255, 1),
        QueryRow(1, 1, flag=5),
        QueryRow(5, 1, name1="Child"),
        QueryRow(8, 1, expression="Id < 0"),
    ]


@pytest.mark.parametrize("sql", [PLAIN, JOINED, PARAMETERS, DELETE])
def test_sql_round_trips_through_the_rows(sql: str) -> None:
    assert SavedQuery("Q", rows_from_sql(sql)).sql == sql


def test_unsupported_statements_are_refused() -> None:
    for sql in ("UPDATE Parent SET Name = 'x'", "SELECT Id", "Parent", "INSERT INTO Parent (Name) VALUES ('a')"):
        with pytest.raises(AccessError):
            rows_from_sql(sql)


def test_create_query_writes_the_catalog_object_and_rows() -> None:
    db = AccessDatabase(TEMPLATE)
    key = [IndexSpec("PK", ("Id",), primary=True)]
    db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50)], key)
    db.create_table("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key)
    before = len(db.queries())
    saved = db.create_query("ParentsAbove1", PLAIN)
    action = db.create_query("Purge", DELETE)
    assert saved.sql == PLAIN and saved.type == 1 and action.type == 5
    assert [q.name for q in db.queries()][before:] == ["ParentsAbove1", "Purge"]
    assert db.query("parentsabove1").sql == PLAIN and db.query("Purge").sql == DELETE
    entry = next(e for e in db.catalog() if e.name == "ParentsAbove1")
    purge = next(e for e in db.catalog() if e.name == "Purge")
    assert entry.type == OBJECT_QUERY and entry.id < 0 and entry.parent_id == db._container("Tables").id  # pyright: ignore[reportPrivateUsage]
    assert entry.flags == 0 and purge.flags == QUERY_ACTION_FLAG and purge.id == entry.id + 1
    props = next(r for r in db.table("MSysObjects").rows() if r["Name"] == "ParentsAbove1")["LvProp"]
    from pyopenvba.access._props import parse_property_blob

    assert isinstance(props, bytes) and parse_property_blob(props).decoded() == {"ODBCTimeout": 60, "MaxRecords": 0}
    assert len([r for r in db.table("MSysACEs").rows() if r["ObjectId"] == entry.id]) == 3
    with pytest.raises(AccessError):
        db.create_query("Parent", PLAIN)
    with pytest.raises(AccessError):
        db.query("Nope")
    for name in ("MSysObjects", "MSysQueries", "MSysACEs"):
        check_indexes(db.table(name))
    again = AccessDatabase(db.to_bytes())
    assert again.query("ParentsAbove1").rows == saved.rows
