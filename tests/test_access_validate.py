"""A table's own rules: DefaultValue, Required and the validation rules.

What each rule means was measured against the engine; see
:mod:`pyopenvba.access._validate`.  These check the writers keep to it
without an engine present.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._validate import default_value
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def _table(sql: str = "CREATE TABLE T (A LONG, B TEXT(20), C TEXT(30))") -> AccessDatabase:
    db = AccessDatabase(TEMPLATE)
    db.execute(sql)
    return db


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("5", 5),
        ("1+1", 2),
        ("3*2", 6),
        ("0", 0),
        ('"abc"', "abc"),
        ("'abc'", "abc"),
        # A name the expression cannot resolve is its own text, which is
        # how the engine reads an unquoted default.
        ("hello", "hello"),
        ("a & b", "ab"),
        ("Left('abcdef',3)", "abc"),
        ("=1+1", 2),
        ("", None),
    ],
)
def test_a_default_reads_as_the_engine_reads_it(text: str, expected: object) -> None:
    assert default_value(text) == expected


def test_a_default_may_call_a_date_function() -> None:
    assert default_value("=Date()") == dt.datetime.combine(dt.date.today(), dt.time())


def test_a_column_left_out_takes_its_default() -> None:
    db = _table()
    table = db.table("T")
    table.set_properties({"DefaultValue": "7"}, column="A")
    table.set_properties({"DefaultValue": "hello"}, column="C")
    db.table("T").insert_row({"B": "one"})
    db.execute("INSERT INTO T (B) VALUES ('two')")
    assert [(r["A"], r["B"], r["C"]) for r in db.table("T").rows()] == [
        (7, "one", "hello"),
        (7, "two", "hello"),
    ]


def test_a_column_that_is_given_keeps_what_it_is_given() -> None:
    db = _table()
    db.table("T").set_properties({"DefaultValue": "7"}, column="A")
    db.execute("INSERT INTO T (A, B) VALUES (3, 'x')")
    db.table("T").insert_row({"A": None, "B": "y"})
    assert [r["A"] for r in db.table("T").rows()] == [3, None]


def test_a_required_column_refuses_a_null() -> None:
    db = _table()
    db.table("T").set_properties({"Required": True}, column="B")
    for values in ({"A": 1}, {"A": 1, "B": None}):
        with pytest.raises(AccessError, match=r"You must enter a value in the 'T.B' field."):
            db.table("T").insert_row(values)
    row = db.table("T").insert_row({"A": 1, "B": "here"})
    with pytest.raises(AccessError, match=r"You must enter a value in the 'T.B' field."):
        db.table("T").update_row(row, {"B": None})
    assert db.table("T").row_count == 1


def test_a_column_rule_is_about_its_own_column() -> None:
    db = _table()
    db.table("T").set_properties({"ValidationRule": ">0"}, column="A")
    with pytest.raises(AccessError, match="prohibited by the validation rule"):
        db.execute("INSERT INTO T (A, B) VALUES (-1, 'x')")
    db.execute("INSERT INTO T (A, B) VALUES (3, 'x')")
    assert db.table("T").row_count == 1


def test_a_rule_lets_a_null_through() -> None:
    db = _table()
    db.table("T").set_properties({"ValidationRule": ">0"}, column="A")
    db.execute("INSERT INTO T (B) VALUES ('x')")
    db.table("T").set_properties({"ValidationRule": "[A]>0"})
    db.execute("INSERT INTO T (B) VALUES ('y')")
    assert db.table("T").row_count == 2


def test_the_validation_text_is_the_message_when_there_is_one() -> None:
    db = _table()
    table = db.table("T")
    table.set_properties({"ValidationRule": ">0"}, column="A")
    table.set_properties({"ValidationText": "must be positive"}, column="A")
    with pytest.raises(AccessError, match="^must be positive$"):
        db.execute("INSERT INTO T (A, B) VALUES (-1, 'x')")


def test_a_table_rule_reads_the_whole_row() -> None:
    db = _table()
    db.table("T").set_properties({"ValidationRule": "[A]>0 And [B]<>'no'"})
    with pytest.raises(AccessError, match=r"set for 'T'\."):
        db.execute("INSERT INTO T (A, B) VALUES (1, 'no')")
    db.execute("INSERT INTO T (A, B) VALUES (1, 'yes')")
    assert db.table("T").row_count == 1


def test_not_null_and_default_come_off_the_ddl() -> None:
    db = _table("CREATE TABLE T (A LONG NOT NULL, B TEXT(20) DEFAULT 'set', C TEXT(30))")
    table = db.table("T")
    assert table.column_properties("A") == {"Required": True}
    assert table.column_properties("B") == {"DefaultValue": "'set'"}
    assert table.column_properties("C") == {}
    db.execute("INSERT INTO T (A) VALUES (1)")
    assert [(r["A"], r["B"]) for r in db.table("T").rows()] == [(1, "set")]
    with pytest.raises(AccessError, match="You must enter a value"):
        db.execute("INSERT INTO T (C) VALUES ('x')")


def test_a_column_rule_keeps_the_nul_the_engine_leaves() -> None:
    db = _table()
    table = db.table("T")
    table.set_properties({"ValidationRule": ">0"}, column="A")
    table.set_properties({"ValidationRule": "[A]>0"})
    assert table.column_properties("A") == {"ValidationRule": ">0" + chr(0)}
    assert table.properties() == {"ValidationRule": "[A]>0"}
    # Written twice, the NUL is not doubled.
    table.set_properties({"ValidationRule": ">0" + chr(0)}, column="A")
    assert db.table("T").column_properties("A") == {"ValidationRule": ">0" + chr(0)}


def test_the_engines_own_properties_take_its_type_and_flags() -> None:
    db = _table()
    table = db.table("T")
    table.set_properties({"DefaultValue": "5", "Required": True}, column="A")
    table.set_properties({"Caption": "shown"}, column="A")
    records = table.property_blob().column_properties["A"]
    assert [(name, v.type, v.flags) for name, v in records.items()] == [
        ("DefaultValue", 12, 1),
        ("Required", 1, 1),
        ("Caption", 10, 0),
    ]


def test_a_system_table_carries_no_rules() -> None:
    db = _table()
    assert not db.table("MSysObjects").rules()
