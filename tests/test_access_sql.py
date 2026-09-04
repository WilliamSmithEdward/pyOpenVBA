"""The SQL executor: SELECT, INSERT, UPDATE and DELETE over the engine."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access._pages import row_bytes
from pyopenvba.access._queries import parse_from
from pyopenvba.access._rows import split_row
from pyopenvba.access._sql import Parser, like_match
from pyopenvba.access._tdef import TYPE_BINARY, TYPE_NUMERIC, TYPE_TEXT
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def _shop(tmp_path: Path) -> AccessDatabase:
    db = AccessDatabase.create_new(tmp_path / "shop.accdb")
    db.create_table(
        "Customers",
        [
            ColumnSpec("Id", "long", autonumber=True),
            ColumnSpec("Name", "text", 50),
            ColumnSpec("City", "text", 30),
            ColumnSpec("Balance", "currency"),
            ColumnSpec("Joined", "datetime"),
            ColumnSpec("Active", "boolean"),
        ],
    )
    db.create_table(
        "Orders",
        [
            ColumnSpec("Id", "long", autonumber=True),
            ColumnSpec("CustomerId", "long"),
            ColumnSpec("Amount", "double"),
            ColumnSpec("Note", "memo"),
        ],
    )
    for statement in (
        "INSERT INTO Customers (Name, City, Balance, Joined, Active) VALUES ('Ada', 'London', 10.5, #3/1/2024#, TRUE)",
        "INSERT INTO Customers (Name, City, Balance, Joined, Active) VALUES ('Bob', 'Paris', -2, #12/31/2023 13:45:00#, FALSE)",
        "INSERT INTO Customers (Name, City, Balance, Joined, Active) VALUES ('Cy', NULL, 0, NULL, TRUE)",
        "INSERT INTO Customers (Name, City) VALUES ('Dee', 'London')",
        "INSERT INTO Orders (CustomerId, Amount, Note) VALUES (1, 100, 'first')",
        "INSERT INTO Orders (CustomerId, Amount, Note) VALUES (1, 250.25, NULL)",
        "INSERT INTO Orders (CustomerId, Amount, Note) VALUES (2, 5, 'small')",
        "INSERT INTO Orders (CustomerId, Amount, Note) VALUES (9, 1, 'orphan')",
    ):
        assert db.execute(statement) == 1
    return db


def _names(rows: list[dict[str, object]] | int) -> list[object]:
    assert isinstance(rows, list)
    return [r["Name"] for r in rows]


def test_select_star_returns_every_column_in_definition_order(tmp_path: Path) -> None:
    rows = _shop(tmp_path).execute("SELECT * FROM Customers")
    assert isinstance(rows, list)
    assert list(rows[0]) == ["Id", "Name", "City", "Balance", "Joined", "Active"]
    assert rows[0]["Balance"] == Decimal("10.5")
    assert rows[1]["Joined"] == dt.datetime(2023, 12, 31, 13, 45)
    assert rows[3]["Balance"] is None and rows[3]["Active"] is False


def test_where_compares_text_case_blind_and_orders_descending(tmp_path: Path) -> None:
    """`Len(Name)` is the third column, so the engine calls it Expr1002:
    the number is 1000 plus where the column sits, not a count of the
    ones before it that needed a name."""
    rows = _shop(tmp_path).execute("SELECT Name, Balance * 2 AS Doubled, Len(Name) FROM Customers WHERE City = 'london' ORDER BY Name DESC")
    assert rows == [
        {"Name": "Dee", "Doubled": None, "Expr1002": 3},
        {"Name": "Ada", "Doubled": Decimal("21.0000"), "Expr1002": 3},
    ]


def test_like_in_between_and_null_tests(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert _names(db.execute("SELECT Name FROM Customers WHERE Name LIKE '*d*' OR Balance < 0")) == ["Ada", "Bob", "Dee"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE City IS NULL")) == ["Cy"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE City IS NOT NULL AND Balance IS NULL")) == ["Dee"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Joined BETWEEN #1/1/2024# AND #12/31/2024#")) == ["Ada"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Name IN ('bob', 'cy') AND NOT Active")) == ["Bob"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Name NOT IN ('bob', 'cy')")) == ["Ada", "Dee"]


def test_null_comparisons_drop_rows_instead_of_matching(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert _names(db.execute("SELECT Name FROM Customers WHERE Balance <> 0")) == ["Ada", "Bob"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE NOT (Balance > 0)")) == ["Bob", "Cy"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Balance > 0 OR City = 'Paris'")) == ["Ada", "Bob"]


def test_top_distinct_and_null_first_ordering(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert _names(db.execute("SELECT TOP 2 Name FROM Customers ORDER BY Id")) == ["Ada", "Bob"]
    assert db.execute("SELECT DISTINCT City FROM Customers ORDER BY City") == [{"City": None}, {"City": "London"}, {"City": "Paris"}]
    assert db.execute("SELECT DISTINCT City FROM Customers ORDER BY City DESC") == [{"City": "Paris"}, {"City": "London"}, {"City": None}]


def test_group_by_having_and_whole_table_aggregates(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    rows = db.execute("SELECT City, Count(*) AS N, Sum(Balance) AS Total, Max(Joined) FROM Customers GROUP BY City HAVING Count(*) > 1 ORDER BY City")
    assert rows == [{"City": "London", "N": 2, "Total": Decimal("10.5000"), "Expr1003": dt.datetime(2024, 3, 1)}]
    assert db.execute("SELECT Count(*), Avg(Amount), Min(Amount), Count(Note) FROM Orders") == [
        {"Expr1000": 4, "Expr1001": 89.0625, "Expr1002": 1.0, "Expr1003": 3}
    ]
    assert db.execute("SELECT Avg(Balance) AS Mean FROM Customers") == [{"Mean": Decimal("2.8333")}]
    assert db.execute("SELECT Count(*) AS N FROM Orders WHERE Amount > 1000") == [{"N": 0}]


def test_inner_left_and_right_joins(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    inner = db.execute("SELECT c.Name, o.Amount FROM Customers AS c INNER JOIN Orders AS o ON c.Id = o.CustomerId ORDER BY o.Amount")
    assert inner == [{"Name": "Bob", "Amount": 5.0}, {"Name": "Ada", "Amount": 100.0}, {"Name": "Ada", "Amount": 250.25}]
    left = db.execute("SELECT c.Name, o.Amount FROM Customers c LEFT JOIN Orders o ON c.Id = o.CustomerId ORDER BY c.Name, o.Amount")
    assert isinstance(left, list)
    assert [(r["Name"], r["Amount"]) for r in left] == [("Ada", 100.0), ("Ada", 250.25), ("Bob", 5.0), ("Cy", None), ("Dee", None)]
    right = db.execute("SELECT Customers.Name, Orders.Note FROM Customers RIGHT JOIN Orders ON Customers.Id = Orders.CustomerId ORDER BY Orders.Id")
    assert isinstance(right, list)
    assert [(r["Name"], r["Note"]) for r in right] == [("Ada", "first"), ("Ada", None), ("Bob", "small"), (None, "orphan")]
    grouped = db.execute("SELECT c.Name, Sum(o.Amount) AS Spent FROM Customers c INNER JOIN Orders o ON c.Id = o.CustomerId GROUP BY c.Name ORDER BY Sum(o.Amount) DESC")
    assert grouped == [{"Name": "Ada", "Spent": 350.25}, {"Name": "Bob", "Spent": 5.0}]


def test_functions_concatenation_and_parameters(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    rows = db.execute("SELECT Name, IIf(Active, 'yes', 'no') AS A, Nz(City, '-') & '!' AS C, UCase(Left(Name, 1)) & LCase(Mid(Name, 2)) AS Cased FROM Customers WHERE Id = 3")
    assert rows == [{"Name": "Cy", "A": "yes", "C": "-!", "Cased": "Cy"}]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Balance > [MinBalance]", {"MinBalance": 0})) == ["Ada"]
    assert db.execute("SELECT Year(Joined) AS Y, Month(Joined) AS M, Day(Joined) AS D FROM Customers WHERE Id = 2") == [{"Y": 2023, "M": 12, "D": 31}]
    assert db.execute("SELECT Round(Amount / 3, 2) AS R, Int(Amount) AS I, Abs(-Amount) AS A FROM Orders WHERE Id = 2") == [{"R": 83.42, "I": 250, "A": 250.25}]
    assert db.execute("SELECT InStr(Name, 'b') AS P, Trim('  x  ') AS T, Right(Name, 2) AS R FROM Customers WHERE Id = 2") == [{"P": 1, "T": "x", "R": "ob"}]


def test_update_delete_and_insert_select(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert db.execute("UPDATE Customers SET Balance = Balance + 1, City = UCase(City) WHERE Active = TRUE") == 2
    rows = db.execute("SELECT Name, Balance, City FROM Customers ORDER BY Id")
    assert isinstance(rows, list)
    assert [(r["Balance"], r["City"]) for r in rows] == [
        (Decimal("11.5000"), "LONDON"),
        (Decimal("-2.0000"), "Paris"),
        (Decimal("1.0000"), None),
        (None, "London"),
    ]
    assert db.execute("DELETE FROM Orders WHERE Amount < 10") == 2
    assert db.execute("INSERT INTO Orders (CustomerId, Amount) SELECT Id, Balance FROM Customers WHERE Balance > 0") == 2
    rows = db.execute("SELECT CustomerId, Amount FROM Orders ORDER BY Id")
    assert isinstance(rows, list)
    assert [(r["CustomerId"], r["Amount"]) for r in rows] == [(1, 100.0), (1, 250.25), (1, 11.5), (3, 1.0)]
    assert db.execute("DELETE FROM Orders") == 4
    assert db.execute("SELECT Count(*) AS N FROM Orders") == [{"N": 0}]
    reopened = AccessDatabase(db.to_bytes())
    assert reopened.table("Orders").row_count == 0
    assert reopened.table("Customers").row_count == 4


def test_values_are_coerced_to_the_column_type(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert db.execute("INSERT INTO Customers (Name, Balance, Active, Joined) VALUES (42, 7, 1, '2020-02-03')") == 1
    rows = db.execute("SELECT Name, Balance, Active, Joined FROM Customers WHERE Id = 5")
    assert rows == [{"Name": "42", "Balance": Decimal("7.0000"), "Active": True, "Joined": dt.datetime(2020, 2, 3)}]
    assert db.execute("UPDATE Orders SET CustomerId = 2.6 WHERE Id = 1") == 1
    assert db.execute("SELECT CustomerId FROM Orders WHERE Id = 1") == [{"CustomerId": 3}]


def test_a_value_list_with_no_column_list_names_every_column(tmp_path: Path) -> None:
    """The engine counts the columns, AutoNumber included, and refuses a
    list of any other length."""
    db = AccessDatabase.create_new(tmp_path / "wide.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long"), ColumnSpec("B", "text", 20)])
    with pytest.raises(AccessError, match="different number of columns"):
        db.execute("INSERT INTO T VALUES (7, 'seven')")
    assert db.execute("INSERT INTO T VALUES (5, 7, 'seven')") == 1
    assert db.execute("SELECT * FROM T") == [{"Id": 5, "A": 7, "B": "seven"}]


def test_an_explicit_autonumber_is_kept_and_moves_the_counter(tmp_path: Path) -> None:
    """The counter follows the number written, not the larger of the two:
    10 then 3 leaves the next row at 4."""
    db = AccessDatabase.create_new(tmp_path / "counter.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long")])
    db.execute("INSERT INTO T VALUES (10, 1)")
    db.execute("INSERT INTO T VALUES (3, 2)")
    db.execute("INSERT INTO T (A) VALUES (3)")
    assert db.execute("SELECT Id, A FROM T ORDER BY A") == [{"Id": 10, "A": 1}, {"Id": 3, "A": 2}, {"Id": 4, "A": 3}]


def test_an_autonumber_below_zero_counts_up_from_there(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "negative.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long")])
    db.execute("INSERT INTO T VALUES (-5, 1)")
    db.execute("INSERT INTO T (A) VALUES (2)")
    assert db.execute("SELECT Id FROM T ORDER BY A") == [{"Id": -5}, {"Id": -4}]
    assert db.table("T").definition.next_autonumber == 0xFFFFFFFC


def test_an_autonumber_column_is_not_updateable_and_takes_no_null(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "auto.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long")])
    db.execute("INSERT INTO T (A) VALUES (1)")
    with pytest.raises(AccessError, match="AutoNumber, which no query updates"):
        db.execute("UPDATE T SET Id = 77")
    with pytest.raises(AccessError, match="AutoNumber, which takes no Null"):
        db.execute("INSERT INTO T VALUES (NULL, 7)")
    assert db.execute("SELECT Id, A FROM T") == [{"Id": 1, "A": 1}]


def test_a_query_truncates_text_to_the_column_but_not_a_memo(tmp_path: Path) -> None:
    """Assigning over-long text to a field is refused; a query cuts it to
    fit, which is why the same value behaves differently by the door it
    comes through."""
    db = AccessDatabase.create_new(tmp_path / "text.accdb")
    db.create_table("T", [ColumnSpec("Id", "long"), ColumnSpec("T", "text", 3), ColumnSpec("M", "memo")])
    db.execute("INSERT INTO T VALUES (1, 'abcdef', String(300, 'x'))")
    db.execute("INSERT INTO T VALUES (2, 'ab' & 'cdef', NULL)")
    db.execute("UPDATE T SET T = T & T & T WHERE Id = 2")
    rows = db.execute("SELECT Id, T, Len(M) AS N FROM T ORDER BY Id")
    assert rows == [{"Id": 1, "T": "abc", "N": 300}, {"Id": 2, "T": "abc", "N": None}]
    with pytest.raises(AccessError, match="exceed its size"):
        db.table("T").insert_row({"Id": 3, "T": "abcdef"})


def test_a_byte_takes_the_low_byte_and_the_others_overflow(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "byte.accdb")
    db.create_table("T", [ColumnSpec("Id", "long"), ColumnSpec("B", "byte"), ColumnSpec("S", "integer")])
    db.execute("INSERT INTO T VALUES (1, 300, 1)")
    db.execute("INSERT INTO T VALUES (2, -1, 2)")
    db.execute("INSERT INTO T VALUES (3, 255.5, 3)")
    assert db.execute("SELECT B FROM T ORDER BY Id") == [{"B": 44}, {"B": 255}, {"B": 0}]
    with pytest.raises(AccessError, match="70000 is not an Integer"):
        db.execute("INSERT INTO T VALUES (4, 1, 70000)")


def test_a_number_written_to_a_date_column_is_its_serial(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "date.accdb")
    db.create_table("T", [ColumnSpec("Id", "long"), ColumnSpec("D", "datetime")])
    db.execute("INSERT INTO T VALUES (1, 44000)")
    assert db.execute("SELECT D FROM T") == [{"D": dt.datetime(2020, 6, 18)}]
    with pytest.raises(AccessError, match="cannot read date"):
        db.execute("INSERT INTO T VALUES (2, 'not a date')")


def test_an_action_query_that_fails_part_way_writes_nothing(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "atomic.accdb")
    db.create_table("Src", [ColumnSpec("Id", "long"), ColumnSpec("A", "long")])
    db.create_table(
        "Dst",
        [ColumnSpec("Id", "long"), ColumnSpec("A", "long")],
        [IndexSpec("PK", ("Id",), primary=True)],
    )
    db.execute("INSERT INTO Src VALUES (1, 1)")
    db.execute("INSERT INTO Src VALUES (1, 2)")
    with pytest.raises(AccessError):
        db.execute("INSERT INTO Dst SELECT * FROM Src")
    assert db.execute("SELECT Count(*) AS N FROM Dst") == [{"N": 0}]
    db.table("Dst").set_properties({"ValidationRule": "<100"}, column="A")
    db.execute("INSERT INTO Dst VALUES (1, 50)")
    db.execute("INSERT INTO Dst VALUES (2, 50)")
    with pytest.raises(AccessError):
        db.execute("UPDATE Dst SET A = A + Id * 40")
    assert db.execute("SELECT A FROM Dst ORDER BY Id") == [{"A": 50}, {"A": 50}]
    # The bytes come back too, not just the rows -- except the header's
    # row count, which keeps the one row the refused INSERT had written
    # before its second row failed (measured: the engine's rollback puts
    # the rows back and leaves the counters).
    reopened = AccessDatabase(db.to_bytes())
    assert reopened.execute("SELECT Count(*) AS N FROM Dst") == [{"N": 2}]
    assert reopened.table("Dst").row_count == 3


def _pair(tmp_path: Path) -> AccessDatabase:
    """Two tables joined on Id, with one Id matched twice and one not at all."""
    db = AccessDatabase.create_new(tmp_path / "pair.accdb")
    db.create_table("L", [ColumnSpec("Id", "long"), ColumnSpec("A", "long")])
    db.create_table("R", [ColumnSpec("Id", "long"), ColumnSpec("B", "long")])
    for statement in (
        "INSERT INTO L VALUES (1, 0)",
        "INSERT INTO L VALUES (2, 0)",
        "INSERT INTO L VALUES (3, 0)",
        "INSERT INTO R VALUES (1, 111)",
        "INSERT INTO R VALUES (2, 222)",
        "INSERT INTO R VALUES (2, 333)",
    ):
        db.execute(statement)
    return db


def test_update_over_a_join_writes_both_tables_and_counts_join_rows(tmp_path: Path) -> None:
    """Three join rows, so three affected.  L.Id 2 is reached twice and
    keeps the last join row's value; with three rows a side, the engine
    scans R and probes L, so the join rows come in R's order and 333 is
    the one that stays (measured)."""
    db = _pair(tmp_path)
    assert db.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET L.A = R.B, R.B = 0") == 3
    assert db.execute("SELECT A FROM L ORDER BY Id") == [{"A": 111}, {"A": 333}, {"A": 0}]
    assert db.execute("SELECT Count(*) AS N FROM R WHERE B = 0") == [{"N": 3}]


def test_update_over_a_left_join_reaches_the_unmatched_rows(tmp_path: Path) -> None:
    db = _pair(tmp_path)
    db.execute("UPDATE L SET A = 9")
    assert db.execute("UPDATE L LEFT JOIN R ON L.Id = R.Id SET L.A = R.B WHERE L.Id <> 2") == 2
    assert db.execute("SELECT A FROM L ORDER BY Id") == [{"A": 111}, {"A": 9}, {"A": None}]


def test_update_names_its_columns_by_alias_or_by_the_table_that_holds_them(tmp_path: Path) -> None:
    db = _pair(tmp_path)
    assert db.execute("UPDATE L AS X INNER JOIN R AS Y ON X.Id = Y.Id SET [X].[A] = Y.B WHERE Y.B > 200") == 2
    assert db.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET A = 5 WHERE R.B = 111") == 1
    assert db.execute("SELECT A FROM L ORDER BY Id") == [{"A": 5}, {"A": 333}, {"A": 0}]
    with pytest.raises(AccessError, match="ambiguous"):
        db.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET Id = 7")
    with pytest.raises(AccessError, match="not a table of the statement"):
        db.execute("UPDATE L SET Q.A = 7")


def test_a_set_clause_cannot_hold_a_subquery_but_a_where_clause_can(tmp_path: Path) -> None:
    db = _pair(tmp_path)
    with pytest.raises(AccessError, match="not updateable"):
        db.execute("UPDATE L SET A = (SELECT Max(B) FROM R) WHERE Id = 1")
    assert db.execute("UPDATE L SET A = DMax('B', 'R') WHERE Id IN (SELECT Id FROM R WHERE B > 300)") == 1
    assert db.execute("SELECT A FROM L ORDER BY Id") == [{"A": 0}, {"A": 333}, {"A": 0}]


def test_delete_over_a_join_names_the_table_and_refuses_a_double_match(tmp_path: Path) -> None:
    db = _pair(tmp_path)
    with pytest.raises(AccessError, match="name the table"):
        db.execute("DELETE FROM L INNER JOIN R ON L.Id = R.Id")
    with pytest.raises(AccessError, match="more than once"):
        db.execute("DELETE L.* FROM L INNER JOIN R ON L.Id = R.Id")
    # Nothing was deleted, but the header's row count keeps the two rows
    # the engine had taken out before it reached L.Id 2 a second time.
    assert db.execute("SELECT Count(*) AS N FROM L") == [{"N": 3}]
    assert db.table("L").row_count == 1
    assert db.execute("DELETE L.* FROM L INNER JOIN R ON L.Id = R.Id WHERE R.B = 111") == 1
    assert db.execute("DELETE R.* FROM L INNER JOIN R ON L.Id = R.Id") == 2
    assert db.execute("SELECT Id FROM L ORDER BY Id") == [{"Id": 2}, {"Id": 3}]
    assert db.execute("SELECT B FROM R") == [{"B": 111}]


def _types(db: AccessDatabase, table: str) -> list[tuple[str, str, object, bool]]:
    """Each column as (name, type, size, fixed) the way the header says it."""
    out: list[tuple[str, str, object, bool]] = []
    for column in db.table(table).definition.columns:
        size: object = None
        if column.type_code == TYPE_TEXT:
            size = column.length // 2
        elif column.type_code == TYPE_NUMERIC:
            size = (column.precision, column.scale)
        elif column.type_code == TYPE_BINARY:
            size = column.length
        out.append((column.name, column.type_name, size, column.is_fixed))
    return out


def test_select_into_makes_the_table_the_engine_makes(tmp_path: Path) -> None:
    """A column read straight out keeps its definition, AutoNumber and
    all; an expression gets the type the engine gives it, and sits among
    the variable-length columns unless it is an Integer.  Text from an
    expression is 255 characters, a number with a fraction is a Decimal
    scaled to its digits, and Null is a Binary as wide as a Text(255)."""
    db = _shop(tmp_path)
    assert db.execute(
        "SELECT Id, Name, Balance * 2 AS Doubled, Name & City AS Joined, 5.5 AS Dec, "
        "Len(Name) AS N, Active AND TRUE AS Flag, Joined + 1 AS Later, Null AS Nothing "
        "INTO Made FROM Customers WHERE Id = 1"
    ) == 1
    assert _types(db, "Made") == [
        ("Id", "Long", None, True),
        ("Name", "Text", 50, False),
        ("Doubled", "Currency", None, False),
        ("Joined", "Text", 255, False),
        ("Dec", "Decimal", (28, 1), False),
        ("N", "Long", None, False),
        ("Flag", "Integer", None, True),
        ("Later", "DateTime", None, False),
        ("Nothing", "Binary", 510, False),
    ]
    made = db.table("Made").definition
    assert made.column("Id").auto_number and not made.column("Joined").compressed_unicode
    # Bytes 9-10 of a made header count from one; every column after the
    # Decimal but the Text carries its precision and scale (measured).
    assert [c.header_ordinal for c in made.columns] == list(range(1, 10))
    assert (made.column("N").sort_order, made.column("N").sort_version) == (28, 1)
    assert (made.column("Later").sort_order, made.column("Later").sort_version) == (28, 1)
    assert made.column("Doubled").sort_order == 1033
    assert db.execute("SELECT Doubled, Dec, N, Flag, Nothing FROM Made") == [
        {"Doubled": Decimal("21.0000"), "Dec": Decimal("5.5"), "N": 3, "Flag": -1, "Nothing": None}
    ]
    assert not db.table("Made").definition.logical_indexes


def test_select_into_over_a_join_and_a_group_and_over_nothing(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert db.execute(
        "SELECT Customers.Name, Sum(Orders.Amount) AS Total, Count(*) AS N INTO Totals "
        "FROM Customers INNER JOIN Orders ON Customers.Id = Orders.CustomerId GROUP BY Customers.Name"
    ) == 2
    assert _types(db, "Totals") == [("Name", "Text", 50, False), ("Total", "Double", None, False), ("N", "Long", None, False)]
    assert db.execute("SELECT * FROM Totals ORDER BY Name") == [
        {"Name": "Ada", "Total": 350.25, "N": 2},
        {"Name": "Bob", "Total": 5.0, "N": 1},
    ]
    # No rows still makes the table, and an existing name is refused.
    assert db.execute("SELECT Id, Name INTO Nobody FROM Customers WHERE Id = 999") == 0
    assert db.table("Nobody").row_count == 0
    with pytest.raises(AccessError, match="already exists"):
        db.execute("SELECT * INTO Nobody FROM Customers")
    assert db.execute("SELECT (SELECT Max(Amount) FROM Orders) AS M, (SELECT Name FROM Customers WHERE Id = 2) AS Who INTO Sub FROM Customers WHERE Id = 1") == 1
    assert _types(db, "Sub") == [("M", "Double", None, False), ("Who", "Text", 255, False)]


def test_a_number_written_with_an_exponent_reads(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert db.execute("SELECT 1E3 AS A, 1.5E2 AS B, 1.5E-1 AS C, -1E3 AS D FROM Customers WHERE Id = 1") == [
        {"A": 1000, "B": 150, "C": Decimal("0.15"), "D": -1000}
    ]
    db.execute("SELECT 1E3 AS A, 1.5E-1 AS C, 1E10 AS Big INTO Exps FROM Customers WHERE Id = 1")
    assert _types(db, "Exps") == [("A", "Long", None, False), ("C", "Decimal", (28, 2), False), ("Big", "Decimal", (28, 0), False)]


def test_short_text_is_stored_without_compression(tmp_path: Path) -> None:
    """The engine compresses a Text value only when that makes it shorter,
    so one and two Latin-1 characters go in as plain UTF-16 (measured
    through INSERT, UPDATE and SELECT INTO)."""
    db = AccessDatabase.create_new(tmp_path / "short.accdb")
    db.create_table("T", [ColumnSpec("Id", "long"), ColumnSpec("T", "text", 10, compressed=True)])
    for value in ("x", "xy", "xyz", ""):
        db.execute(f"INSERT INTO T VALUES ({len(value)}, '{value}')")
    definition = db.table("T").definition
    raw: dict[int, bytes | None] = {}
    for row_id, values in db.table("T").rows_with_ids():
        stored = row_bytes(db.store.read(row_id.page), row_id.slot)
        assert stored is not None
        raw[int(values["Id"])] = split_row(definition, stored).values.get(1)  # pyright: ignore[reportArgumentType]
    assert raw[1] == "x".encode("utf-16-le")
    assert raw[2] == "xy".encode("utf-16-le")
    assert raw[3] == b"\xff\xfexyz"
    assert db.execute("SELECT T FROM T WHERE Id > 0 ORDER BY Id") == [{"T": "x"}, {"T": "xy"}, {"T": "xyz"}]


def _sized(tmp_path: Path, left: int, right: list[tuple[int, int]]) -> AccessDatabase:
    """L with Ids 1..left, R with the given (Id, B) rows, in that order."""
    db = AccessDatabase.create_new(tmp_path / f"sized{left}.accdb")
    db.create_table("L", [ColumnSpec("Id", "long"), ColumnSpec("A", "long")])
    db.create_table("R", [ColumnSpec("Id", "long"), ColumnSpec("B", "long")])
    for i in range(1, left + 1):
        db.execute(f"INSERT INTO L VALUES ({i}, 0)")
    for i, b in right:
        db.execute(f"INSERT INTO R VALUES ({i}, {b})")
    return db


def test_a_join_scans_the_smaller_side_and_probes_the_other_last_in_first_out(tmp_path: Path) -> None:
    """Measured on the engine: with two rows against three, L is scanned
    and R's two rows for Id 2 come back newest first; with three against
    three the second-listed table is scanned and they come in stored
    order.  What a SELECT with no ORDER BY answers depends on it, and so
    does which join row's write an UPDATE keeps."""
    rows = [(1, 111), (2, 222), (2, 333)]
    two = _sized(tmp_path, 2, rows)
    assert two.execute("SELECT L.Id, R.B FROM L INNER JOIN R ON L.Id = R.Id") == [
        {"Id": 1, "B": 111}, {"Id": 2, "B": 333}, {"Id": 2, "B": 222}
    ]
    three = _sized(tmp_path, 3, rows)
    assert three.execute("SELECT L.Id, R.B FROM L INNER JOIN R ON L.Id = R.Id") == [
        {"Id": 1, "B": 111}, {"Id": 2, "B": 222}, {"Id": 2, "B": 333}
    ]
    assert three.execute("SELECT L.Id, R.B FROM R INNER JOIN L ON L.Id = R.Id") == [
        {"Id": 1, "B": 111}, {"Id": 2, "B": 333}, {"Id": 2, "B": 222}
    ]
    # The last join row's write is the one that stays, so the two shapes
    # leave different values behind.
    assert two.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET L.A = R.B") == 3
    assert two.execute("SELECT A FROM L WHERE Id = 2") == [{"A": 222}]
    assert three.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET L.A = R.B") == 3
    assert three.execute("SELECT A FROM L WHERE Id = 2") == [{"A": 333}]


def test_a_later_join_row_reads_what_an_earlier_one_wrote(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "revisit.accdb")
    db.create_table("L", [ColumnSpec("Id", "long"), ColumnSpec("A", "long")])
    db.create_table("R", [ColumnSpec("Id", "long"), ColumnSpec("B", "long")])
    for statement in ("INSERT INTO L VALUES (1, 5)", "INSERT INTO L VALUES (1, 7)", "INSERT INTO R VALUES (1, 100)"):
        db.execute(statement)
    # R is scanned (one row), L probed newest first: 100 + 7, then 107 + 5.
    assert db.execute("UPDATE L INNER JOIN R ON L.Id = R.Id SET R.B = R.B + L.A, L.A = R.B") == 2
    assert db.execute("SELECT B FROM R") == [{"B": 112}]
    assert db.execute("SELECT A FROM L ORDER BY Id, A") == [{"A": 100}, {"A": 107}]


def test_joins_in_parentheses_read_as_access_writes_them(tmp_path: Path) -> None:
    """Every query Access saves over three tables is
    ``(A INNER JOIN B ON ...) INNER JOIN C ON ...``; the group can also be
    the right side.  The rows come out in the engine's order either way."""
    db = AccessDatabase.create_new(tmp_path / "three.accdb")
    db.create_table("A", [ColumnSpec("Id", "long"), ColumnSpec("X", "long")])
    db.create_table("B", [ColumnSpec("Id", "long"), ColumnSpec("Y", "long")])
    db.create_table("C", [ColumnSpec("Id", "long"), ColumnSpec("Z", "long")])
    for statement in ("INSERT INTO A VALUES (1, 1)", "INSERT INTO A VALUES (1, 2)", "INSERT INTO B VALUES (1, 10)",
                      "INSERT INTO B VALUES (1, 20)", "INSERT INTO B VALUES (1, 30)", "INSERT INTO C VALUES (1, 100)", "INSERT INTO C VALUES (1, 200)"):
        db.execute(statement)
    left_deep = db.execute("SELECT A.X, B.Y, C.Z FROM (A INNER JOIN B ON A.Id = B.Id) INNER JOIN C ON A.Id = C.Id")
    assert left_deep == [{"X": x, "Y": y, "Z": z} for x in (1, 2) for y in (30, 20, 10) for z in (200, 100)]
    right_nested = db.execute("SELECT A.X, B.Y, C.Z FROM C INNER JOIN (A INNER JOIN B ON A.Id = B.Id) ON C.Id = A.Id")
    assert right_nested == [{"X": x, "Y": y, "Z": z} for x in (1, 2) for z in (200, 100) for y in (30, 20, 10)]
    tables, joins = parse_from("(A INNER JOIN B ON A.Id = B.Id) INNER JOIN C ON A.Id = C.Id")
    assert [t.name1 for t in tables] == ["A", "B", "C"]
    assert [(j.name1, j.name2, j.expression) for j in joins] == [("A", "B", "A.Id = B.Id"), ("B", "C", "A.Id = C.Id")]
    assert db.execute("SELECT A.X, C.Z FROM A, C") == [{"X": 1, "Z": 100}, {"X": 2, "Z": 100}, {"X": 1, "Z": 200}, {"X": 2, "Z": 200}]


def test_groups_and_distinct_rows_come_out_in_key_order(tmp_path: Path) -> None:
    """The engine groups by sorting: Null first, then ascending, text
    case-blind, and a group wears the first value it met (measured)."""
    db = AccessDatabase.create_new(tmp_path / "groups.accdb")
    db.create_table("T", [ColumnSpec("Id", "long"), ColumnSpec("A", "long"), ColumnSpec("N", "text", 10), ColumnSpec("B", "long")])
    for statement in ("INSERT INTO T VALUES (1, 30, 'pear', 2)", "INSERT INTO T VALUES (2, 10, 'Apple', 1)", "INSERT INTO T VALUES (3, 20, 'fig', 2)",
                      "INSERT INTO T VALUES (4, 10, 'apple', 1)", "INSERT INTO T VALUES (5, NULL, 'Fig', 2)", "INSERT INTO T VALUES (6, 20, NULL, 1)",
                      "INSERT INTO T VALUES (7, -5, 'zed', 3)"):
        db.execute(statement)
    assert db.execute("SELECT A, Count(*) AS C FROM T GROUP BY A") == [
        {"A": None, "C": 1}, {"A": -5, "C": 1}, {"A": 10, "C": 2}, {"A": 20, "C": 2}, {"A": 30, "C": 1}
    ]
    assert db.execute("SELECT N, Count(*) AS C FROM T GROUP BY N") == [
        {"N": None, "C": 1}, {"N": "Apple", "C": 2}, {"N": "fig", "C": 2}, {"N": "pear", "C": 1}, {"N": "zed", "C": 1}
    ]
    assert db.execute("SELECT A, B, Count(*) AS C FROM T GROUP BY B, A") == [
        {"A": 10, "B": 1, "C": 2}, {"A": 20, "B": 1, "C": 1}, {"A": None, "B": 2, "C": 1},
        {"A": 20, "B": 2, "C": 1}, {"A": 30, "B": 2, "C": 1}, {"A": -5, "B": 3, "C": 1},
    ]
    assert db.execute("SELECT DISTINCT B, A FROM T") == [
        {"B": 1, "A": 10}, {"B": 1, "A": 20}, {"B": 2, "A": None}, {"B": 2, "A": 20}, {"B": 2, "A": 30}, {"B": 3, "A": -5}
    ]
    assert db.execute("SELECT DISTINCT N FROM T ORDER BY N DESC") == [{"N": "zed"}, {"N": "pear"}, {"N": "fig"}, {"N": "Apple"}, {"N": None}]


def test_a_refused_insert_keeps_the_numbers_it_reserved(tmp_path: Path) -> None:
    """Every row an INSERT attempts takes the next AutoNumber before its
    values are looked at, and the rows written before the failing one
    stay counted in the header; the rollback puts the rows back and leaves
    both (measured).  A refused UPDATE moves neither."""
    db = AccessDatabase.create_new(tmp_path / "reserved.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long"), ColumnSpec("D", "datetime")])
    db.create_table("S", [ColumnSpec("A", "long")])
    for statement in ("INSERT INTO T (A) VALUES (1)", "INSERT INTO T (A) VALUES (2)", "INSERT INTO S VALUES (10)",
                      "INSERT INTO S VALUES (20)", "INSERT INTO S VALUES (500)", "INSERT INTO S VALUES (40)"):
        db.execute(statement)
    db.table("T").set_properties({"ValidationRule": "<100"}, column="A")
    with pytest.raises(AccessError):
        db.execute("INSERT INTO T (A, D) VALUES (9, 'not a date')")
    assert db.table("T").definition.next_autonumber == 3
    with pytest.raises(AccessError):
        db.execute("INSERT INTO T (A) SELECT A FROM S")
    assert db.table("T").definition.next_autonumber == 6
    assert db.table("T").row_count == 4 and db.execute("SELECT Count(*) AS N FROM T") == [{"N": 2}]
    with pytest.raises(AccessError):
        db.execute("INSERT INTO T (Id, A) VALUES (50, 900)")
    assert db.table("T").definition.next_autonumber == 7
    with pytest.raises(AccessError):
        db.execute("UPDATE T SET A = 900")
    assert db.table("T").definition.next_autonumber == 7 and db.table("T").row_count == 4
    assert db.execute("INSERT INTO T (A) VALUES (3)") == 1
    assert db.execute("SELECT Id FROM T WHERE A = 3") == [{"Id": 8}]


def test_a_statement_writing_its_own_autonumbers_ends_at_the_largest(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "largest.accdb")
    db.create_table("T", [ColumnSpec("Id", "long", autonumber=True), ColumnSpec("A", "long")])
    db.create_table("S", [ColumnSpec("Id", "long"), ColumnSpec("A", "long")])
    db.execute("INSERT INTO S VALUES (30, 1)")
    db.execute("INSERT INTO S VALUES (10, 2)")
    assert db.execute("INSERT INTO T SELECT Id, A FROM S") == 2
    assert db.table("T").definition.next_autonumber == 30
    db.execute("SELECT Id, A INTO M FROM S")
    assert db.table("M").definition.next_autonumber == 0  # a copied Long is no AutoNumber
    db.execute("SELECT Id, A INTO N FROM T")
    assert db.table("N").definition.next_autonumber == 30
    assert db.execute("INSERT INTO N (A) VALUES (9)") == 1 and db.execute("SELECT Max(Id) AS M FROM N") == [{"M": 31}]


def _numbers(tmp_path: Path) -> AccessDatabase:
    """One row of every numeric type: Long 10, Double 1.5, Currency 2.5,
    Byte 3, Integer 4, Single 1.25, a Date, a Decimal(18,4) 0.0625 and a
    Large Number 1e10."""
    db = AccessDatabase.create_new(tmp_path / "numbers.accdb")
    db.create_table(
        "N",
        [ColumnSpec("Id", "long"), ColumnSpec("L", "long"), ColumnSpec("D", "double"), ColumnSpec("C", "currency"),
         ColumnSpec("Y", "byte"), ColumnSpec("S", "integer"), ColumnSpec("F", "single"), ColumnSpec("Dt", "datetime"),
         ColumnSpec("X", "decimal", (18, 4)), ColumnSpec("H", "bigint"), ColumnSpec("T", "text", 10)],
    )
    db.execute("INSERT INTO N VALUES (1, 10, 1.5, 2.5, 3, 4, 1.25, #2020-01-02#, 0.0625, 10000000000, '7')")
    return db


def _one(db: AccessDatabase, expression: str) -> object:
    rows = db.execute(f"SELECT {expression} AS V FROM N")
    assert isinstance(rows, list)
    return rows[0]["V"]


def test_whole_double_and_currency_arithmetic_keep_the_engine_types(tmp_path: Path) -> None:
    """Measured over every pair of types: whole numbers stay whole, a
    Double or Single makes a Double, Currency stays Currency against a
    whole number and under `+` against a Double, but `*` or `/` with a
    Double is a Double, and `/` of whole numbers is a Double."""
    db = _numbers(tmp_path)
    assert _one(db, "L * 2") == 20 and isinstance(_one(db, "L * 2"), int)
    assert _one(db, "S + Y") == 7 and isinstance(_one(db, "S + Y"), int)
    assert _one(db, "L / 2") == 5.0 and isinstance(_one(db, "L / 2"), float)
    assert _one(db, "L + D") == 11.5 and isinstance(_one(db, "L + D"), float)
    assert _one(db, "F * F") == 1.5625 and isinstance(_one(db, "F * F"), float)
    assert _one(db, "C + 1") == Decimal("3.5000") and isinstance(_one(db, "C + 1"), Decimal)
    assert _one(db, "C * C") == Decimal("6.2500") and str(_one(db, "C * C")) == "6.2500"
    assert _one(db, "C + D") == Decimal("4.0000") and isinstance(_one(db, "C + D"), Decimal)
    assert _one(db, "C * D") == 3.75 and isinstance(_one(db, "C * D"), float)
    assert _one(db, "C / 2") == 1.25 and isinstance(_one(db, "C / 2"), float)
    assert _one(db, "T + 1") == 8.0 and isinstance(_one(db, "T + 1"), float)
    assert _one(db, "-C") == Decimal("-2.5000") and _one(db, "-L") == -10


def test_a_literal_with_a_fraction_is_a_decimal_and_two_decimals_need_one_scale(tmp_path: Path) -> None:
    """`5.5` is a Decimal, `1E3` a Long and `3000000000` a Decimal; two
    Decimals stay Decimal only when their scales agree, else a Double; a
    Decimal against a whole number stays Decimal, against a Double becomes
    one; a quotient of Decimals runs to 28 places (measured)."""
    db = _numbers(tmp_path)
    assert _one(db, "5.5") == Decimal("5.5") and isinstance(_one(db, "5.5"), Decimal)
    assert _one(db, "1E3") == 1000 and isinstance(_one(db, "1E3"), int)
    assert isinstance(_one(db, "3000000000"), Decimal) and isinstance(_one(db, "1.5E-1"), Decimal)
    assert _one(db, "5.5 + 6.5") == Decimal("12.0") and isinstance(_one(db, "5.5 + 6.5"), Decimal)
    assert _one(db, "5.5 + 2.25") == 7.75 and isinstance(_one(db, "5.5 + 2.25"), float)
    assert _one(db, "2.50 + 2.5") == Decimal("5.0") and isinstance(_one(db, "2.50 + 2.5"), Decimal)
    assert _one(db, "5.5 * 5.5") == Decimal("30.25")
    assert _one(db, "5.5 + L") == Decimal("15.5") and isinstance(_one(db, "5.5 + L"), Decimal)
    assert _one(db, "5.5 + D") == 7.0 and isinstance(_one(db, "5.5 + D"), float)
    assert _one(db, "5.5 / 6.5") == Decimal("0.8461538461538461538461538462")
    assert _one(db, "5.5 / 2") == Decimal("2.75")
    assert _one(db, "X / 3") == Decimal("0.0208333333333333333333333333")
    assert _one(db, "X + 1.2345") == Decimal("1.2970") and isinstance(_one(db, "X + 5.5"), float)
    assert _one(db, "C + 5.5") == Decimal("8.0000") and isinstance(_one(db, "C * 1.5"), float)


def test_currency_plus_a_decimal_bearing_expression_is_a_decimal(tmp_path: Path) -> None:
    """A Currency added to an expression that holds a Decimal anywhere in it
    answers a Decimal even where that expression's own value is a Double;
    Abs, CDbl and Round launder the Decimal away, Int does not (measured)."""
    db = _numbers(tmp_path)
    assert isinstance(_one(db, "D * 5.5"), float)
    assert _one(db, "C + D * 5.5") == Decimal("10.7500") and isinstance(_one(db, "C + D * 5.5"), Decimal)
    assert _one(db, "C + (0.125 + 0.5)") == Decimal("3.1250")
    assert _one(db, "C + Int(5.5)") == Decimal("7.5000")
    assert _one(db, "C + Abs(5.5)") == Decimal("8.0000") and str(_one(db, "C + Abs(5.5)")) == "8.0000"
    assert isinstance(_one(db, "L + (0.125 + 0.5)"), float)
    assert isinstance(_one(db, "C * (0.125 + 0.5)"), float)


def test_dates_and_large_numbers_through_the_operators(tmp_path: Path) -> None:
    """A Date plus or less a number is a Date, a Date less a Date a Double,
    a Date times anything a number; a Large Number takes everything into
    itself, rounded half to even, except a Currency, which lands one above
    the floor (measured, an artefact of the engine's)."""
    db = _numbers(tmp_path)
    assert _one(db, "Dt + 1") == dt.datetime(2020, 1, 3)
    assert _one(db, "Dt - Dt") == 0.0 and isinstance(_one(db, "Dt - Dt"), float)
    assert isinstance(_one(db, "Dt * 2"), float) and _one(db, "-Dt") == dt.datetime(1779, 12, 27)
    assert _one(db, "Dt + 5.5") == Decimal("43837.5")
    assert _one(db, "H + 1.5") == 10000000002 and _one(db, "H + 0.5") == 10000000000 and _one(db, "H + 2.5") == 10000000002
    assert _one(db, "H + C") == 10000000003 and _one(db, "H - C") == 9999999998
    assert isinstance(_one(db, "H + D"), int) and _one(db, "H / 4") == 2500000000


def test_functions_and_aggregates_answer_in_the_engine_types(tmp_path: Path) -> None:
    db = _numbers(tmp_path)
    assert _one(db, "Abs(L)") == 10.0 and isinstance(_one(db, "Abs(L)"), float)
    assert _one(db, "Abs(C)") == 2.5 and isinstance(_one(db, "Abs(C)"), float)
    assert _one(db, "Int(L)") == 10 and isinstance(_one(db, "Int(L)"), int)
    assert _one(db, "Int(D)") == 1.0 and isinstance(_one(db, "Int(D)"), float)
    assert _one(db, "Int(C)") == Decimal("2.0000") and isinstance(_one(db, "Int(C)"), Decimal)
    assert _one(db, "Int(Dt)") == dt.datetime(2020, 1, 2)
    assert _one(db, "Round(L)") == 10.0 and isinstance(_one(db, "Round(L)"), float)
    assert _one(db, "Val('12')") == 12.0 and isinstance(_one(db, "Val('12')"), float)
    assert _one(db, "CCur(L)") == Decimal("10.0000") and _one(db, "CLng(D)") == 2
    assert _one(db, "Sum(L)") == 10.0 and isinstance(_one(db, "Sum(L)"), float)
    assert _one(db, "Sum(C)") == Decimal("2.5000") and isinstance(_one(db, "Sum(C)"), Decimal)
    assert _one(db, "Sum(X)") == Decimal("0.0625") and _one(db, "Avg(X)") == Decimal("0.0625")
    assert _one(db, "Avg(C * 1.5)") == Decimal("3.75") and isinstance(_one(db, "Sum(C * 1.5)"), float)
    assert _one(db, "Max(C)") == Decimal("2.5000") and _one(db, "Min(L)") == 10
    assert _one(db, "IIf(True, L, D)") == 10.0 and isinstance(_one(db, "IIf(True, L, D)"), float)
    assert _one(db, "IIf(True, C, L)") == Decimal("2.5000")
    assert _one(db, "IIf(True, 5.5, 1)") == 5.5 and isinstance(_one(db, "IIf(True, 5.5, 1)"), float)
    assert _one(db, "IIf(True, L, 'x')") == "10"


def test_a_variance_over_currency_is_accumulated_in_currency(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "var.accdb")
    db.create_table("V", [ColumnSpec("Id", "long"), ColumnSpec("C", "currency")])
    for i, c in enumerate(("0.125", "0.25", "0.375"), start=1):
        db.execute(f"INSERT INTO V VALUES ({i}, {c})")
    assert db.execute("SELECT Var(C) AS A, VarP(C) AS B FROM V") == [{"A": 0.0156, "B": 0.0104}]
    rows = db.execute("SELECT Var(C / 1) AS A FROM V")
    assert rows == [{"A": 0.015625}]


def test_unknown_and_ambiguous_columns_are_errors(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    with pytest.raises(AccessError, match="no column or parameter"):
        db.execute("SELECT Nope FROM Customers")
    with pytest.raises(AccessError, match="ambiguous"):
        db.execute("SELECT Id FROM Customers INNER JOIN Orders ON Customers.Id = Orders.CustomerId")
    with pytest.raises(AccessError, match="no table or query named"):
        db.execute("SELECT * FROM Nowhere")
    with pytest.raises(AccessError, match="unknown type"):
        db.execute("CREATE TABLE X (A NOSUCHTYPE)")
    with pytest.raises(AccessError, match="must start with"):
        db.execute("SHOW TABLES")


@pytest.mark.parametrize(
    ("value", "pattern", "expected"),
    [
        ("Alpha", "a*", True),
        ("Alpha", "*PHA", True),
        ("Alpha", "A?pha", True),
        ("A1", "A#", True),
        ("Ab", "A#", False),
        ("cat", "[a-c]at", True),
        ("cat", "[!a-c]at", False),
        ("a.b", "a.b", True),
        ("axb", "a.b", False),
    ],
)
def test_like_wildcards(value: str, pattern: str, expected: bool) -> None:
    assert like_match(value, pattern) is expected


def test_parser_reads_operators_by_precedence() -> None:
    def env(name: str, qualifier: str | None) -> object:
        return None

    assert Parser.parse("1 + 2 * 3 = 7 AND NOT 'a' & 'b' = 'ab' OR Len('xy') = 2").eval(env) is True
    assert Parser.parse("Null = 1").eval(env) is None
    assert Parser.parse("2 > 1 AND Null").eval(env) is None
    assert Parser.parse("1 > 2 AND Null").eval(env) is False
    assert Parser.parse("1 > 2 OR Null").eval(env) is None
    assert Parser.parse("-(3 - 5)").eval(env) == 2
    assert Parser.parse("#1/2/2024# > #1/1/2024#").eval(env) is True
    assert Parser.parse("'it''s'").eval(env) == "it's"


def test_subqueries_in_a_where_clause(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert _names(db.execute("SELECT Name FROM Customers WHERE Id IN (SELECT CustomerId FROM Orders) ORDER BY Name")) == ["Ada", "Bob"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Id NOT IN (SELECT CustomerId FROM Orders) ORDER BY Name")) == ["Cy", "Dee"]
    assert _names(db.execute("SELECT Name FROM Customers WHERE Balance = (SELECT Max(Balance) FROM Customers)")) == ["Ada"]
    # Correlated, both ways round.
    exists = "SELECT Name FROM Customers AS c WHERE EXISTS (SELECT 1 FROM Orders AS o WHERE o.CustomerId = c.Id) ORDER BY Name"
    assert _names(db.execute(exists)) == ["Ada", "Bob"]
    assert _names(db.execute(exists.replace("WHERE EXISTS", "WHERE NOT EXISTS"))) == ["Cy", "Dee"]


def test_a_subquery_as_a_value_and_as_a_table(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    rows = db.execute("SELECT Name, (SELECT Sum(Amount) FROM Orders AS o WHERE o.CustomerId = c.Id) AS Spent FROM Customers AS c ORDER BY Name")
    assert rows == [
        {"Name": "Ada", "Spent": 350.25},
        {"Name": "Bob", "Spent": 5.0},
        {"Name": "Cy", "Spent": None},
        {"Name": "Dee", "Spent": None},
    ]
    derived = db.execute("SELECT t.City, t.N FROM (SELECT City, Count(*) AS N FROM Customers GROUP BY City) AS t WHERE t.N > 1")
    assert derived == [{"City": "London", "N": 2}]
    joined = db.execute(
        "SELECT c.Name FROM Customers AS c INNER JOIN (SELECT CustomerId, Sum(Amount) AS S FROM Orders GROUP BY CustomerId) AS o "
        "ON c.Id = o.CustomerId WHERE o.S > 100"
    )
    assert joined == [{"Name": "Ada"}]
    with pytest.raises(AccessError, match="more than one row"):
        db.execute("SELECT Name FROM Customers WHERE Id = (SELECT Id FROM Customers)")


def test_a_saved_query_can_be_selected_from(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    db.create_query("Londoners", "SELECT Customers.Name, Customers.Balance FROM Customers WHERE Customers.City = 'London'")
    assert _names(db.execute("SELECT Name FROM Londoners ORDER BY Name")) == ["Ada", "Dee"]
    assert db.execute("SELECT Count(*) AS N FROM Londoners") == [{"N": 2}]


def test_union_keeps_or_drops_duplicates(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    both = db.execute("SELECT City FROM Customers UNION SELECT Name FROM Customers ORDER BY City")
    assert isinstance(both, list)
    assert [r["City"] for r in both] == [None, "Ada", "Bob", "Cy", "Dee", "London", "Paris"]
    everything = db.execute("SELECT City FROM Customers UNION ALL SELECT City FROM Customers")
    assert isinstance(everything, list) and len(everything) == 8
    assert db.execute("SELECT City FROM Customers WHERE Id = 1 UNION SELECT City FROM Customers WHERE Id = 4") == [{"City": "London"}]


def test_dml_takes_a_subquery(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    assert db.execute("UPDATE Customers SET City = 'Gone' WHERE Id NOT IN (SELECT CustomerId FROM Orders)") == 2
    assert _names(db.execute("SELECT Name FROM Customers WHERE City = 'Gone' ORDER BY Name")) == ["Cy", "Dee"]
    assert db.execute("DELETE FROM Orders WHERE CustomerId IN (SELECT Id FROM Customers WHERE Name = 'Bob')") == 1
    assert db.execute("SELECT Count(*) AS N FROM Orders") == [{"N": 3}]


def _sales(tmp_path: Path) -> AccessDatabase:
    db = AccessDatabase.create_new(tmp_path / "sales.accdb")
    db.execute("CREATE TABLE Sales (Id AUTOINCREMENT CONSTRAINT PK PRIMARY KEY, Region TEXT(20), Quarter TEXT(10), Amount CURRENCY)")
    for region, quarter, amount in (("North", "Q1", 10), ("North", "Q2", 20), ("South", "Q1", 30), ("South", "Q1", 5), ("East", "Q3", 7)):
        db.execute(f"INSERT INTO Sales (Region, Quarter, Amount) VALUES ('{region}', '{quarter}', {amount})")
    return db


def test_a_crosstab_pivots_its_rows(tmp_path: Path) -> None:
    db = _sales(tmp_path)
    rows = db.execute("TRANSFORM Sum(Amount) AS Total SELECT Region FROM Sales GROUP BY Region PIVOT Quarter")
    assert rows == [
        {"Region": "East", "Q1": None, "Q2": None, "Q3": Decimal("7.0000")},
        {"Region": "North", "Q1": Decimal("10.0000"), "Q2": Decimal("20.0000"), "Q3": None},
        {"Region": "South", "Q1": Decimal("35.0000"), "Q2": None, "Q3": None},
    ]


def test_a_crosstab_takes_an_in_list_and_an_aggregate_heading(tmp_path: Path) -> None:
    db = _sales(tmp_path)
    rows = db.execute(
        "TRANSFORM Count(*) AS N SELECT Region, Sum(Amount) AS Total FROM Sales GROUP BY Region "
        "ORDER BY Region DESC PIVOT Quarter IN ('Q1', 'Q2', 'Q4')"
    )
    assert rows == [
        {"Region": "South", "Total": Decimal("35.0000"), "Q1": 2, "Q2": None, "Q4": None},
        {"Region": "North", "Total": Decimal("30.0000"), "Q1": 1, "Q2": 1, "Q4": None},
        {"Region": "East", "Total": Decimal("7.0000"), "Q1": None, "Q2": None, "Q4": None},
    ]
    with pytest.raises(AccessError, match="no HAVING"):
        db.execute("TRANSFORM Sum(Amount) SELECT Region FROM Sales GROUP BY Region HAVING Sum(Amount) > 1 PIVOT Quarter")
    with pytest.raises(AccessError, match="needs an aggregate"):
        db.execute("TRANSFORM Amount SELECT Region FROM Sales GROUP BY Region PIVOT Quarter")


def test_jet_operators() -> None:
    def env(name: str, qualifier: str | None) -> object:
        return None

    assert Parser.parse("7 Mod 3").eval(env) == 1
    assert Parser.parse("-7 Mod 3").eval(env) == -1
    assert Parser.parse(r"7 \ 2").eval(env) == 3
    assert Parser.parse(r"-7 \ 2").eval(env) == -3
    assert Parser.parse("2 ^ 10").eval(env) == 1024.0
    # VBA's order: * binds tighter than Mod, which binds tighter than +.
    assert Parser.parse("1 + 7 Mod 3 * 2").eval(env) == 2
    assert Parser.parse("(1 + 7) Mod 3").eval(env) == 2


def test_a_transaction_undoes_everything_or_nothing(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    before = db.to_bytes()
    with pytest.raises(RuntimeError), db.transaction():
        db.execute("DELETE FROM Orders")
        db.execute("INSERT INTO Customers (Name) VALUES ('Eve')")
        db.execute("CREATE TABLE Extra (A LONG)")
        raise RuntimeError("undo")
    assert db.to_bytes() == before
    assert db.table_names() == ["Customers", "Orders"]
    assert db.table("Orders").row_count == 4
    assert _names(db.execute("SELECT Name FROM Customers ORDER BY Id"))[-1] == "Dee"
    with db.transaction():
        db.execute("INSERT INTO Customers (Name) VALUES ('Eve')")
    assert _names(db.execute("SELECT Name FROM Customers ORDER BY Id"))[-1] == "Eve"


def _rows(count: int) -> AccessDatabase:
    """A table of ``count`` numbered rows, for the statement shapes."""
    db = AccessDatabase(TEMPLATE)
    db.execute("CREATE TABLE T (N LONG, T TEXT(20))")
    for i in range(1, count + 1):
        db.execute(f"INSERT INTO T (N, T) VALUES ({i}, 'row {i}')")
    return db


def test_top_with_percent_takes_a_share_of_the_rows() -> None:
    """``TOP n PERCENT`` keeps that share, rounded up (measured: one
    percent of four rows is one)."""
    db = _rows(6)
    assert db.execute("SELECT TOP 50 PERCENT T.N FROM T ORDER BY T.N DESC") == [{"N": 6}, {"N": 5}, {"N": 4}]
    assert db.execute("SELECT TOP 1 PERCENT T.N FROM T ORDER BY T.N DESC") == [{"N": 6}]


def test_a_quantified_comparison_reads_every_row_or_any_of_them() -> None:
    db = _rows(5)
    assert db.execute("SELECT T.N FROM T WHERE T.N > ALL (SELECT N FROM T WHERE N <= 3) ORDER BY T.N") == [{"N": 4}, {"N": 5}]
    assert db.execute("SELECT T.N FROM T WHERE T.N > ANY (SELECT N FROM T WHERE N <= 3) ORDER BY T.N") == [
        {"N": 2},
        {"N": 3},
        {"N": 4},
        {"N": 5},
    ]
    assert db.execute("SELECT T.N FROM T WHERE T.N < SOME (SELECT N FROM T WHERE N = 2)") == [{"N": 1}]
    # ALL over nothing holds, ANY over nothing does not.
    everything = db.execute("SELECT T.N FROM T WHERE T.N > ALL (SELECT N FROM T WHERE N > 99)")
    assert isinstance(everything, list) and len(everything) == 5
    assert db.execute("SELECT T.N FROM T WHERE T.N > ANY (SELECT N FROM T WHERE N > 99)") == []


def test_order_by_a_number_names_that_column() -> None:
    db = _rows(3)
    assert db.execute("SELECT T.N, T.T FROM T ORDER BY 1 DESC") == [
        {"N": 3, "T": "row 3"},
        {"N": 2, "T": "row 2"},
        {"N": 1, "T": "row 1"},
    ]
    with pytest.raises(AccessError, match="ORDER BY 7 names no column"):
        db.execute("SELECT T.N FROM T ORDER BY 7")


def test_a_name_two_sources_share_is_qualified() -> None:
    db = _rows(2)
    rows = db.execute("SELECT a.N, b.N FROM T AS a INNER JOIN T AS b ON a.N = b.N ORDER BY a.N")
    assert rows == [{"a.N": 1, "b.N": 1}, {"a.N": 2, "b.N": 2}]
    # One of them alone keeps the plain name.
    assert db.execute("SELECT a.N FROM T AS a ORDER BY a.N") == [{"N": 1}, {"N": 2}]


def names_of(db: AccessDatabase, sql: str) -> list[str]:
    rows = db.execute(sql)
    assert isinstance(rows, list) and rows
    first = rows[0]
    assert isinstance(first, dict)
    return list(first)


@pytest.mark.parametrize(
    ("sql", "expected"),
    [
        # A column with no name of its own is named for where it sits.
        ("SELECT Id + 1, Balance * 2 FROM Customers", ["Expr1000", "Expr1001"]),
        ("SELECT Name, Id + 1 FROM Customers", ["Name", "Expr1001"]),
        ("SELECT Name, City, Id + 1, Balance * 2 FROM Customers",
         ["Name", "City", "Expr1002", "Expr1003"]),
        # A name the list holds twice is kept by the last one only.
        ("SELECT Id, Id FROM Customers", ["Expr1000", "Id"]),
        ("SELECT Id, Id, Id FROM Customers", ["Expr1000", "Expr1001", "Id"]),
        ("SELECT Id, Name, Id FROM Customers", ["Expr1000", "Name", "Id"]),
        ("SELECT Id, Id, Name, Name FROM Customers",
         ["Expr1000", "Id", "Expr1002", "Name"]),
        ("SELECT Id + 0 AS E, Id, Id FROM Customers", ["E", "Expr1001", "Id"]),
    ],
)
def test_output_columns_are_named_as_the_engine_names_them(
    tmp_path: Path, sql: str, expected: list[str]
) -> None:
    assert names_of(_shop(tmp_path), sql) == expected


def test_a_repeated_column_does_not_swallow_its_twin(tmp_path: Path) -> None:
    """Rows come back keyed by name, so before this the two columns
    collapsed into one and the row was a column short -- which is what
    made `INSERT INTO t (a, b) SELECT x, x FROM u` refuse to run."""
    rows = _shop(tmp_path).execute("SELECT Id, Id, Name FROM Customers ORDER BY Id")
    assert isinstance(rows, list)
    first = rows[0]
    assert isinstance(first, dict)
    assert len(first) == 3
    assert first["Expr1000"] == first["Id"]


def test_an_insert_from_a_select_can_repeat_a_column(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    db.execute("CREATE TABLE Pairs (Id LONG CONSTRAINT PK PRIMARY KEY, Same LONG, Other LONG)")

    customers = db.execute("SELECT Id FROM Customers")
    assert isinstance(customers, list)

    assert db.execute(
        "INSERT INTO Pairs (Id, Same, Other) SELECT Id, Id, Id FROM Customers"
    ) == len(customers)

    rows = db.execute("SELECT * FROM Pairs ORDER BY Id")
    assert isinstance(rows, list)
    assert all(r["Id"] == r["Same"] == r["Other"] for r in rows)


def test_two_output_columns_cannot_share_a_name(tmp_path: Path) -> None:
    """The engine refuses rather than dropping one, and so does this."""
    with pytest.raises(AccessError, match="duplicate output alias"):
        _shop(tmp_path).execute("SELECT Id AS X, Name AS X FROM Customers")


def test_a_column_two_sources_share_is_named_for_its_table(tmp_path: Path) -> None:
    """Qualifying makes the two distinct, so neither is renumbered."""
    db = _shop(tmp_path)
    db.execute("CREATE TABLE Notes (Id LONG CONSTRAINT PK2 PRIMARY KEY)")
    db.execute("INSERT INTO Notes (Id) VALUES (1)")

    assert names_of(db, "SELECT c.Id, n.Id FROM Customers AS c, Notes AS n") == ["c.Id", "n.Id"]


@pytest.mark.parametrize(
    "expression",
    ["Balance / 0", "Id / 0", "Id \\ 0", "Id Mod 0", "1 / 0", "0 / 0"],
)
def test_dividing_by_zero_answers_null(tmp_path: Path, expression: str) -> None:
    """A query that meets a zero goes on and answers Null for that row --
    it does not stop -- which is what the engine does."""
    rows = _shop(tmp_path).execute(f"SELECT Id, {expression} AS D FROM Customers ORDER BY Id")
    assert isinstance(rows, list) and rows
    assert all(row["D"] is None for row in rows)


def test_a_zero_in_one_row_does_not_stop_the_query(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    db.execute("CREATE TABLE Ratios (Id LONG CONSTRAINT PK PRIMARY KEY, Over LONG, Under LONG)")
    for number, (over, under) in enumerate([(10, 2), (10, 0), (9, 3)], start=1):
        db.execute(f"INSERT INTO Ratios (Id, Over, Under) VALUES ({number}, {over}, {under})")

    rows = db.execute("SELECT Id, Over / Under AS R FROM Ratios ORDER BY Id")

    assert isinstance(rows, list)
    assert [row["R"] for row in rows] == [5.0, None, 3.0]


@pytest.mark.parametrize(
    ("expression", "message"),
    [("Sqr(-1)", "not negative"), ("Log(0)", "above zero"), ("Log(-1)", "above zero")],
)
def test_a_maths_function_given_the_wrong_number_says_so(
    tmp_path: Path, expression: str, message: str
) -> None:
    """The engine refuses these too; what matters here is that the refusal
    is this package's own error and not the maths library's."""
    with pytest.raises(AccessError, match=message):
        _shop(tmp_path).execute(f"SELECT {expression} AS X FROM Customers")


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        # Text on one side and a number on the other is an addition.
        ("'5' + 5", 10),
        ("5 + '5'", 10),
        ("' 5 ' + 1", 6),
        ("'5.5' + 1", 6.5),
        # Text that will not read as a number makes the whole thing Null.
        ("'a' + 5", None),
        ("'' + 1", None),
        # Two strings join, and `&` always joins.
        ("'5' + '5'", "55"),
        ("'a' + 'b'", "ab"),
        ("'5' & 5", "55"),
    ],
)
def test_plus_adds_a_number_written_as_text(
    tmp_path: Path, expression: str, expected: object
) -> None:
    rows = _shop(tmp_path).execute(f"SELECT {expression} AS A FROM Customers WHERE Id = 1")
    assert isinstance(rows, list) and rows
    assert rows[0]["A"] == expected


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        # A query ignores case unless the comparison argument says not to.
        ("Replace('abcABC', 'b', 'X')", "aXcAXC"),
        ("Replace('abcABC', 'B', 'X')", "aXcAXC"),
        ("Replace('abcABC', 'b', 'X', 1, -1, 1)", "aXcAXC"),
        ("Replace('abcABC', 'b', 'X', 1, -1, 0)", "aXcABC"),
        # `start` says where the answer begins, not just where to look.
        ("Replace('abcabc', 'b', 'X', 3)", "caXc"),
        ("Replace('abcabc', 'b', 'X', 1, 1)", "aXcabc"),
        ("Replace('abc', 'z', 'X')", "abc"),
        ("InStr('abcABC', 'B')", 2),
        ("InStr(1, 'abcABC', 'B', 0)", 5),
        ("StrComp('a', 'A')", 0),
        ("StrComp('a', 'A', 0)", 1),
    ],
)
def test_the_text_functions_honour_their_comparison_argument(
    tmp_path: Path, expression: str, expected: object
) -> None:
    rows = _shop(tmp_path).execute(f"SELECT {expression} AS A FROM Customers WHERE Id = 1")
    assert isinstance(rows, list) and rows
    assert rows[0]["A"] == expected


def test_replace_starting_before_the_string_is_refused(tmp_path: Path) -> None:
    with pytest.raises(AccessError, match="starts at 1"):
        _shop(tmp_path).execute("SELECT Replace('abc', 'b', 'X', 0) AS A FROM Customers")


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        # Half goes to the even digit, and on the decimal as written --
        # the double nearest 2.345 sits above it and the one nearest 2.675
        # below, so rounding the double would answer 2.35 and 2.67.
        ("Round(2.345, 2)", 2.34),
        ("Round(2.355, 2)", 2.36),
        ("Round(2.675, 2)", 2.68),
        ("Round(-2.345, 2)", -2.34),
        ("Round(1.005, 2)", 1.0),
        ("Round(0.125, 2)", 0.12),
        ("Round(0.135, 2)", 0.14),
        ("Round(2.5)", 2.0),
        ("Round(3.5)", 4.0),
    ],
)
def test_round_goes_half_to_even_on_the_decimal_as_written(
    tmp_path: Path, expression: str, expected: float
) -> None:
    rows = _shop(tmp_path).execute(f"SELECT {expression} AS A FROM Customers WHERE Id = 1")
    assert isinstance(rows, list) and rows
    assert rows[0]["A"] == expected


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        ("CDbl('1e3')", 1000.0),
        ("CDbl('1E-2')", 0.01),
        ("CLng('1e3')", 1000),
        ("'1e3' + 1", 1001.0),
        ("CStr(True)", "-1"),
        ("CStr(False)", "0"),
        ("CStr(1.5)", "1.5"),
    ],
)
def test_numbers_written_as_text_read_the_way_the_engine_reads_them(
    tmp_path: Path, expression: str, expected: object
) -> None:
    rows = _shop(tmp_path).execute(f"SELECT {expression} AS A FROM Customers WHERE Id = 1")
    assert isinstance(rows, list) and rows
    assert rows[0]["A"] == expected


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        ("True & ''", "-1"),
        ("False & ''", "0"),
        ("CStr(True)", "-1"),
        ("'x' & True", "x-1"),
        ("True + 1", 0),
        ("-True", 1),
        ("Abs(True)", 1),
    ],
)
def test_a_boolean_is_written_as_the_number_it_is(
    tmp_path: Path, expression: str, expected: object
) -> None:
    """A query writes True as -1, not as its name."""
    rows = _shop(tmp_path).execute(f"SELECT {expression} AS A FROM Customers WHERE Id = 1")
    assert isinstance(rows, list) and rows
    assert rows[0]["A"] == expected


def test_ordering_by_an_alias_matches_ordering_by_what_it_names(tmp_path: Path) -> None:
    """DAO will not take an alias there, so the engine cannot answer for
    the alias form directly; what it can answer for is the expression,
    and the two have to agree."""
    db = _shop(tmp_path)
    by_alias = db.execute("SELECT Id, Balance * 2 AS Doubled FROM Customers ORDER BY Doubled, Id")
    by_expression = db.execute("SELECT Id, Balance * 2 AS Doubled FROM Customers ORDER BY Balance * 2, Id")

    assert by_alias == by_expression


def test_grouping_by_an_expression_orders_by_it_too(tmp_path: Path) -> None:
    db = _shop(tmp_path)
    by_alias = db.execute(
        "SELECT Len(Name) AS L, Count(*) AS N FROM Customers GROUP BY Len(Name) ORDER BY L"
    )
    by_expression = db.execute(
        "SELECT Len(Name) AS L, Count(*) AS N FROM Customers GROUP BY Len(Name) ORDER BY Len(Name)"
    )

    assert by_alias == by_expression


def money_table(tmp_path: Path) -> AccessDatabase:
    """A Currency column and a Decimal one holding the same numbers, which
    is what tells the two rules apart."""
    db = AccessDatabase.create_new(tmp_path / "money.accdb")
    db.create_table(
        "M",
        [
            ColumnSpec("Id", "Long"),
            ColumnSpec("Money4", "Currency"),
            ColumnSpec("Fix4", "Decimal", size=(18, 4)),
            ColumnSpec("Fix2", "Decimal", size=(18, 2)),
            ColumnSpec("L", "Long"),
        ],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
    )
    rows = [
        ("0.25", "0.25", "0.25", 3),
        ("1.0", "1.0", "1.0", 7),
        ("12345.6789", "12345.6789", "12345.67", 2),
    ]
    for number, (money, four, two, whole) in enumerate(rows, start=1):
        db.execute(
            f"INSERT INTO M (Id, Money4, Fix4, Fix2, L) "
            f"VALUES ({number}, {money}, {four}, {two}, {whole})"
        )
    return db


def test_averaging_a_currency_column_rounds_to_the_places_it_holds(tmp_path: Path) -> None:
    rows = money_table(tmp_path).execute("SELECT Avg(Money4) AS A FROM M")
    assert isinstance(rows, list)
    assert str(rows[0]["A"]) == "4115.6430"


def test_averaging_a_decimal_column_keeps_what_the_division_gives(tmp_path: Path) -> None:
    """Both columns hold the same numbers and arrive as `Decimal`, so only
    the column says which rule applies.  The engine's decimal is 96 bits
    wide and carries every digit that fits."""
    rows = money_table(tmp_path).execute("SELECT Avg(Fix4) AS A FROM M")
    assert isinstance(rows, list)
    assert str(rows[0]["A"]) == "4115.6429666666666666666666667"


def test_an_exact_average_keeps_only_the_places_it_needs(tmp_path: Path) -> None:
    rows = money_table(tmp_path).execute("SELECT Avg(Fix2) AS A FROM M")
    assert isinstance(rows, list)
    assert str(rows[0]["A"]) == "4115.64"


def test_averaging_the_other_numbers_is_unchanged(tmp_path: Path) -> None:
    db = money_table(tmp_path)
    rows = db.execute("SELECT Avg(L) AS A, Sum(Money4) AS S, Min(Fix4) AS N FROM M")
    assert isinstance(rows, list)
    assert rows[0]["A"] == 4.0
    assert str(rows[0]["S"]) == "12346.9289"
    assert str(rows[0]["N"]) == "0.2500"


def test_an_average_follows_the_type_the_expression_has(tmp_path: Path) -> None:
    """Currency plus a whole number is still Currency, so its average
    rounds to four places; divided, it is a Double and keeps them all
    (measured)."""
    db = money_table(tmp_path)
    assert db.execute("SELECT Avg(Money4 + 0) AS A FROM M") == [{"A": Decimal("4115.6430")}]
    rows = db.execute("SELECT Avg(Money4 / 1) AS A FROM M")
    assert isinstance(rows, list) and isinstance(rows[0]["A"], float)
