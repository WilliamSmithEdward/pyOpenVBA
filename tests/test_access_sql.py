"""The SQL executor: SELECT, INSERT, UPDATE and DELETE over the engine."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access._sql import Parser, like_match
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
    # The bytes have to come back too, not just the rows.
    assert AccessDatabase(db.to_bytes()).table("Dst").row_count == 2


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


def test_an_average_of_an_expression_is_not_treated_as_currency(tmp_path: Path) -> None:
    """The rule follows the column, so anything but a plain read of one
    keeps what the division gives."""
    rows = money_table(tmp_path).execute("SELECT Avg(Money4 + 0) AS A FROM M")
    assert isinstance(rows, list)
    assert str(rows[0]["A"]) != "4115.6430"
