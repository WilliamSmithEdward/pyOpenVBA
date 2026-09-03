"""The SQL executor: SELECT, INSERT, UPDATE and DELETE over the engine."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec
from pyopenvba.access._sql import Parser, like_match
from pyopenvba.access_read import AccessError


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
    rows = _shop(tmp_path).execute("SELECT Name, Balance * 2 AS Doubled, Len(Name) FROM Customers WHERE City = 'london' ORDER BY Name DESC")
    assert rows == [
        {"Name": "Dee", "Doubled": None, "Expr1000": 3},
        {"Name": "Ada", "Doubled": Decimal("21.0000"), "Expr1000": 3},
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
    assert rows == [{"City": "London", "N": 2, "Total": Decimal("10.5000"), "Expr1000": dt.datetime(2024, 3, 1)}]
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
