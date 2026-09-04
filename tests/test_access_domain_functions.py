"""The domain functions: DLookup and its family.

Each is a query over another table, so each runs as one.  The live gate
compares every answer with Access's own.
"""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access._sql import DOMAIN_FUNCTIONS
from pyopenvba.access_read import AccessError


@pytest.fixture
def db(tmp_path: Path) -> AccessDatabase:
    database = AccessDatabase.create_new(tmp_path / "orders.accdb")
    orders = database.create_table(
        "Orders",
        [
            ColumnSpec("Id", "Long"),
            ColumnSpec("Customer", "Text", size=40),
            ColumnSpec("Total", "Currency"),
        ],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
    )
    for number, (customer, total) in enumerate(
        [("Ada", 10), ("Bob", 20), ("Ada", 30), ("Cat", 40)], start=1
    ):
        orders.insert_row({"Id": number, "Customer": customer, "Total": total})
    return database


def rows_of(db: AccessDatabase, sql: str) -> list[dict[str, object]]:
    rows = db.execute(sql)
    assert isinstance(rows, list)
    return rows


def one(db: AccessDatabase, expression: str) -> object:
    """What a domain function answers, evaluated in a query of its own."""
    rows = rows_of(db, f"SELECT {expression} AS v FROM Orders WHERE Id = 1")
    return next(iter(rows[0].values()))


@pytest.mark.parametrize(
    ("expression", "expected"),
    [
        ("DLookup('Customer', 'Orders', 'Id = 3')", "Ada"),
        ("DLookup('Total', 'Orders', 'Id = 2')", Decimal("20.0000")),
        ("DCount('*', 'Orders')", 4),
        ('DCount("Id", "Orders", "Customer = \'Ada\'")', 2),
        ("DSum('Total', 'Orders')", Decimal("100.0000")),
        ("DAvg('Total', 'Orders')", Decimal("25.0000")),
        ("DMin('Total', 'Orders')", Decimal("10.0000")),
        ("DMax('Total', 'Orders')", Decimal("40.0000")),
        ('DMax("Total", "Orders", "Customer = \'Ada\'")', Decimal("30.0000")),
        ("DFirst('Customer', 'Orders')", "Ada"),
        ("DLast('Customer', 'Orders')", "Cat"),
    ],
)
def test_each_domain_function_answers(
    db: AccessDatabase, expression: str, expected: object
) -> None:
    assert one(db, expression) == expected


def test_nothing_found_is_null_but_nothing_counted_is_zero(db: AccessDatabase) -> None:
    assert one(db, "DLookup('Total', 'Orders', 'Id = 99')") is None
    assert one(db, "DCount('Id', 'Orders', 'Id = 99')") == 0
    assert one(db, "DSum('Total', 'Orders', 'Id = 99')") is None


def test_a_criteria_can_name_the_row_it_is_evaluated_in(db: AccessDatabase) -> None:
    """Which is what makes a domain function worth having: each row's own
    customer decides what its total is summed over."""
    rows = rows_of(
        db,
        "SELECT Customer, "
        "DSum('Total', 'Orders', 'Customer = ' & Chr(34) & Customer & Chr(34)) AS Mine "
        "FROM Orders ORDER BY Id",
    )
    assert [(r["Customer"], r["Mine"]) for r in rows] == [
        ("Ada", Decimal("40.0000")),
        ("Bob", Decimal("20.0000")),
        ("Ada", Decimal("40.0000")),
        ("Cat", Decimal("40.0000")),
    ]


def test_the_spread_functions_answer(db: AccessDatabase) -> None:
    variance = one(db, "DVar('Total', 'Orders')")
    deviation = one(db, "DStDev('Total', 'Orders')")
    assert isinstance(variance, float) and isinstance(deviation, float)
    assert round(variance, 6) == round(500 / 3, 6)
    assert round(deviation * deviation, 6) == round(variance, 6)


def test_a_domain_that_does_not_exist_is_refused(db: AccessDatabase) -> None:
    with pytest.raises(AccessError):
        one(db, "DLookup('Total', 'Nothing')")


def test_too_few_arguments_are_refused(db: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="takes an expression, a domain"):
        one(db, "DLookup('Total')")


def test_every_domain_function_is_reachable(db: AccessDatabase) -> None:
    """The table drives the dispatch, so a name in it that does not run is
    a name that lies."""
    for name in DOMAIN_FUNCTIONS:
        expression = "'*'" if name == "DCOUNT" else "'Total'"
        assert one(db, f"{name}({expression}, 'Orders')") is not None
