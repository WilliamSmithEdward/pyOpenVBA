"""Live Access gate for the domain functions (opt-in).

`DLookup` and its family are Access's own, not the database engine's, so
DAO cannot answer for them: the oracle here has to be the application.
Each expression is evaluated twice -- once by the executor and once by
Access's `Eval` on the same database -- and the two must agree.

Opt-in: set ``RUN_LIVE_ACCESS_VBA=1`` on a Windows machine with desktop
Access and ``pyvbaharness`` installed.  ``pyvbaharness`` is a test-time
oracle only; pyOpenVBA never uses COM.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS_VBA") != "1" or sys.platform != "win32",
    reason="live Access gate: set RUN_LIVE_ACCESS_VBA=1 on Windows with Access installed",
)

_TIMEOUT = 60.0
_PROBE = """
Public Function EvalIt(ByVal expression As String) As Variant
    EvalIt = Eval(expression)
End Function
"""

#: Written as Access would write them, since Access is what evaluates the
#: right-hand side.
EXPRESSIONS = [
    'DLookup("Customer", "Orders", "Id = 3")',
    'DLookup("Total", "Orders", "Id = 2")',
    'DLookup("Total", "Orders", "Id = 99")',
    'DCount("*", "Orders")',
    "DCount(\"Id\", \"Orders\", \"Customer = 'Ada'\")",
    'DCount("Id", "Orders", "Id = 99")',
    'DSum("Total", "Orders")',
    'DAvg("Total", "Orders")',
    'DMin("Total", "Orders")',
    'DMax("Total", "Orders")',
    "DMax(\"Total\", \"Orders\", \"Customer = 'Ada'\")",
    'DFirst("Customer", "Orders")',
    'DLast("Customer", "Orders")',
    'DStDev("Total", "Orders")',
    'DStDevP("Total", "Orders")',
    'DVar("Total", "Orders")',
    'DVarP("Total", "Orders")',
]


@pytest.fixture
def orders(tmp_path: Path) -> Path:
    path = tmp_path / "orders.accdb"
    database = AccessDatabase.create_new(path)
    table = database.create_table(
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
        table.insert_row({"Id": number, "Customer": customer, "Total": total})
    database.save(path)
    return path


def answered(database: AccessDatabase, expression: str) -> object:
    """What the executor makes of one domain function."""
    rows = database.execute(f"SELECT {expression} AS v FROM Orders WHERE Id = 1")
    assert isinstance(rows, list) and rows
    row = rows[0]
    assert isinstance(row, dict)
    return next(iter(row.values()))


def agree(mine: object, theirs: object) -> bool:
    if mine is None or theirs is None:
        return mine is None and theirs is None
    try:
        return abs(float(mine) - float(theirs)) < 1e-9  # pyright: ignore[reportArgumentType]
    except (TypeError, ValueError):
        return str(mine) == str(theirs)


def test_every_domain_function_answers_what_access_answers(orders: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    database = AccessDatabase(orders)
    mine = {expression: answered(database, expression) for expression in EXPRESSIONS}

    differences: list[str] = []
    with harness.AccessSession() as access:
        access.open_document(orders, read_only=False)
        for expression in EXPRESSIONS:
            result = access.run_vba(_PROBE, proc="EvalIt", args=(expression,), timeout=_TIMEOUT)
            assert result.outcome == "passed", f"{expression}: {result.outcome}"
            if not agree(mine[expression], result.value):
                differences.append(f"{expression}: ours {mine[expression]!r}, Access {result.value!r}")
    assert not differences, "\n".join(differences)
