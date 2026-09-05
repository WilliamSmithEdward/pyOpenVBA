"""Compose a Power Query step by step, and edit the steps afterwards.

    python examples/power_query_steps.py [path.xlsx]

A query's Applied Steps are the top-level bindings of its ``let``, and
each one names the one before it. That chain of names is what makes a
query awkward to edit as text: insert a step in the middle and the step
after it still points at the wrong one.

So this example keeps the steps in a Python list and writes ``#PREV``
where a step means "whatever came before me". :func:`let` resolves that
token when it renders the ``let``, which leaves inserting, removing and
reordering as plain list operations with the wiring taken care of.

The workbook it writes is a small sales pipeline:

* **Orders** and **Products** are the sources, written in M so the file
  needs nothing from outside.
* **OrderLines** is the intricate one: eleven steps that join the two
  sources, price the lines, discount them by quantity, drop the ones that
  are too small, sort and number what is left.
* **ProductTotals** and **MonthlySales** group that result two ways.

Three of them are loaded onto the first sheet. Open the file and press
Refresh All. Every query here was refreshed in Excel before it was
committed.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

from pyopenvba import PowerQueryWorkbook
from pyopenvba.powerquery import quote_name

CRLF = "\r\n"
#: What a step writes where it means the step before it.
PREVIOUS = "#PREV"
#: One step: the name Excel shows in Applied Steps, and its expression.
Step = tuple[str, str]


def let(steps: list[Step], result: str | None = None) -> str:
    """An M ``let`` expression from an ordered list of steps.

    Every ``#PREV`` becomes the name of the step before, spelled the way
    M spells it, so a step called ``Added Net`` is referred to as
    ``#"Added Net"``. The last step is what the query returns unless
    `result` names another one.
    """
    if not steps:
        raise ValueError("a let expression needs at least one step")
    lines: list[str] = []
    previous: str | None = None
    for name, expression in steps:
        if PREVIOUS in expression:
            if previous is None:
                raise ValueError(f"the first step {name!r} has no step before it to reference")
            expression = expression.replace(PREVIOUS, quote_name(previous))
        lines.append(f"    {quote_name(name)} = {expression}")
        previous = name
    body = ("," + CRLF).join(lines)
    return f"let{CRLF}{body}{CRLF}in{CRLF}    {quote_name(result or steps[-1][0])}"


def after(steps: list[Step], name: str, step: Step) -> list[Step]:
    """The steps with `step` inserted after the one called `name`."""
    at = _index(steps, name)
    return [*steps[: at + 1], step, *steps[at + 1 :]]


def without(steps: list[Step], name: str) -> list[Step]:
    """The steps with the one called `name` taken out."""
    _index(steps, name)
    return [step for step in steps if step[0] != name]


def replacing(steps: list[Step], name: str, expression: str) -> list[Step]:
    """The steps with one step's expression changed."""
    at = _index(steps, name)
    return [*steps[:at], (name, expression), *steps[at + 1 :]]


def _index(steps: list[Step], name: str) -> int:
    for at, (existing, _expression) in enumerate(steps):
        if existing == name:
            return at
    known = ", ".join(existing for existing, _ in steps)
    raise ValueError(f"there is no step called {name!r}; there is: {known}")


# --- the two sources ----------------------------------------------------------

ORDERS: list[Step] = [
    (
        "Source",
        "Table.FromRecords({" + CRLF
        + '        [Id = 1,  Date = #date(2026, 1, 8),  Sku = "W-1", Qty = 12],' + CRLF
        + '        [Id = 2,  Date = #date(2026, 1, 19), Sku = "G-2", Qty = 3],' + CRLF
        + '        [Id = 3,  Date = #date(2026, 1, 27), Sku = "D-3", Qty = 25],' + CRLF
        + '        [Id = 4,  Date = #date(2026, 2, 3),  Sku = "W-1", Qty = 7],' + CRLF
        + '        [Id = 5,  Date = #date(2026, 2, 14), Sku = "G-2", Qty = 9],' + CRLF
        + '        [Id = 6,  Date = #date(2026, 2, 21), Sku = "K-4", Qty = 2],' + CRLF
        + '        [Id = 7,  Date = #date(2026, 3, 2),  Sku = "D-3", Qty = 40],' + CRLF
        + '        [Id = 8,  Date = #date(2026, 3, 11), Sku = "W-1", Qty = 5],' + CRLF
        + '        [Id = 9,  Date = #date(2026, 3, 18), Sku = "K-4", Qty = 15],' + CRLF
        + '        [Id = 10, Date = #date(2026, 3, 29), Sku = "G-2", Qty = 1]})',
    ),
    (
        "Changed Type",
        f"Table.TransformColumnTypes({PREVIOUS}, {{" + CRLF
        + '        {"Id", Int64.Type}, {"Date", type date}, {"Sku", type text}, {"Qty", Int64.Type}})',
    ),
]

PRODUCTS: list[Step] = [
    (
        "Source",
        "Table.FromRecords({" + CRLF
        + '        [Sku = "W-1", Name = "Widget",    Category = "Hardware", Price = 9.99],' + CRLF
        + '        [Sku = "G-2", Name = "Gadget",    Category = "Hardware", Price = 24.5],' + CRLF
        + '        [Sku = "D-3", Name = "Doohickey", Category = "Supplies", Price = 1.75],' + CRLF
        + '        [Sku = "K-4", Name = "Kit",       Category = "Supplies", Price = 45.0]})',
    ),
    (
        "Changed Type",
        f"Table.TransformColumnTypes({PREVIOUS}, {{" + CRLF
        + '        {"Sku", type text}, {"Name", type text},' + CRLF
        + '        {"Category", type text}, {"Price", Currency.Type}})',
    ),
]

# --- the query this example is about ------------------------------------------

#: The columns the last step keeps, and the order the sheet shows them in.
LINE_COLUMNS = ["Line", "Date", "Sku", "Product", "Category", "Qty", "Gross", "Discount", "Net"]


def keep(columns: list[str]) -> str:
    """The expression for a step that keeps these columns, in this order."""
    listed = ", ".join(f'"{column}"' for column in columns)
    return f"Table.SelectColumns({PREVIOUS}, {{{listed}}})"


ORDER_LINES: list[Step] = [
    ("Source", "Orders"),
    ("Merged Products", f'Table.NestedJoin({PREVIOUS}, {{"Sku"}}, Products, {{"Sku"}}, "Product", JoinKind.LeftOuter)'),
    (
        "Expanded Product",
        f'Table.ExpandTableColumn({PREVIOUS}, "Product",' + CRLF
        + '        {"Name", "Category", "Price"}, {"Product", "Category", "Price"})',
    ),
    ("Added Gross", f'Table.AddColumn({PREVIOUS}, "Gross", each [Qty] * [Price], Currency.Type)'),
    (
        "Added Discount",
        f'Table.AddColumn({PREVIOUS}, "Discount",' + CRLF
        + "        each if [Qty] >= 20 then 0.15 else if [Qty] >= 10 then 0.1 else 0, type number)",
    ),
    ("Added Net", f'Table.AddColumn({PREVIOUS}, "Net", each [Gross] * (1 - [Discount]), Currency.Type)'),
    ("Removed Price", f'Table.RemoveColumns({PREVIOUS}, {{"Price"}})'),
    ("Filtered Rows", f"Table.SelectRows({PREVIOUS}, each [Net] > 20)"),
    ("Sorted Rows", f'Table.Sort({PREVIOUS}, {{{{"Date", Order.Ascending}}, {{"Net", Order.Descending}}}})'),
    ("Added Line Number", f'Table.AddIndexColumn({PREVIOUS}, "Line", 1, 1, Int64.Type)'),
    ("Reordered Columns", keep(LINE_COLUMNS)),
]

PRODUCT_TOTALS: list[Step] = [
    ("Source", "OrderLines"),
    (
        "Grouped Rows",
        f'Table.Group({PREVIOUS}, {{"Category", "Product"}}, {{' + CRLF
        + '        {"Lines", each Table.RowCount(_), Int64.Type},' + CRLF
        + '        {"Units", each List.Sum([Qty]), Int64.Type},' + CRLF
        + '        {"Revenue", each List.Sum([Net]), Currency.Type}})',
    ),
    ("Sorted Rows", f'Table.Sort({PREVIOUS}, {{{{"Revenue", Order.Descending}}}})'),
]

MONTHLY_SALES: list[Step] = [
    ("Source", "OrderLines"),
    ("Added Month", f'Table.AddColumn({PREVIOUS}, "Month", each Date.StartOfMonth([Date]), type date)'),
    (
        "Grouped Rows",
        f'Table.Group({PREVIOUS}, {{"Month"}}, {{' + CRLF
        + '        {"Lines", each Table.RowCount(_), Int64.Type},' + CRLF
        + '        {"Revenue", each List.Sum([Net]), Currency.Type}})',
    ),
    ("Sorted Rows", f'Table.Sort({PREVIOUS}, {{{{"Month", Order.Ascending}}}})'),
]


def build(path: str | Path) -> Path:
    """Write the workbook at `path` and return the path."""
    with PowerQueryWorkbook.create_new(path) as book:
        sources = book.add_group("Sources")
        reports = book.add_group("Reports")

        book.add_query("Orders", let(ORDERS), group=sources, description="Ten order lines, written in M.")
        book.add_query("Products", let(PRODUCTS), group=sources, description="The price list the orders join to.")

        book.add_query(
            "OrderLines", let(ORDER_LINES), group=reports,
            description="Eleven steps: join, price, discount, trim, sort, number.",
        )
        book.add_query(
            "ProductTotals", let(PRODUCT_TOTALS), group=reports,
            description="OrderLines grouped by category and product.",
        )
        book.add_query(
            "MonthlySales", let(MONTHLY_SALES), group=reports,
            description="OrderLines grouped by the month each line falls in.",
        )

        book.load_to_sheet("OrderLines", LINE_COLUMNS, cell="A1")
        book.load_to_sheet("ProductTotals", ["Category", "Product", "Lines", "Units", "Revenue"], cell="L1")
        book.load_to_sheet("MonthlySales", ["Month", "Lines", "Revenue"], cell="S1")
        return book.save()


def add_margin(path: str | Path) -> list[str]:
    """Put a margin column into OrderLines, and drop a step on the way.

    Two list operations and one changed expression. The step after the
    new one picks it up on its own, because it asks for whatever came
    before it, and pyOpenVBA rewrites the metadata items from the new
    expression when the formula is set.
    """
    steps = after(
        ORDER_LINES,
        "Added Net",
        ("Added Margin", f'Table.AddColumn({PREVIOUS}, "Margin", each [Net] - [Gross] * 0.6, Currency.Type)'),
    )
    steps = without(steps, "Removed Price")
    steps = replacing(steps, "Reordered Columns", keep([*LINE_COLUMNS, "Margin"]))

    book = PowerQueryWorkbook(path)
    book.query("OrderLines").formula = let(steps)
    book.save()
    return PowerQueryWorkbook(path).query("OrderLines").steps


if __name__ == "__main__":
    written = build(sys.argv[1] if len(sys.argv) > 1 else "power_query_steps.xlsx")
    print(written)
    for query in PowerQueryWorkbook(written).queries():
        print(f"  {query.name:14} {len(query.steps):2} steps  {query.load_target}")

    # The edit runs on a copy, so the workbook above stays as it was built.
    copy = Path(tempfile.mkdtemp()) / "edited.xlsx"
    shutil.copyfile(written, copy)
    print("\nafter add_margin():")
    for name in add_margin(copy):
        print(f"  {name}")
