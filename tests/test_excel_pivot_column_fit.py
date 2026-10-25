"""A pivot table's columns fitted to its cells each time it is laid out, replayed against the in-memory model.

tests/fixtures/pivot_column_fit/pivot_column_fit.json holds what
scripts/measure_pivot_column_fit.py saw in live Excel: the elements of every
built-in style that draw bold; the width Excel fits a compact label column
to, for ten texts at every depth of up to four row fields with and without
the +/- buttons and at four compact indents, and the same for a subtotal's
label; outline and tabular tables with column fields; and tables edited a
step at a time, with every column's width after each step. The widths are
pixels on a 96-DPI display.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._pivot_fit import BOLD

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "pivot_column_fit" / "pivot_column_fit.json")
                                    .read_text(encoding="utf-8"))


def test_each_style_draws_bold_what_excels_does() -> None:
    """Every built-in pivot table and table style, the elements whose font it draws bold."""
    measured = {name.casefold(): frozenset(entry["bold"]) for name, entry in RECORD["styles"].items()
                if "p" in entry["kinds"] or "t" in entry["kinds"]}
    assert BOLD == measured


def grid(totals: bool) -> dict[str, str]:
    """The widths the model fits each grid case's label column to, from the macro Excel ran."""
    app = ExcelApplication()
    app.add_module(RECORD["grid"], name="Grid")
    found: dict[str, str] = {}
    for index, text in enumerate(RECORD["texts"]):
        pairs = (one.split("=", 1) for one in str(app.run("Grid", index, text, totals)).split(";") if "=" in one)
        found.update({key: value for key, value in pairs})
    return found


def test_a_compact_label_column_fits_as_excels_does() -> None:
    """Each label 88/7 pixels in beside the +/- buttons and a depth further each, 9 pixels a step of the compact
    indent and 90/7 at least; with no buttons 9 a step; none with one row field; outer items bold to two deep."""
    found = grid(totals=False)
    assert {key: found.get(key) for key in RECORD["compact"]} == RECORD["compact"]


def test_a_subtotal_label_fits_as_excels_does() -> None:
    """A subtotal's label where its indent puts it, the buttons or none, bold where its depth's subtotal rows are."""
    found = grid(totals=True)
    assert {key: found.get(key) for key in RECORD["subtotals"]} == RECORD["subtotals"]


def scenarios(part: str) -> dict[str, str]:
    """The widths of every column after each step of the part's scenarios, run in the model."""
    table = {name: (entry["rows"], entry["columns"], entry["step"]) for name, entry in RECORD[part].items()}
    app = ExcelApplication()
    app.add_module(install(table), name="Scenarios")
    pairs = (one.split("^", 1) for one in str(app.run("Scenarios")).split("|") if "^" in one)
    return {name: widths for name, widths in pairs}


def install(table: dict[str, tuple[str, str, str]]) -> str:
    """The macro Excel ran for ``table``'s scenarios, W reading the widths again (as the measuring script makes it)."""
    steps: list[str] = []
    for index, (_rows, _columns, step) in enumerate(table.values(), 1):
        lines = [f"    Case {index}"]
        for line in step.splitlines():
            lines.append('        out = out & "w=" & Widths(ws) & " "' if line == "W" else "        " + line)
        steps.append("\n".join(lines) + "\n")
    calls = "".join(f'    out = out & Run("{name}", "{rows}", "{columns}", {index})\n'
                    for index, (name, (rows, columns, _step)) in enumerate(table.items(), 1))
    return RECORD["source"].replace("{steps}", "".join(steps)).replace("{calls}", calls)


@pytest.fixture(scope="module")
def layouts() -> dict[str, str]:
    return scenarios("layouts")


@pytest.fixture(scope="module")
def steps() -> dict[str, str]:
    return scenarios("steps")


@pytest.mark.parametrize("name", list(RECORD["layouts"]))
def test_outline_and_tabular_columns_fit_as_excels_do(layouts: dict[str, str], name: str) -> None:
    """Where the columns show their fields' captions, each run of item columns a subtotal's or the grand total's
    column closes as wide as its widest; an outer row field's column neither compact nor innermost kept."""
    assert layouts[name] == RECORD["layouts"][name]["widths"]


@pytest.mark.parametrize("name", list(RECORD["steps"]))
def test_each_step_fits_the_columns_as_excels_does(steps: dict[str, str], name: str) -> None:
    """A layout, a refresh, a caption, a style and the rest fit the columns again, narrower too and over a width
    a macro set, the table's cells alone; HasAutoFormat off, a style's flags and the +/- buttons fit nothing."""
    assert steps[name] == RECORD["steps"][name]["widths"]
