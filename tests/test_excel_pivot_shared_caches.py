"""Pivot tables sharing a cache across sheets, refreshed and set, replayed against the in-memory model.

tests/fixtures/pivot_shared_caches/ holds what scripts/measure_pivot_shared_caches.py
saw in live Excel: a workbook for each scenario, a cache with two tables on
two sheets and a third beside the first, and for some a second cache with a
table of its own, the source changed and one table, its cache or the
workbook refreshed, or a setting of the cache set on one table; every sheet
dumped with its tables' names, caches' indexes, places and cells and its
columns' widths, before the workbook was saved and after it was opened
again, the workbook as Excel saved it and its pivot parts; and, in a session
of its own, the events an application's handlers saw as a refresh laid out
tables on two sheets.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_pivot_layouts import caches, column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_shared_caches"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_shared_caches.json").read_text(encoding="utf-8"))
SCENARIOS: list[dict[str, Any]] = RECORD["scenarios"]


def dumps(text: str) -> dict[str, str]:
    pairs = (one.split("^", 1) for one in text.split("|") if "^" in one)
    return {name: dump for name, dump in pairs}


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it."""
    body = "".join(RECORD["scenario"].replace("{name}", scenario["name"]).replace("{extra}", str(scenario["extra"]))
                   .replace("{step}", str(index)).replace("{folder}", str(folder))
                   for index, scenario in enumerate(SCENARIOS, 1))
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{scenarios}", body)
            .replace('folder & "\\" & name', f'folder & "{os.sep}" & name'))
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    return app


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], Path]:
    folder = tmp_path_factory.mktemp("pivot_shared_caches")
    return dumps(str(built(folder).run("Build"))), folder


def sheets(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as package:
        return set(re.findall(r'<sheet\b[^>]*name="([^"]*)"', package.read("xl/workbook.xml").decode("utf-8")))


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_every_table_shows_what_excel_shows(made: tuple[dict[str, str], Path], scenario: dict[str, Any]) -> None:
    """Every sheet's tables, their caches' indexes, places and cells and the sheet's widths, after the scenario."""
    assert made[0][scenario["name"]].split("@@", 1)[0] == scenario["dump"]


def unordered(dump: str) -> str:
    """The dump with each sheet's tables in name order. Excel lists the two tables of sheet one in either order
    from one open to the next: the files of these scenarios name the same parts in the same relationships, and the
    one measuring session listed them Third, First for some and First, Third for others. The model lists them by
    their relationships' ids."""
    def sheet(match: re.Match[str]) -> str:
        tables = re.findall(r" \{[^{}]*\}", match.group())
        return re.sub(r" \{[^{}]*\}", "", match.group())[:-1] + "".join(sorted(tables)) + "]"

    return re.sub(r"\[[^\[\]]*\]", sheet, dump)


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_tables_read_back_from_the_file_the_model_saved(made: tuple[dict[str, str], Path],
                                                          scenario: dict[str, Any]) -> None:
    assert unordered(made[0][scenario["name"]].split("@@", 1)[1]) == unordered(scenario["reopened"])


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_file_keeps_the_caches_as_excels_does(made: tuple[dict[str, str], Path],
                                                  scenario: dict[str, Any]) -> None:
    """Each table's cache, its fields and records, the tables of one cache sharing it."""
    path = f"{scenario['name']}.xlsx"
    assert caches(made[1] / path) == caches(FIXTURES / path)


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_every_cell_and_width_is_excels(made: tuple[dict[str, str], Path], scenario: dict[str, Any]) -> None:
    path = f"{scenario['name']}.xlsx"
    names = sheets(FIXTURES / path)
    assert sheet_cells(made[1] / path, names) == sheet_cells(FIXTURES / path, names)
    assert column_widths(made[1] / path, names) == column_widths(FIXTURES / path, names)


HOWS = {"refresh_table": 1, "refresh_all": 2, "refresh_all_reordered": 3}


@pytest.mark.parametrize("how", list(HOWS))
def test_a_refresh_raises_the_events_excel_raises(how: str) -> None:
    """Each table of the cache, on whatever sheet, changes, is calculated and updates as Excel's do, and
    RefreshAll refreshes the caches from the last to the first."""
    measured = RECORD["events"]
    app = ExcelApplication()
    app.add_module("Public Log As String\n", name="EventLog")
    app.add_module(measured["sink"], name="SharedSink", kind="class")
    app.add_module(measured["watch"], name="Watcher")
    assert app.run("Watch", HOWS[how]) == measured["logs"][how]
