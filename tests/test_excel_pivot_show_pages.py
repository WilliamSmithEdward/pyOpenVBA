"""PivotTable.ShowPages, a copy of a pivot table on a sheet of its own for each item of a page field, replayed
against the in-memory model.

tests/fixtures/pivot_show_pages/ holds what scripts/measure_pivot_show_pages.py
saw in live Excel: a workbook for each scenario, a source list and a pivot
table with a page field on a sheet of its own, and ShowPages called on it by
the field's name, the field or nothing, with items hidden or chosen, beside a
second page field, with the table lower down, with formatted values or a
column field, with another workbook active, and with items whose names are
numbers, dates, blank, long, odd or taken; one scenario makes tables with no
name given beside names taken; every sheet dumped in order with each table's
name, place, page choices and cells, before the workbook was saved and after
it was opened again, each workbook as Excel saved it and each table's
definition from it; and, in a session of its own, the events an application's
handlers saw as ShowPages made three sheets.
"""
from __future__ import annotations

import json
import os
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_pivot_layouts import column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_show_pages"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_show_pages.json").read_text(encoding="utf-8"))
SCENARIOS: list[dict[str, Any]] = RECORD["scenarios"]
#: What Excel writes afresh each time: revision ids, cache numbering and refresh stamps.
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')
#: Scenarios with a sheet of several tables, which Excel lists in an order that turns round from one session to
#: the next: its own default_names.xlsx, opened in two fresh sessions, listed the report sheet's seven tables in
#: the order the measuring session read them, started one table later. The model lists them by their
#: relationships' ids.
ROTATING = {"default_names"}


def relationships(package: zipfile.ZipFile, part: str) -> list[tuple[str, str]]:
    """A part's relationships, their ids and the parts they name, in the order the file lists them."""
    folder, name = posixpath.split(part)
    rels = f"{folder}/_rels/{name}.rels"
    if rels not in package.namelist():
        return []
    found: list[tuple[str, str]] = []
    for element in re.findall(r"<Relationship\b[^>]*>", package.read(rels).decode("utf-8")):
        ident = re.search(r'\bId="([^"]+)"', element)
        target = re.search(r'\bTarget="([^"]+)"', element)
        if ident and target:
            found.append((ident.group(1), posixpath.normpath(posixpath.join(folder, target.group(1)))))
    return found


def sheets(package: zipfile.ZipFile) -> dict[str, str]:
    """Each sheet's part, by its name, in tab order."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = dict(relationships(package, "xl/workbook.xml"))
    return {name: targets[relation]
            for name, relation in re.findall(r'<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"', workbook)}


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition, by its sheet's name and the table's, as the file keeps it."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for sheet, part in sheets(package).items():
            for _, table in relationships(package, part):
                if not re.search(r"pivotTable\d+\.xml$", table):
                    continue
                text = re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(table).decode("utf-8"))
                name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
                out[f"{sheet}/{name.group(1) if name else table}"] = REVISIONS.sub("", text).replace(" />", "/>")
    return out


def table_relationships(path: Path) -> dict[str, list[tuple[str, str]]]:
    """Each sheet's relationships to its pivot tables, in the file's order: the id and the table it names."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, list[tuple[str, str]]] = {}
        for sheet, part in sheets(package).items():
            for ident, table in relationships(package, part):
                if re.search(r"pivotTable\d+\.xml$", table):
                    name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"',
                                     package.read(table).decode("utf-8"))
                    out.setdefault(sheet, []).append((ident, name.group(1) if name else table))
        return out


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it."""
    body = "".join(RECORD["scenario"].replace("{name}", scenario["name"]).replace("{names}", scenario["names"])
                   .replace("{at}", scenario["at"]).replace("{step}", str(index)).replace("{folder}", str(folder))
                   for index, scenario in enumerate(SCENARIOS, 1))
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{scenarios}", body)
            .replace('folder & "\\" & name', f'folder & "{os.sep}" & name'))
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    return app


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], Path]:
    """Every scenario run in the model, in one session as Excel ran them, and the folder its workbooks went to."""
    folder = tmp_path_factory.mktemp("pivot_show_pages")
    return dumps(str(built(folder).run("Build"))), folder


def recorded(scenario: dict[str, Any]) -> dict[str, str]:
    return {key: text.replace(" />", "/>") for key, text in scenario["files"].items()}


def test_the_record_holds_the_files() -> None:
    assert {scenario["name"]: tables(FIXTURES / f"{scenario['name']}.xlsx") for scenario in SCENARIOS} == {
        scenario["name"]: recorded(scenario) for scenario in SCENARIOS}


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_sheets_are_excels(made: tuple[dict[str, str], Path], scenario: dict[str, Any]) -> None:
    """Every sheet in order, the active one, and each table's name, place, page choices and cells, with the error
    ShowPages raises."""
    assert made[0][scenario["name"]].split("@@", 1)[0] == scenario["dump"]


def unordered(dump: str) -> tuple[str, list[str]]:
    return re.sub(r"\{[^{}]*\}", "", dump), sorted(re.findall(r"\{[^{}]*\}", dump))


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_sheets_read_back_from_the_file_the_model_saved(made: tuple[dict[str, str], Path],
                                                          scenario: dict[str, Any]) -> None:
    reopened = made[0][scenario["name"]].split("@@", 1)[1]
    if scenario["name"] in ROTATING:
        assert unordered(reopened) == unordered(scenario["reopened"])
    else:
        assert reopened == scenario["reopened"]


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_file_keeps_each_copy_as_excels_does(made: tuple[dict[str, str], Path],
                                                 scenario: dict[str, Any]) -> None:
    """Each table's definition whole: its name, its place at the top left of its sheet, the page field choosing
    the copy's item alone and everything else the table it copies has."""
    assert tables(made[1] / f"{scenario['name']}.xlsx") == recorded(scenario)


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_each_sheet_names_its_tables_as_excels_does(made: tuple[dict[str, str], Path],
                                                    scenario: dict[str, Any]) -> None:
    """The ids each sheet gives its tables, in the order Office writes a part's relationships."""
    path = f"{scenario['name']}.xlsx"
    assert table_relationships(made[1] / path) == table_relationships(FIXTURES / path)


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_every_cell_is_excels(made: tuple[dict[str, str], Path], scenario: dict[str, Any]) -> None:
    """Each sheet's cells, values and formats, the copies' pivot buttons and number formats included."""
    path = f"{scenario['name']}.xlsx"
    with zipfile.ZipFile(FIXTURES / path) as package:
        names = set(sheets(package))
    assert sheet_cells(made[1] / path, names) == sheet_cells(FIXTURES / path, names)


@pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda scenario: scenario["name"])
def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], Path], scenario: dict[str, Any]) -> None:
    """Each sheet's columns as wide as Excel fits them to its table, the copies' and the table they copy (see
    tests/test_excel_pivot_column_fit.py)."""
    path = f"{scenario['name']}.xlsx"
    names = set(table_relationships(FIXTURES / path))
    assert column_widths(made[1] / path, names) == column_widths(FIXTURES / path, names)


def test_each_sheet_raises_the_events_excel_raises() -> None:
    """NewSheet once the sheet holds its table laid out, then the sheet it went before deactivating, it
    activating, and its table's cells changing, calculated and updated, one sheet after another."""
    measured = RECORD["events"]
    app = ExcelApplication()
    app.add_module("Public Log As String\n", name="EventLog")
    app.add_module(measured["sink"], name="ShowSink", kind="class")
    app.add_module(measured["watch"], name="Watcher")
    assert app.run("Watch") == measured["log"]
