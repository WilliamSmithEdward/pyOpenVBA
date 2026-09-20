"""Refreshing a query in the in-memory workbook.

The evaluator is held to its own tests; this is the wiring: that a
refresh finds where the query loads to, writes the rows there, keeps
the table in the file in step with them, and that a query can read the
workbook it lives in.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._refresh import load_targets
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.powerquery._opc import OpcFile

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
LOADED = FIXTURES / "loaded_to_sheet.xlsx"


def opened(tmp_path: Path, name: str = "copy.xlsx") -> ExcelApplication:
    copy = tmp_path / name
    shutil.copyfile(LOADED, copy)
    return ExcelApplication.open(copy)


def test_the_load_target_is_found_from_the_table(tmp_path: Path) -> None:
    app = opened(tmp_path)
    targets = load_targets(app.workbook)
    assert "Loaded" in targets
    assert targets["Loaded"].sheet == "Sheet1"
    assert targets["Loaded"].area.address(absolute=False) == "A1:B3"


def test_refreshing_the_fixture_reproduces_its_rows(tmp_path: Path) -> None:
    """The fixture's own M, evaluated here, gives what is on its sheet."""
    app = opened(tmp_path)
    before = app.sheet(1).rows()
    assert app.refresh_query("Loaded") == [["A", "B"], [1.0, "x"], [2.0, "y"]]
    assert app.sheet(1).rows() == before


def test_a_new_formula_lands_on_the_sheet(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = (
        'let Source = #table({"N", "Sq"}, {{1, 1}, {2, 4}, {3, 9}}) in Source'
    )
    app.refresh_query("Loaded")
    assert app.sheet(1).rows() == [["N", "Sq"], [1, 1], [2, 4], [3, 9]]


def test_a_macro_can_refresh_through_the_object_model(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.add_module(
        "Sub Go()\n"
        '    ActiveWorkbook.Queries("Loaded").Formula = '
        '"let Source = #table({""N""}, {{7}}) in Source"\n'
        '    ActiveWorkbook.Queries("Loaded").Refresh\n'
        "End Sub\n",
        name="Module1",
    )
    app.run("Go")
    assert app.sheet(1).rows() == [["N"], [7]]


def test_the_table_in_the_file_follows_the_new_rows(tmp_path: Path) -> None:
    """A table that says it ends early shows less than the data under it."""
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = (
        'let Source = #table({"N", "Sq"}, {{1, 1}, {2, 4}, {3, 9}, {4, 16}}) in Source'
    )
    app.refresh_query("Loaded")
    out = tmp_path / "out.xlsx"
    app.save(out)

    package = OpcFile.parse(out.read_bytes())
    table = package.read("xl/tables/table1.xml").decode("utf-8")
    assert 'ref="A1:B5"' in table
    assert 'name="N"' in table and 'name="Sq"' in table
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    assert "Sheet1!$A$1:$B$5" in workbook


def test_a_shorter_result_clears_what_was_there(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = 'let Source = #table({"N"}, {{1}}) in Source'
    app.refresh_query("Loaded")
    assert app.sheet(1).rows() == [["N"], [1]]


def test_a_query_can_read_the_workbook_it_lives_in(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = (
        "let Source = Excel.CurrentWorkbook(), Names = Table.ColumnNames(Source) in "
        'Table.FromRows({Names}, {"C1", "C2"})'
    )
    assert app.refresh_query("Loaded") == [["C1", "C2"], ["Name", "Content"]]


def test_one_query_can_name_another(tmp_path: Path) -> None:
    app = opened(tmp_path)
    book = app.workbook
    book.queries_.vba_get(
        "Add", ["Numbers", 'let Source = #table({"N"}, {{1}, {2}, {3}}) in Source']
    )
    book.queries_.entries[0].formula = (
        'let Source = Table.AddColumn(Numbers, "Double", each [N] * 2) in Source'
    )
    assert app.refresh_query("Loaded") == [["N", "Double"], [1, 2], [2, 4], [3, 6]]


def test_a_query_that_names_itself_is_reported(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = "let Source = Loaded in Source"
    with pytest.raises(Exception) as raised:
        app.refresh_query("Loaded")
    assert "refers to itself" in str(raised.value)


def test_a_source_that_cannot_be_reached_says_so(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = (
        'let Source = Web.Contents("https://example.invalid/data.csv") in Source'
    )
    with pytest.raises(VBAUnsupportedError) as raised:
        app.refresh_query("Loaded")
    assert "Web.Contents" in str(raised.value)


def test_a_refreshed_workbook_reopens_with_its_rows(tmp_path: Path) -> None:
    app = opened(tmp_path)
    app.workbook.queries_.entries[0].formula = (
        'let Source = #table({"N", "Sq"}, {{1, 1}, {2, 4}}) in Source'
    )
    app.refresh_query("Loaded")
    out = tmp_path / "out.xlsx"
    app.save(out)
    assert ExcelApplication.open(out).sheet(1).rows() == [["N", "Sq"], [1, 1], [2, 4]]
