"""AutoFilter by dates grouped by year, month, day, hour, minute or second, as Excel filters, reads back and saves.

tests/fixtures/date_filters/ holds the workbook
scripts/measure_date_filters.py had Excel make, a sheet for each layout
of a date filter, its dumps before the save and once Excel opened the file
again, and copies of the file with a filter written the classic way,
which Excel read and saved.
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

FIXTURES = Path(__file__).parent / "fixtures" / "date_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "date_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
CLASSIC: list[dict[str, Any]] = RECORD["classic"]
#: The revision id a filter's file gives it, which Excel makes anew for each.
_UID = re.compile(r' xr:uid="\{[0-9A-F-]{36}\}"')


def filters(path: Path) -> dict[str, list[str]]:
    """Each sheet's autoFilter element as the file keeps it, its revision id left out."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            text = package.read("xl/" + targets[relationship].lstrip("/").removeprefix("xl/")).decode("utf-8")
            out[name] = [_UID.sub("", one) for one in
                         re.findall(r"<autoFilter\b[^>]*?/>|<autoFilter\b.*?</autoFilter>", text, re.DOTALL)]
    return out


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran: the error each raised, its dumps before the save and
    once the file was opened again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("date_filters")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{table}", RECORD["table"])
                   .replace("{steps}", layout["steps"]) for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    entries = [one.split("^", 1) for one in session.split("|") if "^" in one]
    return ({name: text for name, text in entries if text.startswith("E")},
            {name: text for name, text in entries if not text.startswith("E")},
            dict(one.split("^", 1) for one in reopened.split("|") if "^" in one), folder)


def test_the_record_holds_each_file() -> None:
    held = filters(FIXTURES / "date_filters.xlsx")
    assert all(held[layout["name"]] == [_UID.sub("", one) for one in layout["file"]] for layout in LAYOUTS)
    assert all(filters(FIXTURES / f"{one['name']}_resaved.xlsx")[one["sheet"]]
               == [_UID.sub("", element) for element in one["resaved"]] for one in CLASSIC)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_date_filter_shows_the_rows_excel_shows(made: tuple[dict[str, str], dict[str, str], dict[str, str], Path],
                                                  layout: dict[str, Any]) -> None:
    """A row shows where its cell holds a date, in a date format, inside a year, month, day, hour, minute or second
    named, or one of the values an array Criteria1 names beside them; a year named without a date is error 1004.
    The filter reads back as those values do, Operator 0 and no criteria where there are none."""
    assert (made[0].get(layout["name"], ""), made[1][layout["name"]]) == (layout["error"], layout["dump"])


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_date_filter_reads_back_from_the_file_as_excel_reads_it(
        made: tuple[dict[str, str], dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[2][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_holds_each_date_filter_as_excel_saved_it(
        made: tuple[dict[str, str], dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    """In the Excel 2017 extension of the column's filter: the values first, then the dates, the latest year first
    and within one the earliest moment, a group before those inside it."""
    assert filters(made[3] / "date_filters.xlsx")[layout["name"]] == [_UID.sub("", one) for one in layout["file"]]


@pytest.mark.parametrize("case", CLASSIC, ids=lambda case: case["name"])
def test_a_classic_date_filter_reads_and_saves_as_excel_does(case: dict[str, Any], tmp_path: Path) -> None:
    """A file's filters element of dates reads back as xlFilterValues with the values beside it, shows the rows
    Excel showed, and is saved again as the file had it."""
    app = ExcelApplication()
    app.add_module(RECORD["read"], name="Reader")
    saved = tmp_path / "resaved.xlsx"
    shown = dict(one.split("^", 1) for one in str(app.run("Read", str(FIXTURES / f"{case['name']}.xlsx"),
                                                         str(saved))).split("|") if "^" in one)
    assert shown[case["sheet"]] == case["dump"]
    assert filters(saved)[case["sheet"]] == [_UID.sub("", element) for element in case["resaved"]]
