"""AutoFilter's dynamic filters on dates -- today, this week, last month, a quarter of any year -- as Excel has them.

tests/fixtures/dynamic_filters/ holds the workbook
scripts/measure_dynamic_filters.py had Excel make, a sheet for each of
the 32 dynamic filters on dates over rows of dates around the day it was
made, its dumps before the save and once Excel opened the file again, and
the moment it was made, where the model's clock stands for the replay.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "dynamic_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "dynamic_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: The moment Excel made the workbook, its rows counted from that day.
NOW = dt.datetime.strptime(RECORD["now"], "%Y-%m-%d %H:%M:%S")
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
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its clock where Excel's stood: its dumps before the
    save and once the file was opened again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("dynamic_filters")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{criteria}", str(layout["criteria"]))
                   for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.freeze_clock(NOW)
    app.add_module(code, name="Builder")
    head, rest = str(app.run("Build")).split("|", 1)
    assert head == "now=" + RECORD["now"]
    session, reopened = rest.split("~", 1)
    return (dict(one.split("^", 1) for one in session.split("|") if "^" in one),
            dict(one.split("^", 1) for one in reopened.split("|") if "^" in one), folder)


def test_the_record_holds_the_file() -> None:
    held = filters(FIXTURES / "dynamic_filters.xlsx")
    assert all(held[layout["name"]] == [_UID.sub("", one) for one in layout["file"]] for layout in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_dynamic_filter_shows_the_rows_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                     layout: dict[str, Any]) -> None:
    """A row shows holding a date, in a date format, in the period: a day, a week from Sunday, a month, a quarter
    or a year counted from today, the year to date, or a quarter or a month of any year. The filter reads back as
    xlFilterDynamic and its Criteria1."""
    assert (layout["error"], made[0][layout["name"]]) == ("", layout["dump"])


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_dynamic_filter_reads_back_from_the_file_as_excel_reads_it(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_holds_each_dynamic_filter_as_excel_saved_it(made: tuple[dict[str, str], dict[str, str], Path],
                                                              layout: dict[str, Any]) -> None:
    """A period counted from today with its first serial and the one after its last; a quarter or a month of any
    year by its name alone."""
    assert filters(made[2] / "dynamic_filters.xlsx")[layout["name"]] == [_UID.sub("", one) for one in layout["file"]]
