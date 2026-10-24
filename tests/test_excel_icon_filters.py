"""Which icon an icon set shows each cell, as AutoFilter by icon finds it, and how such a filter reads back and saves.

tests/fixtures/icon_filters/ holds the workbook
scripts/measure_icon_filters.py had Excel make, a sheet for each icon set
over numbers, a word, a blank and an error, filtered by each icon in turn
and by none, its dumps after each filter and once Excel opened the file
again.
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

FIXTURES = Path(__file__).parent / "fixtures" / "icon_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "icon_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
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
    """Every layout made in the model by the macro Excel ran, its dumps after each filter and once the file was
    opened again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("icon_filters")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{set}", str(layout["set"]))
                   .replace("{count}", str(layout["count"])).replace("{filter_set}", str(layout["filter_set"]))
                   .replace("{last}", str(layout["last"])) for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{cases}", RECORD["cases"]).replace("{layouts}", body).replace(
        "{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    return (dict(one.split("^", 1) for one in session.split("|") if "^" in one),
            dict(one.split("^", 1) for one in reopened.split("|") if "^" in one), folder)


def test_the_record_holds_the_file() -> None:
    held = filters(FIXTURES / "icon_filters.xlsx")
    assert all(held[layout["name"]] == [_UID.sub("", one) for one in layout["file"]] for layout in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_each_cell_shows_the_icon_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                              layout: dict[str, Any]) -> None:
    """A number shows the icon of the highest criterion it reaches, at or above its value, or above it where the
    criterion says so, the criteria standing among the range's numbers as a colour scale's points do; the order
    turned round turns the icons round, and a place given no icon shows none, as a word, a blank and an error do.
    Each filter reads back xlFilterIcon and its Icon, or xlFilterNoIcon and Empty."""
    assert made[0][layout["name"]] == layout["dumps"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_an_icon_filter_reads_back_from_the_file_as_excel_reads_it(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_holds_each_icon_filter_as_excel_saved_it(made: tuple[dict[str, str], dict[str, str], Path],
                                                           layout: dict[str, Any]) -> None:
    """The set the icon is from and its place there, from 0."""
    assert filters(made[2] / "icon_filters.xlsx")[layout["name"]] == [_UID.sub("", one) for one in layout["file"]]
