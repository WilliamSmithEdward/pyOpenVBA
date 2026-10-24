"""AutoFilter by a cell's fill or font colour, as Excel filters, reads back, saves and opens it again.

tests/fixtures/colour_filters/ holds the workbook
scripts/measure_colour_filters.py had Excel make, a sheet for each layout
of filters by colour, its dumps before the save and once Excel opened the
file again, and Excel's save of it and of the file opened again.
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

FIXTURES = Path(__file__).parent / "fixtures" / "colour_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "colour_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: The revision id a filter's file gives it, which Excel makes anew for each.
_UID = re.compile(r' xr:uid="\{[0-9A-F-]{36}\}"')


def filters(path: Path) -> dict[str, object]:
    """Each sheet's filter as the file keeps it, its colour filters holding the dxf they name, and the dxfs, as the
    record keeps them."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        styles = package.read("xl/styles.xml").decode("utf-8")
        found = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
        dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", found.group(1) or "", re.DOTALL) if found else []
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        sheets: dict[str, object] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            own = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
            tables = [package.read("xl/tables/" + target.rsplit("/", 1)[-1]).decode("utf-8")
                      for target in re.findall(r'Target="([^"]*tables/[^"]+)"', package.read(own).decode("utf-8"))] \
                if own in names else []
            elements = re.findall(r"<autoFilter\b[^>]*?/>|<autoFilter\b.*?</autoFilter>", "".join([text, *tables]),
                                  re.DOTALL)
            sheets[name] = [re.sub(r'dxfId="(\d+)"', lambda match: f'dxf="{dxfs[int(match.group(1))]}"', one)
                            for one in elements]
    return {"dxfs": dxfs, "sheets": sheets}


def masked(held: dict[str, Any]) -> dict[str, Any]:
    return {"dxfs": held["dxfs"],
            "sheets": {name: [_UID.sub("", one) for one in elements] for name, elements in held["sheets"].items()}}


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its dumps before the save and once the file was
    opened again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("colour_filters")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{table}", RECORD["table"])
                   .replace("{steps}", layout["steps"]) for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    return (dict(one.split("^", 1) for one in session.split("|") if "^" in one),
            dict(one.split("^", 1) for one in reopened.split("|") if "^" in one), folder)


def test_the_record_holds_each_file() -> None:
    assert filters(FIXTURES / "colour_filters.xlsx") == RECORD["file"]
    assert filters(FIXTURES / "colour_filters_resaved.xlsx") == RECORD["resaved"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_colour_filter_shows_the_rows_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                    layout: dict[str, Any]) -> None:
    """A row shows where its cell has a solid fill of the colour, a conditional format's or a theme colour's
    included, or no fill at all; or a font set outright in it, or the automatic one. Each filter reads back its
    operator, one criterion, and the colour as an Interior or a font's as a number."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_colour_filter_reads_back_from_the_file_as_excel_reads_it(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("name", ["colour_filters", "colour_filters_resaved"])
def test_the_file_holds_each_filter_and_its_dxf_as_excel_saved_them(
        made: tuple[dict[str, str], dict[str, str], Path], name: str) -> None:
    """Each colour filter keeps its colour in a dxf of its own, made as the filter was, so the dxfs stand among the
    rules' in the order they were made; opened again, they stand in the order Excel reads the parts."""
    assert masked(filters(made[2] / f"{name}.xlsx")) == masked(RECORD["file" if name == "colour_filters" else "resaved"])
