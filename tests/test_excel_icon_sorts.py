"""Sorting on the icon a conditional format's icon set shows, as the Sort object a recorded macro uses does it.

tests/fixtures/icon_sorts/ holds the workbook scripts/measure_icon_sorts.py
had Excel make, a sheet for each layout of a sort on icons over a list an
icon set covers, its dumps after the sort and once Excel opened the file
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
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "icon_sorts"
RECORD: dict[str, Any] = json.loads((FIXTURES / "icon_sorts.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def conditions(path: Path) -> dict[str, list[str]]:
    """Each sheet's sortState as its ref and each condition's attributes, as a file keeps them."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            text = package.read("xl/" + targets[relationship].lstrip("/").removeprefix("xl/")).decode("utf-8")
            out[name] = held(re.findall(r"<sortState\b.*?</sortState>|<sortState\b[^>]*/>", text, re.DOTALL))
    return out


def held(states: list[str]) -> list[str]:
    found: list[str] = []
    for state in states:
        ref = re.search(r'<sortState\b[^>]*?\bref="([^"]*)"', state)
        found.append(f"ref={ref.group(1) if ref else ''}")
        found += [" ".join(f"{key}={value}" for key, value in re.findall(r'(\w+)="([^"]*)"', condition))
                  for condition in re.findall(r"<sortCondition\b[^>]*/>", state)]
    return found


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its dumps after the sort and once the file was opened
    again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("icon_sorts")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{steps}", layout["steps"])
                   for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    return (dict(one.split("^", 1) for one in session.split("|") if "^" in one),
            dict(one.split("^", 1) for one in reopened.split("|") if "^" in one), folder)


def test_the_record_holds_the_file() -> None:
    saved = conditions(FIXTURES / "icon_sorts.xlsx")
    assert all(saved[layout["name"]] == held(layout["file"]) for layout in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_sort_on_icons_orders_rows_as_excel_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                   layout: dict[str, Any]) -> None:
    """The rows showing the icon go first, or last for xlDescending, the rest keeping their order; an icon no row
    shows, one of another set and a field given none move nothing. Each field reads back its icon, or Nothing."""
    assert (layout["error"], made[0][layout["name"]]) == ("", layout["dump"])


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_sort_on_icons_reads_back_from_the_file_as_excel_reads_it(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_sort_on_icons_as_excel_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                         layout: dict[str, Any]) -> None:
    """A condition sorting by icon, naming the icon's set and its place there from 0, or none."""
    assert conditions(made[2] / "icon_sorts.xlsx")[layout["name"]] == held(layout["file"])


def test_a_table_sort_on_icons_reports_itself() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Probe()\nDim t As Object\nActiveSheet.Range("A1:A3").Value = 1\n'
                   'Set t = ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.Range("A1:A3"), , xlYes)\n'
                   't.Sort.SortFields.Add Key:=t.ListColumns(1).Range, SortOn:=xlSortOnIcon\nEnd Sub\n',
                   name="Probe")
    with pytest.raises(VBAUnsupportedError, match="icons"):
        app.run("Probe")
