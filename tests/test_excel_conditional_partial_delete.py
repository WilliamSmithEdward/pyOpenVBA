"""FormatCondition.Delete on a rule reached through a range holding part of it, as Excel deletes it.

tests/fixtures/conditional_partial_delete/ holds a workbook for each
layout scripts/measure_conditional_partial_delete.py had Excel make: rules
over a block of numbers and one Delete, through part of a rule's range,
all of it, more than it, or a rule second in a cell's FormatConditions;
Excel's read of the rules before and after the Delete; and its save.
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

FIXTURES = Path(__file__).parent / "fixtures" / "conditional_partial_delete"
RECORD: dict[str, Any] = json.loads((FIXTURES / "conditional_partial_delete.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: A rule's id in the worksheet's Excel 2010 extension, which Excel makes anew for each rule.
_ID = re.compile(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}")


def masked(texts: list[str]) -> list[str]:
    return [_ID.sub("{id}", text) for text in texts]


def holdings(path: Path) -> dict[str, object]:
    """The sheet's conditional formats as the file keeps them, and the dxfs, as the record keeps them."""
    with zipfile.ZipFile(path) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
    groups = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", sheet, re.DOTALL)
    extension = re.findall(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", sheet, re.DOTALL)
    dxfs = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
    listed = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dxfs.group(1) or "", re.DOTALL) if dxfs else []
    return {"rules": groups, "extension": extension, "dxfs": listed}


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, what it read of each, and where it saved them."""
    folder = tmp_path_factory.mktemp("partial_delete")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{rules}", layout["rules"])
                   .replace("{delete}", layout["delete"]) for layout in LAYOUTS)
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    shown = dict(one.split("^", 1) for one in str(app.run("Build")).split("|") if "^" in one)
    return shown, folder


def test_the_record_holds_each_file() -> None:
    assert all(holdings(FIXTURES / f"{layout['name']}.xlsx") == layout["file"] for layout in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_delete_leaves_the_rules_excel_leaves(made: tuple[dict[str, str], Path], layout: dict[str, Any]) -> None:
    """Through part of a rule's range, Delete takes that part out of the rule, which keeps its priority; through all
    of it or more, the rule goes."""
    assert made[0][layout["name"]] == f"{layout['before']}~{layout['after']}"


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_save_holds_the_rules_excel_saved(made: tuple[dict[str, str], Path], layout: dict[str, Any]) -> None:
    """The rule left is saved over what is left of its range, with its priority, its formula moved to the range's
    first cell, and its dxf; the extension's rules follow, each id one Excel makes anew."""
    held = holdings(made[1] / f"{layout['name']}.xlsx")
    for slot in ("rules", "extension"):
        assert masked(held[slot]) == masked(layout["file"][slot]), slot  # type: ignore[arg-type]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_save_holds_the_dxfs_excel_saved(made: tuple[dict[str, str], Path], layout: dict[str, Any]) -> None:
    """A rule cut keeps its dxf where it was, and a rule deleted whole leaves its dxf behind
    (tests/fixtures/dxf_lifetime/)."""
    assert holdings(made[1] / f"{layout['name']}.xlsx")["dxfs"] == layout["file"]["dxfs"]
