"""A deleted sheet goes from the file, with all it alone holds (scripts/measure_deleted_sheet_parts.py)."""

from __future__ import annotations

import importlib.util
import json
import os
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._xml import attributes
from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "deleted_sheet_parts.xlsx"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "deleted_sheet_parts.json").read_text(encoding="utf-8"))


def _deleted(target: Path, edit: str = MEASURED["edit"]) -> None:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(edit, name="E")
    app.run("E.Edit")
    app.save(target)


def _reached(package: zipfile.ZipFile) -> set[str]:
    names = set(package.namelist())
    reached = {"[Content_Types].xml"}
    waiting = ["_rels/.rels"]
    while waiting:
        rels = waiting.pop()
        if rels not in names or rels in reached:
            continue
        reached.add(rels)
        folder = posixpath.dirname(posixpath.dirname(rels))
        for tag in re.findall(r"<Relationship\b[^>]*/>", package.read(rels).decode("utf-8")):
            found = attributes(tag)
            if found.get("TargetMode") == "External":
                continue
            target = found["Target"]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            reached.add(part)
            waiting.append(posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"))
    return reached


def test_nothing_of_the_sheet_stays_in_the_file(tmp_path: Path) -> None:
    """Excel's file holds none of the markers the sheet held, whatever part they were in; nor does the model's."""
    assert MEASURED["excel"]["markers"] == {}
    assert set(MEASURED["start"]["markers"]) == set(MEASURED["markers"])
    target = tmp_path / "saved.xlsx"
    _deleted(target)
    with zipfile.ZipFile(target) as package:
        left = {marker: [name for name in package.namelist() if marker.encode() in package.read(name)] for marker in MEASURED["markers"]}
    assert not {marker: names for marker, names in left.items() if names}


def test_the_package_holds_only_what_is_reached(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    _deleted(target)
    with zipfile.ZipFile(target) as package:
        names = set(package.namelist())
        assert names == _reached(package)
        overrides = re.findall(r'PartName="/([^"]*)"', package.read("[Content_Types].xml").decode("utf-8"))
        assert [part for part in overrides if part not in names] == []


def test_the_kinds_of_part_left_are_excels(tmp_path: Path) -> None:
    """The model keeps the numbers the parts had, where Excel numbers the sheets again."""
    def kinds(parts: list[str]) -> list[str]:
        return sorted(re.sub(r"\d+", "", part) for part in parts)

    target = tmp_path / "saved.xlsx"
    _deleted(target)
    with zipfile.ZipFile(target) as package:
        assert kinds(package.namelist()) == kinds(MEASURED["excel"]["parts"])


def test_a_pivot_cache_nothing_uses_goes(tmp_path: Path) -> None:
    assert MEASURED["excel"]["workbook"]["pivot_caches"] == []
    target = tmp_path / "saved.xlsx"
    _deleted(target)
    with zipfile.ZipFile(target) as package:
        assert b"pivotCache" not in package.read("xl/workbook.xml")
        assert b"pivotCache" not in package.read("xl/_rels/workbook.xml.rels")


def test_what_the_other_sheets_read_stays(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    _deleted(target)
    again = ExcelApplication.open(target, with_vba=False)
    again.add_module('Public Function Go() As String\n    With ActiveWorkbook\n        Go = .Worksheets.Count & "|" & .Worksheets("Keep").Range("B1").Value & "|" & .Worksheets("Keep").ChartObjects.Count & "|" & .Worksheets("Data").Range("A1").Value\n    End With\nEnd Function\n', name="E")
    assert again.run("E.Go") == "2|kept text|1|h"


def test_a_second_save_changes_nothing_more(tmp_path: Path) -> None:
    first, second = tmp_path / "first.xlsx", tmp_path / "second.xlsx"
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(MEASURED["edit"], name="E")
    app.run("E.Edit")
    app.save(first)
    app.save(second)
    with zipfile.ZipFile(first) as one, zipfile.ZipFile(second) as two:
        assert sorted(one.namelist()) == sorted(two.namelist())


def test_a_save_with_no_sheet_deleted_keeps_every_part(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    ExcelApplication.open(BOOK, with_vba=False).save(target)
    with zipfile.ZipFile(BOOK) as old, zipfile.ZipFile(target) as new:
        assert sorted(old.namelist()) == sorted(new.namelist())
        assert new.read("xl/sharedStrings.xml") == old.read("xl/sharedStrings.xml")


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_reads_the_models_save_as_its_own(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    target = tmp_path / "saved.xlsx"
    _deleted(target)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        answer = excel.run_vba(MEASURED["make"], "Reads", (str(target),), timeout=60.0, line_numbers=False)
    assert answer.value == MEASURED["excel"]["reads"]
