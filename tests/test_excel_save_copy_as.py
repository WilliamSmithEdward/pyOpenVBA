"""Workbook.SaveCopyAs, replayed from live Excel (scripts/measure_save_copy_as.py, tests/fixtures/save_copy_as.json).

The probe runs in the model as Excel ran it, but for what the test does in the model's stead: ChDrive and ChDir,
which the model does not have, so the test runs in the probe's folder; MkDir, which it does not do, so the test makes
the folder; and the probe's hyperlinks, whose copy into another folder Excel moves and the model refuses
(test_a_copy_elsewhere_of_a_workbook_linking_to_files_is_refused).
"""

from __future__ import annotations

import base64
import datetime as dt
import json
import os
import re
import sys
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
RECORD = json.loads((FIXTURES / "save_copy_as.json").read_text(encoding="utf-8"))
NOW = dt.datetime(2026, 10, 25, 4, 21, 32)
LEFT_OUT = ("ChDrive ", "ChDir ", "MkDir ", "Hyperlinks.Add", 'link_before^"', 'link_after^"')
#: A new workbook the model saves macro-enabled carries an empty VBA project, where Excel writes none: a matter
#: of SaveAs, which made these files and the copy of one, not of SaveCopyAs.
EMPTY_PROJECT = ("macro.xlsm", "macro_copy.xlsx")
_OVERRIDE = re.compile(r'<Override\b[^>]*PartName="/xl/workbook\.xml"[^>]*ContentType="([^"]+)"')
_CELL = re.compile(r'<c r="([AB]1)"[^>]*?(?:/>|>(.*?)</c>)', re.S)
_VALUE = re.compile(r"<v>(.*?)</v>", re.S)


def model_source(code: str) -> str:
    """Excel's probe without the lines the test does for the model, its paths in this system's separator."""
    code = "\n".join(line for line in code.split("\n") if not any(part in line for part in LEFT_OUT))
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


def portable(text: str) -> str:
    """A path in Excel's record as this system writes it."""
    return text.replace("\\", os.sep)


def written(path: Path) -> dict[str, object]:
    """What a file holds that the measurement read of Excel's: its workbook's content type, whether it has a VBA
    project, its first sheet's A1 and B1, and whether its last author is the one a macro set."""
    if not zipfile.is_zipfile(path):
        return {"zip": False}
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        types = package.read("[Content_Types].xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        core = package.read("docProps/core.xml").decode("utf-8")
    found = _OVERRIDE.search(types)
    cells: dict[str, str | None] = {}
    for match in _CELL.finditer(sheet):
        value = _VALUE.search(match.group(2) or "")
        cells[match.group(1)] = value.group(1) if value else None
    return {"zip": True, "content_type": found.group(1) if found else None, "vba": "xl/vbaProject.bin" in names,
            "cells": cells, "last_author_someone_else": "<cp:lastModifiedBy>Someone else</cp:lastModifiedBy>" in core}


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[list[str]], Path]:
    """The probe and the overwrite with alerts on, run in the model; their records and the folder they wrote in."""
    folder = tmp_path_factory.mktemp("save_copy_as").resolve()
    (folder / "sub").mkdir()
    here = Path.cwd()
    os.chdir(folder)
    try:
        app = ExcelApplication()
        app.add_workbook()
        app.freeze_clock(NOW)
        app.add_module(RECORD["log"], name="CopyLog")
        app.add_module(RECORD["sink"], name="CopySink", kind="class")
        app.add_module(model_source(RECORD["probe"]), name="CopyProbe")
        app.add_module(model_source(RECORD["loud"]), name="CopyLoud")
        answer = str(app.run("Probe", str(folder))) + str(app.run("Loud", str(folder)))
    finally:
        os.chdir(here)
    return [entry.split("^") for entry in answer.split("~") if entry], folder


def test_each_copy_leaves_the_workbook_and_raises_what_excel_raises(replayed: tuple[list[list[str]], Path]) -> None:
    records, _folder = replayed
    measured = [[portable(field) for field in case] for case in RECORD["cases"] if "link" not in case[0]]
    assert records == measured


def test_each_file_holds_what_excels_holds(replayed: tuple[list[list[str]], Path]) -> None:
    _records, folder = replayed
    files = {path.relative_to(folder).as_posix(): written(path) for path in folder.rglob("*") if path.is_file()}
    assert sorted(files) == sorted(RECORD["files"])
    for name, seen in files.items():
        measured = {key: value for key, value in RECORD["files"][name].items() if key in seen}
        if name in EMPTY_PROJECT:
            assert seen.pop("vba") and not measured.pop("vba"), name
        assert seen == measured, name


def test_a_copy_is_the_save_that_follows_it(replayed: tuple[list[list[str]], Path]) -> None:
    # Excel's two files differ only in xl/workbook.xml, where each save writes a revision id of its own at random.
    _records, folder = replayed
    assert RECORD["copy1_differs_from_save"] == ["xl/workbook.xml"]
    with zipfile.ZipFile(folder / "copy1.xlsx") as copy, zipfile.ZipFile(folder / "named.xlsx") as saved:
        assert copy.namelist() == saved.namelist()
        assert all(copy.read(name) == saved.read(name) for name in copy.namelist())


def test_a_copy_elsewhere_of_a_workbook_linking_to_files_is_refused(tmp_path: Path) -> None:
    (tmp_path / "sub").mkdir()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Made(folder As String)\n'
                   '    Worksheets(1).Hyperlinks.Add Worksheets(1).Range("A3"), folder & "\\target.txt"\n'
                   'End Sub\n'
                   'Public Sub Copied(target As String)\n'
                   '    ThisWorkbook.SaveCopyAs target\n'
                   'End Sub\n', name="Links")
    app.run("Made", str(tmp_path))
    # A workbook never saved has its links go relative to the copy's folder wherever the copy goes.
    with pytest.raises(VBAUnsupportedError, match="links to files"):
        app.run("Copied", str(tmp_path / "fresh.xlsx"))
    app.workbook.SaveAs(str(tmp_path / "named.xlsx"), 51)
    app.run("Copied", str(tmp_path / "beside.xlsx"))
    with zipfile.ZipFile(tmp_path / "beside.xlsx") as package:
        rels = package.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    if sys.platform == "win32":
        assert f'Target="{RECORD["files"]["copy1.xlsx"]["link"]}"' in rels
    with pytest.raises(VBAUnsupportedError, match="links to files"):
        app.run("Copied", str(tmp_path / "sub" / "elsewhere.xlsx"))
    assert sorted(path.name for path in tmp_path.rglob("*.xlsx")) == ["beside.xlsx", "named.xlsx"]


def test_a_copy_elsewhere_of_a_workbook_linking_to_other_workbooks_is_refused(tmp_path: Path) -> None:
    # Dest.xlsx, which Excel saved with formulas reading Source.xlsx (tests/fixtures/external_links.json).
    record = json.loads((FIXTURES / "external_links.json").read_text(encoding="utf-8"))
    path = tmp_path / "Dest.xlsx"
    with zipfile.ZipFile(path, "w") as package:
        for name, text in record["dest_parts"].items():
            data = base64.b64decode(text[len("base64:"):]) if text.startswith("base64:") \
                else text.replace("<folder>", str(tmp_path)).encode("utf-8")
            package.writestr(name, data)
    (tmp_path / "sub").mkdir()
    app = ExcelApplication.open(path)
    assert app.workbook.linked_books
    app.workbook.SaveCopyAs(str(tmp_path / "beside.xlsx"))
    with pytest.raises(VBAUnsupportedError, match="links to files"):
        app.workbook.SaveCopyAs(str(tmp_path / "sub" / "elsewhere.xlsx"))
    assert not (tmp_path / "sub" / "elsewhere.xlsx").exists()


def test_save_copy_as_without_a_name_text_gives_is_not_implemented(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    with pytest.raises(VBAUnsupportedError, match="without a filename"):
        app.workbook.SaveCopyAs()
    with pytest.raises(VBAUnsupportedError, match="not text"):
        app.workbook.SaveCopyAs(5)
    assert not list(tmp_path.iterdir())
