"""What a macro deletes leaves the file: every part of the package searched for a marker (scripts/measure_deleted_data.py).

The objects here come from a file Excel saved, as a table or a pivot table a file holds is written differently from one a macro made.
"""

from __future__ import annotations

import importlib.util
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.excel import ExcelFile

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "deleted_data.xlsx"
MACRO_BOOK = FIXTURES / "deleted_data.xlsm"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "deleted_data.json").read_text(encoding="utf-8"))
EDITS: dict[str, dict[str, Any]] = MEASURED["edits"]


def _held(path: Path, markers: list[str]) -> dict[str, list[str]]:
    """Each marker against the parts holding it, of every part of the package decompressed, and a VBA project's source."""
    found: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            data = package.read(name)
            for marker in markers:
                if marker.encode() in data or marker.encode("utf-16-le") in data:
                    found.setdefault(marker, []).append(name)
        project = "xl/vbaProject.bin" in package.namelist()
    if project:
        with ExcelFile(path) as host:
            for module in host.vba_project().modules:
                for marker in markers:
                    if marker in module.source:
                        found.setdefault(marker, []).append(f"vba:{module.name}")
    return found


def _saved(row: dict[str, Any], target: Path) -> None:
    start = MACRO_BOOK if row["macro"] else BOOK
    app = ExcelApplication.open(start, with_vba=row["macro"], sandbox=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Application.DisplayAlerts = False\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("Main")\n    {row["edit"]}\nEnd Sub\n', name="MeasureEdit")
    app.run("MeasureEdit.Edit")
    app.save(target)


@pytest.mark.parametrize("name", [name for name in EDITS if "excel" in EDITS[name]])
def test_no_part_holds_what_excels_file_does_not(name: str, tmp_path: Path) -> None:
    """The check that finds a deleted sheet in a file: every marker in every part, against Excel's save of the same edit."""
    row = EDITS[name]
    target = tmp_path / ("saved.xlsm" if row["macro"] else "saved.xlsx")
    _saved(row, target)
    left = _held(target, MEASURED["markers"])
    for marker in row["should_go"]:
        if marker not in row["excel"]:
            assert marker not in left, (marker, left.get(marker))
    # Nor has it lost what Excel kept.
    for marker, parts in row["excel"].items():
        if marker != "MKVBASUB" and not (marker == "MKMODSUB" and not parts):
            assert marker in left, marker


@pytest.mark.parametrize("clear", ["Clear", "ClearContents"])
def test_clearing_all_of_a_pivot_table_takes_it_and_its_cache(clear: str) -> None:
    """As Excel reads it in the session and from its save: no table and no cache."""
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(f'Public Function Go() As String\n    With ActiveWorkbook\n        .Worksheets("Other").PivotTables("MKPIVOT").TableRange2.{clear}\n        Go = .Worksheets("Other").PivotTables.Count & "|" & .PivotCaches.Count\n    End With\nEnd Function\n', name="E")
    assert app.run("E.Go") == "0|0"


def test_the_markers_were_there_to_find() -> None:
    wanted = {marker for row in EDITS.values() for marker in row["should_go"]}
    start = {**MEASURED["start_macro"], **MEASURED["start"]}
    assert wanted <= set(start)


def test_excel_leaves_a_deleted_sheets_macro_name_in_its_compiled_part() -> None:
    """Though not in any module's source: why a module's source is what this holds the model to."""
    row = EDITS["a sheet with VBA deleted"]
    assert row["excel"]["MKVBASUB"] == ["xl/vbaProject.bin"]
    assert "MKVBATEXT" not in row["excel"]


def test_a_deleted_table_from_a_file_goes_from_the_file(tmp_path: Path) -> None:
    """Left in tableParts over cleared cells, Excel refused the file; after Unlist, the table came back."""
    for name in ("ListObject.Delete", "ListObject.Unlist"):
        target = tmp_path / f"{name}.xlsx"
        _saved(EDITS[name], target)
        with zipfile.ZipFile(target) as package:
            names = package.namelist()
            assert not [part for part in names if part.startswith("xl/tables/")]
            assert b"tablePart" not in package.read("xl/worksheets/sheet1.xml")
            assert b"tables/" not in package.read("xl/worksheets/_rels/sheet1.xml.rels")
            assert b"table+xml" not in package.read("[Content_Types].xml")


def test_a_second_table_stays_when_one_goes(tmp_path: Path) -> None:
    app = ExcelApplication.open(BOOK, with_vba=False, sandbox=False)
    app.add_module('Public Sub Edit()\n    Dim ws As Object\n    Set ws = ActiveWorkbook.Worksheets("Main")\n    ws.Range("M1").Value = "h": ws.Range("M2").Value = 1\n    ws.ListObjects.Add(1, ws.Range("M1:M2"), , 1).Name = "Stays"\nEnd Sub\nPublic Sub Gone()\n    ActiveWorkbook.Worksheets("Main").ListObjects("MKTABLE").Delete\nEnd Sub\n', name="E")
    app.run("E.Edit")
    first = tmp_path / "first.xlsx"
    app.save(first)
    again = ExcelApplication.open(first, with_vba=False, sandbox=False)
    again.add_module('Public Function Gone() As String\n    ActiveWorkbook.Worksheets("Main").ListObjects("MKTABLE").Delete\n    Gone = ActiveWorkbook.Worksheets("Main").ListObjects.Count\nEnd Function\n', name="E")
    assert again.run("E.Gone") == "1"
    second = tmp_path / "second.xlsx"
    again.save(second)
    with zipfile.ZipFile(second) as package:
        assert len(re.findall(rb"<tablePart\b", package.read("xl/worksheets/sheet1.xml"))) == 1
        assert re.search(rb'<tableParts count="1">', package.read("xl/worksheets/sheet1.xml"))
        assert len([part for part in package.namelist() if part.startswith("xl/tables/")]) == 1


def test_an_image_two_pictures_share_stays_until_both_go(tmp_path: Path) -> None:
    with zipfile.ZipFile(BOOK) as package:
        image = package.read("xl/media/image1.png")
    picture = tmp_path / "same.png"
    picture.write_bytes(image)
    app = ExcelApplication.open(BOOK, with_vba=False, sandbox=False)
    app.add_module(f'Public Sub Edit()\n    ActiveWorkbook.Worksheets("Main").Shapes.AddPicture("{picture}", False, True, 1, 1, 20, 20).Name = "Second"\nEnd Sub\nPublic Sub One()\n    ActiveWorkbook.Worksheets("Main").Shapes("MKPIC").Delete\nEnd Sub\nPublic Sub Other()\n    ActiveWorkbook.Worksheets("Main").Shapes("Second").Delete\nEnd Sub\n', name="E")
    app.run("E.Edit")
    app.save(tmp_path / "both.xlsx")
    with zipfile.ZipFile(tmp_path / "both.xlsx") as package:
        assert len([part for part in package.namelist() if part.startswith("xl/media/")]) == 1
    app.run("E.One")
    first = tmp_path / "first.xlsx"
    app.save(first)
    with zipfile.ZipFile(first) as package:
        assert [part for part in package.namelist() if part.startswith("xl/media/")]
    app.run("E.Other")
    second = tmp_path / "second.xlsx"
    app.save(second)
    with zipfile.ZipFile(second) as package:
        assert not [part for part in package.namelist() if part.startswith("xl/media/")]


def test_an_unedited_file_keeps_its_strings_byte_for_byte(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    ExcelApplication.open(BOOK, with_vba=False).save(target)
    with zipfile.ZipFile(BOOK) as old, zipfile.ZipFile(target) as new:
        assert old.read("xl/sharedStrings.xml") == new.read("xl/sharedStrings.xml")
        assert sorted(old.namelist()) == sorted(new.namelist())


def test_an_edit_keeps_the_strings_still_used_at_their_numbers(tmp_path: Path) -> None:
    """Only the text that went is emptied: every number a kept sheet holds stands for what it did."""
    target = tmp_path / "saved.xlsx"
    _saved(EDITS["ClearContents"], target)
    entries = re.compile(rb"<si\b[^>]*?(?:/>|>.*?</si>)", re.DOTALL)
    with zipfile.ZipFile(BOOK) as old, zipfile.ZipFile(target) as new:
        before, after = entries.findall(old.read("xl/sharedStrings.xml")), entries.findall(new.read("xl/sharedStrings.xml"))
    assert len(before) == len(after)
    assert b"MKCLEAR" not in b"".join(after)
    again = ExcelApplication.open(target, with_vba=False)
    again.add_module('Public Function Go() As String\n    With ActiveWorkbook\n        Go = .Worksheets("Main").Range("A2").Value & "|" & .Worksheets("Src").Range("A2").Value & "|" & .Worksheets("MKSHEET").Range("A1").Value\n    End With\nEnd Function\n', name="E")
    assert again.run("E.Go") == "MKOVER|MKRECORD|MKSHEETCELL"


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None
OPENS = """Public Function Tryit(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    Tryit = CStr(Err.Number)
    If Not wb Is Nothing Then wb.Close False
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
End Function
"""


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_opens_the_models_save_after_every_deletion(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    refused: dict[str, object] = {}
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, (name, row) in enumerate(EDITS.items()):
            if "model" not in row:
                continue
            target = tmp_path / f"saved{number}{'.xlsm' if row['macro'] else '.xlsx'}"
            _saved(row, target)
            answer = excel.run_vba(OPENS, "Tryit", (str(target),), timeout=60.0, line_numbers=False)
            if answer.value != "0":
                refused[name] = answer.value
    assert not refused
