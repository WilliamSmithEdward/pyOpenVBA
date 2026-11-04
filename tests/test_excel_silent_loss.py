"""What the model keeps of a sheet's and a workbook's settings from a file Excel saved: what it reads, and what is left
after an unrelated edit and a save (scripts/measure_silent_loss.py, scripts/measure_calc_settings_kept.py)."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

HERE = Path(__file__).parent / "fixtures/silent_loss"
MEASURED = json.loads((HERE / "silent_loss.json").read_text(encoding="utf-8"))
KEPT = json.loads((HERE / "calc_kept.json").read_text(encoding="utf-8"))
ROWS = [row for row in MEASURED["rows"] if "failed" not in row]
#: Hidden, Excel lays out no panes, so its reading of these differs from its own with the window shown, which is what
#: the model answers (tests/fixtures/windows/windows.json).
HIDDEN_LAYOUT = {"freeze_rows", "freeze_columns", "freeze_both"}
#: Panes split without freezing, which the model reports rather than guesses at (see _windows).
REPORTED = {"split", "split_columns"}
BOOK_ELEMENTS = ("workbookPr", "workbookProtection", "bookViews", "calcPr", "fileSharing")
SHEET_ELEMENTS = ("sheetPr", "sheetViews", "sheetFormatPr", "cols", "sheetProtection")
READ = '''Function Read1(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, i As Long, s As String
    Set wb = Workbooks.Open(file)
    Set ws = wb.Worksheets(1)
    wb.Activate
    For i = 0 To {last}
        s = s & "|" & Tried(i, wb, ws)
    Next
    wb.Close False
    Read1 = Mid$(s, 2)
End Function
'''
EDIT = '''Sub Edit1(ByVal file As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(file)
    {action}
    wb.Worksheets(1).Range("H20").Value = 7
    wb.Save
    wb.Close False
End Sub
'''


def _app(source: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    return app


def _reading(row: dict[str, object], file: Path) -> object:
    reads = cast(list[str], row["reads"])
    cases ="\n".join(f"    Case {index}\n        v = {read}" for index, read in enumerate(reads)) or "    Case -1\n        v = 0"
    return _app(MEASURED["help"].format(cases=cases) + READ.format(last=len(reads) - 1)).run("Read1", file.as_posix())


def _elements(file: Path) -> dict[str, str]:
    with zipfile.ZipFile(file) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        sheets = sorted(name for name in package.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name))
        parts = [("book", book, BOOK_ELEMENTS)]
        parts += [(sheet, package.read(sheet).decode("utf-8"), SHEET_ELEMENTS) for sheet in sheets]
    found: dict[str, str] = {}
    for label, text, names in parts:
        for element in names:
            match = re.search(rf"<{element}\b(?:[^>]*/>|.*?</{element}>)", text, re.DOTALL)
            found[f"{label}.{element}"] = match.group(0) if match else ""
    return found


@pytest.mark.parametrize("row", [row for row in ROWS if row["name"] not in REPORTED | HIDDEN_LAYOUT], ids=lambda row: row["name"])
def test_the_model_reads_a_file_as_excel_does(row: dict[str, object]) -> None:
    assert _reading(row, HERE / f"{row['name']}.xlsx") == row["excel"]


@pytest.mark.parametrize("name", sorted(HIDDEN_LAYOUT))
def test_frozen_panes_read_as_excel_lays_them_out(name: str) -> None:
    row = next(one for one in ROWS if one["name"] == name)
    assert _reading(row, HERE / f"{name}.xlsx") == row["model"]


@pytest.mark.parametrize("name", sorted(REPORTED))
def test_a_split_reports_itself(name: str) -> None:
    row = next(one for one in ROWS if one["name"] == name)
    with pytest.raises(VBAUnsupportedError):
        _reading(row, HERE / f"{name}.xlsx")


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
def test_an_unrelated_edit_and_a_save_keep_the_settings(row: dict[str, object], tmp_path: Path) -> None:
    """calcPr's iteration and calcOnSave are the application's, so calc_kept below says what a save does with them."""
    file = tmp_path / "file.xlsx"
    shutil.copy(HERE / f"{row['name']}.xlsx", file)
    given = _elements(file)
    _app(EDIT.format(action="")).run("Edit1", file.as_posix())
    kept = _elements(file)
    lost = {key: (given[key], kept[key]) for key in given if given[key] != kept[key] and key != "book.calcPr"}
    assert not lost


def _calc(file: Path) -> str:
    with zipfile.ZipFile(file) as package:
        found = re.search(r"<calcPr\b[^>]*/>", package.read("xl/workbook.xml").decode("utf-8"))
    assert found is not None
    return re.sub(r'\scalcId="\d+"', "", found.group(0))


@pytest.mark.parametrize("row", KEPT["rows"], ids=lambda row: f"{row['file']}-{row['action']}")
def test_a_macro_saves_calculation_settings_as_excel_does(row: dict[str, str], tmp_path: Path) -> None:
    file = tmp_path / "file.xlsx"
    shutil.copy(HERE / f"{row['file']}.xlsx", file)
    _app(EDIT.format(action=KEPT["actions"][row["action"]])).run("Edit1", file.as_posix())
    assert _calc(file) == re.sub(r'\scalcId="\d+"', "", row["saved"])
