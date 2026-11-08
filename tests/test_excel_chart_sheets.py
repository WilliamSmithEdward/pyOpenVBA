"""A workbook holding chart sheets, as Excel saved it: what a macro reads and walks, and what a save keeps
(scripts/measure_chart_sheets.py)."""

from __future__ import annotations

import importlib.util
import json
import re
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parents[1]
HERE = ROOT / "tests/fixtures/chart_sheets"
MEASURED: dict[str, Any] = json.loads((HERE / "chart_sheets.json").read_text(encoding="utf-8"))
ROWS: list[dict[str, Any]] = [row for row in MEASURED["rows"] if "failed" not in row]


def _module(name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, ROOT / f"scripts/{name}.py")
    assert spec is not None and spec.loader is not None
    module: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


PROBE = _module("measure_chart_sheets")
PARTS = _module("measure_parts_loss")
SAVES = _module("measure_chart_sheet_hidden")
HIDDEN: dict[str, Any] = json.loads((HERE / "hidden.json").read_text(encoding="utf-8"))


def _run(source: str, entry: str, *arguments: object) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    return app.run(entry, *arguments)


READ_ONE = """Function One(ByVal file As String) As String
    Dim wb As Workbook, v As Variant
    Set wb = Workbooks.Open(file)
    On Error GoTo Failed
    v = {read}
    One = CStr(v)
    Exit Function
Failed:
    One = "E" & Err.Number
End Function
"""
READS: list[str] = MEASURED["reads"]
CASES = [(row, index, read) for row in ROWS for index, read in enumerate(READS)]


@pytest.mark.parametrize(("row", "index", "read"), CASES, ids=[f"{row['name']}-{index}" for row, index, _ in CASES])
def test_a_macro_reads_a_chart_sheet_book_as_excel_does(row: dict[str, Any], index: int, read: str) -> None:
    source = READ_ONE.format(read=read)
    assert str(_run(source, "One", (HERE / row["file"]).as_posix())) == row["reads"][index]


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
def test_a_macro_walks_the_tabs_as_excel_does(row: dict[str, Any]) -> None:
    assert _run(PROBE.WALK, "Walk", (HERE / row["file"]).as_posix()) == row["walk"]


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
@pytest.mark.parametrize("route", ["macro", "python"])
def test_a_save_keeps_every_part(row: dict[str, Any], route: str, tmp_path: Path) -> None:
    given = HERE / row["file"]
    file = tmp_path / given.name
    shutil.copy(given, file)
    answer = PARTS.model_answer(PARTS.MODEL_EDIT, "Edit1", file) if route == "macro" else PARTS.python_edit(file)
    assert not str(answer).startswith(("reports", "crash")), answer
    assert PARTS.package_loss(given, file) == {}


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
def test_excel_read_what_the_model_saved(row: dict[str, Any]) -> None:
    """Recorded as measured: the edited cell reached the chart, and an active chart sheet stayed the only selected tab, but
    for a very hidden one, which Excel does not open on."""
    assert row["after"].split("|")[:2] == ["99", "99"]
    first = next(one for one in row["walk"].split("|||")[2].split("|") if one)
    assert row["active"].split("|")[2:] == ["Sheet1" if row["name"] == "very_hidden" else first.split(":")[1], "1"]


@pytest.mark.parametrize("name", sorted(HIDDEN["rows"]))
def test_a_hidden_chart_sheet_activates_but_does_not_select(name: str) -> None:
    assert _run(HIDDEN["answers"], "Answers", (HERE / f"{name}.xlsx").as_posix()) == HIDDEN["rows"][name]["answers"]


@pytest.mark.parametrize("name", sorted(HIDDEN["rows"]))
def test_a_hidden_chart_sheet_activated_saves_as_excel_saves_it(name: str, tmp_path: Path) -> None:
    file = tmp_path / f"{name}.xlsx"
    shutil.copy(HERE / file.name, file)
    _run(HIDDEN["save"], "Save1", file.as_posix())
    assert SAVES.views(file) == HIDDEN["rows"][name]["views"]


@pytest.mark.parametrize("action", ['wb.Worksheets(1).Delete', 'wb.Sheets("C").Move', 'wb.Worksheets(1).Move', 'wb.Sheets("C").Copy',
                                    'wb.Sheets(Array(1, 2)).Select'])
def test_changing_the_tabs_reports_itself(action: str) -> None:
    source = f"Sub Go(ByVal file As String)\n    Dim wb As Workbook\n    Set wb = Workbooks.Open(file)\n    {action}\nEnd Sub\n"
    with pytest.raises(VBAUnsupportedError):
        _run(source, "Go", (HERE / "middle.xlsx").as_posix())


def test_activating_moves_the_selection_in_the_file(tmp_path: Path) -> None:
    import zipfile

    file = tmp_path / "middle.xlsx"
    shutil.copy(HERE / "middle.xlsx", file)
    source = ('Sub Go(ByVal file As String)\n    Dim wb As Workbook\n    Set wb = Workbooks.Open(file)\n    wb.Sheets("C").Activate\n'
              '    wb.Save\n    wb.Close False\nEnd Sub\n')
    _run(source, "Go", file.as_posix())
    with zipfile.ZipFile(file) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        chart = package.read("xl/chartsheets/sheet1.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert 'activeTab="1"' in book
    assert 'tabSelected="1"' in chart
    assert "tabSelected" not in sheet
    assert re.search(r'<sheet name="C"', book)
