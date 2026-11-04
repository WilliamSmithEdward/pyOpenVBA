"""What a formula that read a sheet becomes when the sheet is deleted (scripts/measure_sheet_delete_refs.py)."""

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
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._parse import has_dangling_reference, without_dangling_references

FIXTURES = Path(__file__).parent / "fixtures"
BOOK = FIXTURES / "sheet_delete_refs.xlsx"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "sheet_delete_refs.json").read_text(encoding="utf-8"))
COUNT = len(MEASURED["formulas"])
#: A workbook Excel opens lists hidden names of its own that the model does not: _xlfn.IFERROR, _xlfn.SINGLE.
RAN = {name: row for name, row in MEASURED["edits"].items() if "session" in row["model"]}
REFUSED = {name: row for name, row in MEASURED["edits"].items() if "refuses" in row["model"]}


def _comparable(reads: list[str]) -> list[str]:
    """Reads without Excel's hidden names, and with the workbook's own name, which differs by file, as one word."""
    return [re.sub(r"=[\w.]+\.xlsx!", "=BOOK!", one) for one in reads if "_xlfn" not in one]


def _app(start: Path, edit: str) -> Any:
    app = ExcelApplication.open(start, with_vba=False)
    app.add_module(MEASURED["reader"] + f'Public Function Run(ByVal count As Long) As String\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("Reader")\n    Application.DisplayAlerts = False\n    {edit}\n    Run = ReadAll(wb, count)\nEnd Function\n', name="E")
    return app


def _reader_sheet(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("xl/worksheets/sheet1.xml").decode("utf-8")


@pytest.mark.parametrize("name", list(RAN))
def test_the_session_reads_what_excel_reads(name: str) -> None:
    row = RAN[name]
    session = str(_app(BOOK, row["edit"]).run("E.Run", COUNT)).split("@@")
    assert _comparable(session) == _comparable(row["excel_session"])


@pytest.mark.parametrize("name", list(RAN))
def test_excel_reads_the_models_save_as_its_own(name: str) -> None:
    row = RAN[name]
    assert _comparable(row["model_reopened"]) == _comparable(row["excel_reopened"])


@pytest.mark.parametrize("name", list(RAN))
def test_the_file_holds_the_formulas_excels_file_holds(name: str, tmp_path: Path) -> None:
    """Only the error is written for what a deleted sheet left: <f>#REF!*2</f>, for the session's =#REF!A1*2."""
    row = RAN[name]
    app = _app(BOOK, row["edit"])
    app.run("E.Run", COUNT)
    target = tmp_path / "saved.xlsx"
    app.save(target)
    written = re.findall(r'<c r="([A-Z]+\d+)"(?: s="\d+")?(?: t="(\w+)")?[^>]*><f[^>]*?>(.*?)</f><v>(.*?)</v></c>', _reader_sheet(target))
    assert sorted(f"{ref}|{kind}|{formula}|{value}" for ref, kind, formula, value in written) == sorted(row["excel_file"]["cells"])


@pytest.mark.parametrize("name", list(RAN))
def test_the_names_in_the_file_are_excels(name: str) -> None:
    row = RAN[name]
    assert sorted(row["model_file"]["names"]) == sorted(row["excel_file"]["names"])


def test_excel_will_not_take_what_it_shows() -> None:
    """A deleted sheet's =#REF!A1 is shown, and is error 1004 written back; =#REF! is a formula like any other."""
    row = MEASURED["edits"]["delete, then rewrite a formula"]
    assert row["excel_session"][2].startswith("=#REF!A1*2=")
    # The probe's own =#REF!A1 and =SUM(#REF!A1:B2) never reached their cells, which read empty.
    at = MEASURED["formulas"].index("=#REF!A1")
    assert row["excel_session"][at + 1].startswith("=Empty=")
    assert row["excel_session"][at + 2].startswith("=Empty=")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    On Error Resume Next\n    ActiveSheet.Range("A1").Formula = "=#REF!A1"\n    Go = Err.Number & "|" & ActiveSheet.Range("A1").Formula\n    Err.Clear\n    ActiveSheet.Range("A2").Formula = "=#REF!"\n    Go = Go & "|" & Err.Number & "|" & ActiveSheet.Range("A2").Formula\nEnd Function\n', name="E")
    assert app.run("E.Go") == "1004||0|=#REF!"


@pytest.mark.parametrize(("formula", "dangling", "written"), [
    ("=#REF!A1*2", True, "=#REF!*2"),
    ("=SUM(#REF!$A$1:$B$2)", True, "=SUM(#REF!)"),
    ("=#REF!A:A", True, "=#REF!"),
    ("=SUM(#REF!1:2)", True, "=SUM(#REF!)"),
    ("=#REF!", False, "=#REF!"),
    ('="#REF!A1"', False, '="#REF!A1"'),
    ("='#REF!A1'!B1", False, "='#REF!A1'!B1"),
    ("=A1", False, "=A1"),
])
def test_what_a_deleted_sheet_leaves(formula: str, dangling: bool, written: str) -> None:
    assert has_dangling_reference(formula) == dangling
    assert without_dangling_references(formula) == written


def test_a_rename_after_the_deletion_passes_over_what_it_left(tmp_path: Path) -> None:
    """The book of the probe holds three-dimensional references, which a rename refuses, so this is a plain one."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim wb As Object\n    Set wb = ActiveWorkbook\n    Application.DisplayAlerts = False\n    wb.Worksheets.Add\n    wb.Worksheets(1).Name = "Gone"\n    wb.Worksheets(2).Name = "Stays"\n    wb.Worksheets("Stays").Range("A1").Formula = "=Gone!A1+Stays!B1"\n    wb.Worksheets("Gone").Delete\n    wb.Worksheets("Stays").Name = "Moved"\n    Go = wb.Worksheets("Moved").Range("A1").Formula\nEnd Function\n', name="E")
    assert app.run("E.Go") == "=#REF!A1+Moved!B1"


@pytest.mark.parametrize("name", list(REFUSED))
def test_what_is_still_refused_says_so(name: str) -> None:
    """Refused where the book has what a save would lose, never a wrong answer."""
    row = REFUSED[name]
    with pytest.raises((VBAUnsupportedError, VBARuntimeError)):
        start = FIXTURES / "sheet_delete_refs_objects.xlsx" if name.startswith("objects") else BOOK
        _app(start, row["edit"]).run("E.Run", COUNT)


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None
OPENED = MEASURED["reader"] + """Public Function Opened(ByVal path As String, ByVal count As Long) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    Opened = ReadAll(wb, count)
    wb.Close False
    Application.DisplayAlerts = True
End Function
"""


@pytest.mark.skipif(not LIVE, reason="requires explicitly enabled native Excel")
def test_native_excel_reads_the_models_save_as_recorded(tmp_path: Path) -> None:
    harness: Any = pytest.importorskip("pyvbaharness")
    wrong: list[str] = []
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number, (name, row) in enumerate(RAN.items()):
            app = _app(BOOK, row["edit"])
            app.run("E.Run", COUNT)
            target = tmp_path / f"saved{number}.xlsx"
            app.save(target)
            answer = excel.run_vba(OPENED, "Opened", (str(target), COUNT), timeout=60.0, line_numbers=False)
            if _comparable(str(answer.value).split("@@")) != _comparable(row["excel_reopened"]):
                wrong.append(name)
    assert not wrong
