"""What becomes of a name local to a sheet when the sheet is deleted (scripts/measure_deleted_local_names.py)."""

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

MEASURED: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/deleted_local_names.json").read_text(encoding="utf-8"))
BOOKS: dict[str, dict[str, Any]] = MEASURED["books"]
#: What is left reads a name by the deleted sheet's name, and nothing by the bare name: Excel keeps it as the workbook's.
KEPT = {"qualified only", "qualified twice", "read by a workbook name", "read by a workbook name alone", "read by a local name on the sheet",
        "bare on the deleted sheet", "qualified on the deleted sheet too", "an area", "a constant", "hidden", "with a comment", "a name of another sheet too"}


def _deleted(name: str, folder: Path) -> tuple[str, Path]:
    """The session's reads after the deletion, with the workbook's name as one word, and the model's save."""
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    start = folder / "start.xlsx"
    make = MEASURED["make"].replace("{extra}", BOOKS[name]["extra"])
    made.add_module(make, name="Maker")
    made.run("Maker.Make", str(start))
    app = ExcelApplication.open(start, with_vba=False, sandbox=False)
    app.add_module(MEASURED["reader"] + 'Public Function Go() As String\n    Application.DisplayAlerts = False\n    ActiveWorkbook.Worksheets("Gone").Delete\n    Go = Reads(ActiveWorkbook)\nEnd Function\n', name="E")
    session = str(app.run("E.Go")).replace(start.name, "BOOK")
    saved = folder / "saved.xlsx"
    app.save(saved)
    return session, saved


def _file_says(path: Path) -> dict[str, list[str]]:
    with zipfile.ZipFile(path) as package:
        return {"names": re.findall(r"<definedName\b[^>]*>[^<]*</definedName>", package.read("xl/workbook.xml").decode("utf-8")),
                "formulas": re.findall(r"<f[^>]*>([^<]*)</f>", package.read("xl/worksheets/sheet1.xml").decode("utf-8"))}


@pytest.mark.parametrize("name", list(BOOKS))
def test_the_session_reads_what_excel_reads(name: str, tmp_path: Path) -> None:
    assert _deleted(name, tmp_path)[0] == BOOKS[name]["excel_session"]


@pytest.mark.parametrize("name", list(BOOKS))
def test_the_file_holds_what_excels_file_holds(name: str, tmp_path: Path) -> None:
    assert _file_says(_deleted(name, tmp_path)[1]) == BOOKS[name]["excel_file"]


@pytest.mark.parametrize("name", list(BOOKS))
def test_excel_reads_the_models_save_as_its_own(name: str) -> None:
    assert BOOKS[name]["model_reopened"] == BOOKS[name]["excel_reopened"]


def test_the_rule_is_what_excel_did() -> None:
    """Kept where a name stands in Excel's file with no sheet, and that sets apart every book measured."""
    kept = {name for name, row in BOOKS.items() if any('name="Mine"' in tag and "localSheetId" not in tag and "Keep!" not in tag for tag in row["excel_file"]["names"])}
    assert kept == KEPT


LIVE = os.environ.get("RUN_LIVE_EXCEL") == "1" and importlib.util.find_spec("pyvbaharness") is not None
OPENED = MEASURED["reader"] + """Public Function Opened(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    Opened = Reads(wb)
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
        for number, name in enumerate(BOOKS):
            folder = tmp_path / str(number)
            folder.mkdir()
            saved = _deleted(name, folder)[1]
            answer = str(excel.run_vba(OPENED, "Opened", (str(saved),), timeout=60.0, line_numbers=False).value).replace(saved.name, "BOOK")
            if answer != BOOKS[name]["excel_reopened"]:
                wrong.append(name)
    assert not wrong
