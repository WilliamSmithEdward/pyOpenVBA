"""Application.ExecuteExcel4Macro reading cells and working out expressions (scripts/measure_execute_excel4.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/execute_excel4.json").read_text(encoding="utf-8"))
#: A closed workbook's cells inside an expression, which reports itself.
REPORTED = {"SUM('<folder>\\[closed.xlsx]Data'!R1C1:R1C1)"}
ROWS = [row for row in MEASURED["rows"] if row["text"] not in REPORTED]


def _given(text: str, folder: Path, here: str) -> str:
    return text.replace("<folder>", folder.as_posix()).replace("<here>", here).replace('"', '""')


def _ready(folder: Path) -> tuple[ExcelApplication, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The file's own separator on every platform the tests run on.
    app.add_module(MEASURED["probe"].replace('"\\closed.xlsx"', '"/closed.xlsx"'), name="Probe1")
    return app, str(app.run("Setup", folder.as_posix()))


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["text"] or "empty")
def test_execute_excel4(row: dict[str, str], tmp_path: Path) -> None:
    app, here = _ready(tmp_path)
    app.add_module(f'Function Run1() As String\nRun1 = One("{_given(row["text"], tmp_path, here)}")\nEnd Function\n',
                   name="Runner")
    assert app.run("Run1") == row["result"]


@pytest.mark.parametrize("text", sorted(REPORTED))
def test_a_closed_workbook_in_an_expression_reports_itself(text: str, tmp_path: Path) -> None:
    app, here = _ready(tmp_path)
    app.add_module(f'Function Run1() As String\nRun1 = One("{_given(text, tmp_path, here)}")\nEnd Function\n', name="R")
    with pytest.raises(VBAUnsupportedError):
        app.run("Run1")


def test_the_sandbox_keeps_a_closed_workbook_shut(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Sub Go()\nDebug.Print Application.ExecuteExcel4Macro("\'{tmp_path.as_posix()}/[closed.xlsx]Data\'!R1C1")\n'
                   'End Sub\n', name="M")
    with pytest.raises(Exception, match="(?i)sandbox"):
        app.run("Go")


def test_an_excel4_function_reports_itself() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Go()\nDebug.Print Application.ExecuteExcel4Macro("GET.WORKBOOK(1)")\nEnd Sub\n', name="M")
    with pytest.raises(VBAUnsupportedError, match="Excel 4 function"):
        app.run("Go")


def test_a_closed_workbook_stays_closed(tmp_path: Path) -> None:
    app, here = _ready(tmp_path)
    given = _given("'<folder>\\[closed.xlsx]Data'!R1C1", tmp_path, here)
    app.add_module(f'Function Run1() As String\nRun1 = One("{given}") & "/" & Workbooks.Count & "/" & ActiveWorkbook.Name\n'
                   'End Function\n', name="Runner")
    assert app.run("Run1") == f"Double:42/1/{here}"
