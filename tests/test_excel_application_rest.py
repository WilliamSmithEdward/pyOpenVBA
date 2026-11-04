"""Pane, and what is left of Application and Window that needs no screen (scripts/measure_application_rest.py)."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/application_rest.json").read_text(encoding="utf-8"))
#: Measured but left reporting themselves: what DDERequest's error value holds was not read, and UsedObjects(1) answers
#: from what the measuring session had used.
REPORTED = {"dde_request_nothing", "used_objects_item"}
ROWS = [row for row in MEASURED["rows"] if row["name"] not in REPORTED]


def _run(action: str) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=action), name="Measure")
    return str(app.run("Probe"))


@pytest.mark.parametrize("row", ROWS, ids=lambda row: row["name"])
def test_application_rest(row: dict[str, str]) -> None:
    assert re.sub(r"Book\d+", "Book", _run(row["action"])) == re.sub(r"Book\d+", "Book", row["result"])


@pytest.mark.parametrize("row", [row for row in MEASURED["rows"] if row["name"] in REPORTED], ids=lambda row: row["name"])
def test_what_was_left_out_reports_itself(row: dict[str, str]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(row["action"])


def test_the_sandbox_keeps_register_xll_out() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Go()\nDebug.Print Application.RegisterXLL("nothing.xll")\nEnd Sub\n', name="M")
    with pytest.raises(Exception, match="(?i)sandbox"):
        app.run("Go")


def test_a_real_xll_is_never_loaded(tmp_path: Path) -> None:
    library = tmp_path / "real.xll"
    library.write_bytes(b"not loaded")
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Sub Go()\nDebug.Print Application.RegisterXLL("{library.as_posix()}")\nEnd Sub\n', name="M")
    with pytest.raises(VBAUnsupportedError, match="XLL"):
        app.run("Go")
