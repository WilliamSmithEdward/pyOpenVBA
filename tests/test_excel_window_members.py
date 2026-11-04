"""Window and Windows members that need no measure of the screen (scripts/measure_window_members.py)."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/window_members.json").read_text(encoding="utf-8"))


def _names(text: str) -> str:
    return re.sub(r"Book\d+", "Book", text)


@pytest.mark.parametrize("row", MEASURED["rows"], ids=lambda row: row["name"])
def test_window_members(row: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=row["action"]), name="Probe1")
    assert _names(str(app.run("Probe"))) == _names(row["result"])


def test_the_active_pane_of_a_split_window_reports_itself() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Go()\nRange("B2").Select\nActiveWindow.FreezePanes = True\nDebug.Print ActiveWindow.ActivePane.Index\n'
                   'End Sub\n', name="M")
    with pytest.raises(VBAUnsupportedError, match="active pane"):
        app.run("Go")
