"""The tab a saved workbook opens on, and the first it shows, once a macro has moved, added or deleted tabs
(scripts/measure_active_tab.py)."""

from __future__ import annotations

import importlib.util
import json
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parents[1]
HERE = ROOT / "tests/fixtures/active_tab"
MEASURED: dict[str, Any] = json.loads((HERE / "active_tab.json").read_text(encoding="utf-8"))
ROWS: list[dict[str, str]] = MEASURED["rows"]
_spec = importlib.util.spec_from_file_location("measure_active_tab", ROOT / "scripts/measure_active_tab.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)


def _id(row: dict[str, str]) -> str:
    return f"{row['book']}-{row['case']}"


def _acted(row: dict[str, str], tmp_path: Path) -> Path:
    file = tmp_path / f"{row['book']}.xlsx"
    shutil.copy(HERE / file.name, file)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["act"].format(action=MEASURED["actions"][row["case"]]), name="M")
    app.run("Act", file.as_posix())
    return file


@pytest.mark.parametrize("row", ROWS, ids=_id)
def test_a_save_says_which_tab_as_excel_does(row: dict[str, str], tmp_path: Path) -> None:
    assert PROBE.view(_acted(row, tmp_path)) == row["view"]


@pytest.mark.parametrize("row", ROWS, ids=_id)
def test_a_reopened_workbook_is_on_the_tab_it_was_on(row: dict[str, str], tmp_path: Path) -> None:
    file = _acted(row, tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["read"], name="M")
    assert app.run("ReadIt", file.as_posix()) == row["reads"]
