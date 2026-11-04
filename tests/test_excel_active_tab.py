"""The tab a saved workbook opens on, and the first it shows, once a macro has moved, added or deleted tabs
(scripts/measure_active_tab.py)."""

from __future__ import annotations

import importlib.util
import json
import re
import shutil
import zipfile
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


#: Excel scrolled its tab bar on to keep the active tab in view, here firstSheet 1 with six tabs in a window 700 wide. How far
#: goes by the tabs' widths in the window, which the model does not lay out; what it writes opens all the same, on the same tab.
SCROLLED_ON = {("six_sheets", "move_first_after_sixth")}


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


@pytest.mark.parametrize(("action", "wanted"), [("wb.Sheets(1).Range(\"A1\").Value = 1", 'activeTab="2" firstSheet="2"'),
                                                ("wb.Sheets(1).Delete", 'activeTab="1" firstSheet="1"')])
def test_a_first_tab_past_the_active_one_comes_back_on_opening(action: str, wanted: str, tmp_path: Path) -> None:
    """Measured live with a shown window on six sheets: firstSheet=4, activeTab=2 saves as 2,2, and with sheet 1 deleted as 1,1.
    Excel may not write such a file itself, so this one is the scrolled book with its view changed by hand."""
    source = HERE / "scrolled.xlsx"
    file = tmp_path / "file.xlsx"
    with zipfile.ZipFile(source) as given, zipfile.ZipFile(file, "w", zipfile.ZIP_DEFLATED) as made:
        for info in given.infolist():
            data = given.read(info.filename)
            if info.filename == "xl/workbook.xml":
                data = data.replace(b'firstSheet="2"', b'firstSheet="4"').replace(b'activeTab="4"', b'activeTab="2"')
            elif info.filename == "xl/worksheets/sheet5.xml":
                data = data.replace(b' tabSelected="1"', b"")
            elif info.filename == "xl/worksheets/sheet3.xml":
                data = data.replace(b"<sheetView ", b'<sheetView tabSelected="1" ', 1)
            made.writestr(info, data)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["act"].format(action=action), name="M")
    app.run("Act", file.as_posix())
    assert PROBE.view(file) == wanted


@pytest.mark.parametrize("row", ROWS, ids=_id)
def test_a_save_says_which_tab_as_excel_does(row: dict[str, str], tmp_path: Path) -> None:
    view = PROBE.view(_acted(row, tmp_path))
    if (row["book"], row["case"]) in SCROLLED_ON:
        # The active tab is the same; the first differs by how much of the bar the window has room for.
        assert re.findall(r'activeTab="\d+"', view) == re.findall(r'activeTab="\d+"', row["view"])
        assert view != row["view"]
    else:
        assert view == row["view"]


@pytest.mark.parametrize("row", ROWS, ids=_id)
def test_a_reopened_workbook_is_on_the_tab_it_was_on(row: dict[str, str], tmp_path: Path) -> None:
    file = _acted(row, tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["read"], name="M")
    assert app.run("ReadIt", file.as_posix()) == row["reads"]
