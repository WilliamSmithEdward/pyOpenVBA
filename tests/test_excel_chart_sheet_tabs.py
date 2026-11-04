"""Where tabs land in a workbook holding chart sheets (scripts/measure_chart_sheet_tabs.py)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

HERE = Path(__file__).parent / "fixtures/chart_sheets"
MEASURED: dict[str, Any] = json.loads((HERE / "tabs.json").read_text(encoding="utf-8"))
ROWS: list[dict[str, str]] = MEASURED["rows"]
SAVED: list[dict[str, str]] = MEASURED["saved"]
#: What the model reports rather than does yet: a new chart sheet, a copy of one, a group holding one.
REPORTED = {"add_type_chart", "charts_add", "charts_before_1", "charts_before_2", "charts_after_1", "charts_after_2", "charts_after_last", "charts_two",
            "activated_chart_charts_add", "copy_chart_last", "copy_chart_first", "group_select"}


def _answer(row: dict[str, str]) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=MEASURED["actions"][row["case"]]), name="M")
    return str(app.run("Go", (HERE / f"{row['book']}.xlsx").as_posix()))


#: What the model reports where Excel goes on: a book left with chart sheets alone, and a worksheet a chart reads, deleted;
#: as in any workbook, a copy of a sheet with print names placed first, and a rename where a formula spans sheets.
REPORTED_ON = {(book, "delete_active_worksheet") for book in ("first", "last", "two_charts", "active", "middle", "names")} | {
    (book, "delete_sheet") for book in ("first", "last", "two_charts", "active")} | {("names", "copy_sheet_first"), ("names", "rename_sheet"),
                                                                                     ("names", "added_then_named")}


def _reported(row: dict[str, str]) -> bool:
    return row["case"] in REPORTED or (row["book"], row["case"]) in REPORTED_ON


def _id(row: dict[str, str]) -> str:
    return f"{row['book']}-{row['case']}"


@pytest.mark.parametrize("row", [row for row in ROWS if not _reported(row)], ids=_id)
def test_a_tab_lands_where_excel_puts_it(row: dict[str, str]) -> None:
    assert _answer(row) == row["answer"]


@pytest.mark.parametrize("row", SAVED, ids=_id)
def test_excel_reads_the_models_save_as_its_own(row: dict[str, str]) -> None:
    """Recorded by the probe: the action made by Excel and saved, and made by the model and saved, each read by Excel."""
    assert (row["model"], row["model_view"]) == (row["excel"], row["excel_view"])


@pytest.mark.parametrize("row", [row for row in ROWS if _reported(row)], ids=_id)
def test_what_the_model_does_not_make_reports_itself(row: dict[str, str]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _answer(row)
