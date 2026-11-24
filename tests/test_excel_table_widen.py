"""ListObjects.Add widening its header cells' columns, replayed against the in-memory model.

tests/fixtures/table_widen.json is what scripts/measure_table_widen.py saw in live Excel on a 144-DPI display
(StandardHeight 14.5 at both ends of every case). The model's geometry is a 96-DPI display's, so widths are not
compared: what is compared is whether each case ran or failed, its r, the table it made and whether a column stayed
hidden, and which columns widened. Excel widens the column under every header cell, text a macro typed included, to
fit the header and its dropdown, never narrower, a hidden one not at all, and Resize and ListColumns.Add widen
nothing. The model widens nothing yet: how wide a column gets is the bold AutoFit width plus 22 pixels at 11pt here,
less at bigger sizes and in italic, and the rule at 96 DPI is not measured, so those cases are strict xfails.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "table_widen.json").read_text(encoding="utf-8"))
CASES: list[list[str]] = RECORD["cases"]
#: Cases with a pivot table, which the model's own pivot fit sizes at 96 DPI.
PIVOT = {22}


def _facts(answer: str) -> tuple[str, ...]:
    """What a case's answer says that does not depend on the display: how it ended, r, the table and the hidden
    column, and which columns ended wider than they started."""
    fields = answer.split("|")
    before = [float(one) for one in fields[3].removeprefix("before ").split(",")]
    after = [float(one) for one in fields[4].removeprefix("after ").split(",")]
    rest = tuple(one for one in fields[6:] if not one.startswith("sh "))
    wider = ",".join(str(index + 1) for index, (old, new) in enumerate(zip(before, after, strict=True)) if new > old)
    return (fields[1], fields[2], *rest, f"wider {wider}")


def _ran(index: int) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["source"], name="Probe")
    return str(app.run("Probe", index))


def _widened(index: int) -> bool:
    return _facts(CASES[index][1])[-1] != "wider "


def _cases(widened: bool) -> list[int]:
    return [index for index in range(len(CASES)) if index not in PIVOT and _widened(index) == widened]


@pytest.mark.parametrize("index", _cases(False) + _cases(True))
def test_each_case_runs_and_makes_the_table_excel_makes(index: int) -> None:
    assert _facts(_ran(index))[:-1] == _facts(CASES[index][1])[:-1]


@pytest.mark.parametrize("index", _cases(False))
def test_a_table_over_wide_enough_columns_widens_nothing(index: int) -> None:
    assert _facts(_ran(index))[-1] == _facts(CASES[index][1])[-1]


@pytest.mark.xfail(strict=True, reason="the widths ListObjects.Add widens header columns to are not measured at 96 DPI")
@pytest.mark.parametrize("index", _cases(True))
def test_a_table_widens_its_header_columns_as_excel_does(index: int) -> None:
    assert _facts(_ran(index))[-1] == _facts(CASES[index][1])[-1]


def test_the_fixture_was_measured_at_one_dpi() -> None:
    assert {re.findall(r"sh ([\d.]+)", answer)[0] for _, answer in CASES} == {"14.5"}
