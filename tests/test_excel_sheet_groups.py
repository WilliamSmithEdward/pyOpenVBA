"""Groups of sheets, Sheets(Array(...)) and the window's selected sheets, replayed against the in-memory model.

tests/fixtures/sheet_groups.json is what scripts/measure_sheet_groups.py saw
in live Excel: each case makes a workbook of four sheets that read one
another, runs a statement and reads back what it gave, the active sheet,
the selected sheets and every sheet's visibility and formulas.
tests/fixtures/sheet_group_moves.json is what
scripts/measure_sheet_group_moves.py saw: the tab order after 608 group
moves and 48 group copies beside another sheet, and the file a grouped
selection saves.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._sheet_move import placed
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((FIXTURES / "sheet_groups.json").read_text(encoding="utf-8"))
MOVES: dict[str, Any] = json.loads((FIXTURES / "sheet_group_moves.json").read_text(encoding="utf-8"))
NAMES = ["One", "Two", "Three", "Four", "Five", "Six"]
#: Cases the model reports it does not make, and why.
UNSUPPORTED = {"copy_new_outside": "a copy reading a sheet left behind needs an external link to its workbook"}


def _run(case: dict[str, Any]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + case["code"], name="Probe")
    return app.run(f"Case_{case['name']}")


@pytest.mark.parametrize("case", RECORD["cases"], ids=[case["name"] for case in RECORD["cases"]])
def test_a_group_does_what_excel_does(case: dict[str, Any]) -> None:
    if case["name"] in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            _run(case)
        return
    assert _run(case) == case["answer"]


def test_each_measured_move_puts_the_group_where_excel_does() -> None:
    """The placement rule over every move measured, the shifted ones among them."""
    missed: list[tuple[dict[str, Any], list[str]]] = []
    for move in MOVES["moves"]:
        sheets: list[str] = NAMES[:move["sheets"]]
        group: list[str] = move["group"]
        got = placed(sheets, group, str(move["anchor"]), bool(move["after"]))
        if got != move["order"]:
            missed.append((move, got))
    assert not missed, missed[:3]


def _swept(rows: list[dict[str, Any]], verb: str) -> list[str]:
    """Each row's group moved or copied in a fresh workbook in the model, read back as Excel's sweep reads it."""
    body = ["Public Function Sweep() As String", "    Dim wb As Object, s As Object, out As String, i As Long"]
    for row in rows:
        listed = ", ".join(f'"{name}"' for name in row["group"])
        body += ["    Set wb = Workbooks.Add(xlWBATWorksheet)", '    wb.Worksheets(1).Name = "One"',
                 *(f'    wb.Worksheets.Add(After:=wb.Worksheets({at})).Name = "{name}"'
                   for at, name in enumerate(NAMES[1:row["sheets"]], start=1)),
                 '    wb.Worksheets("One").Activate',
                 f'    wb.Sheets(Array({listed})).{verb} {"After" if row["after"] else "Before"}:=wb.Sheets("{row["anchor"]}")',
                 "    For Each s In wb.Sheets", '        out = out & s.Name & ","', "    Next",
                 '    out = out & ActiveSheet.Name & "|"', "    wb.Close False"]
    body += ["    Sweep = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(body) + "\n", name="Probe")
    return [part for part in str(app.run("Sweep")).split("|") if part]


@pytest.mark.parametrize("verb", ["Move", "Copy"])
def test_each_measured_move_and_copy_leaves_the_tabs_excel_leaves(verb: str) -> None:
    rows: list[dict[str, Any]] = MOVES["moves" if verb == "Move" else "copies"]
    expected = [",".join([*row["order"], row["active"]]) for row in rows]
    assert _swept(rows, verb) == expected


def test_a_grouped_selection_saves_and_opens_again_as_excels_does(tmp_path: Path) -> None:
    """Each selected sheet's tab is selected in the file, the active one's the workbook's active tab, and the group
    is selected again once the file opens."""
    path = tmp_path / "grouped.xlsx"
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MOVES["saved_code"], name="Probe")
    app.run("Saved", str(path))
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        selected: list[bool] = []
        for number in range(1, 5):
            sheet = package.read(f"xl/worksheets/sheet{number}.xml").decode("utf-8")
            view = re.search(r"<sheetView\b[^>]*>", sheet)
            assert view is not None
            selected.append('tabSelected="1"' in view.group(0))
    active = re.search(r'activeTab="(\d+)"', workbook)
    assert {"tab_selected": selected, "active_tab": int(active.group(1)) if active else 0} == MOVES["saved"]
    assert app.run("Reopened", str(path)) == MOVES["reopened"]
