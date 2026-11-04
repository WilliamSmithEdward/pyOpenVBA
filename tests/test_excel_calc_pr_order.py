"""calcPr as Excel writes it: its attributes in order, and calcCompleted only while a calculation is to finish (tests/fixtures/saved_values.json)."""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._calculation_settings import _in_schema_order  # pyright: ignore[reportPrivateUsage]

FIXTURES = Path(__file__).parent / "fixtures"
MEASURED: dict[str, Any] = json.loads((FIXTURES / "saved_values.json").read_text(encoding="utf-8"))
ROWS = {name: row for name, row in {**MEASURED["edits"], **MEASURED["sequences"]}.items() if "excel_cells" in row and "model_cells" in row}


def _calc_pr(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        found = re.search(r"<calcPr\b[^>]*/>", package.read("xl/workbook.xml").decode("utf-8"))
    assert found is not None
    return found.group(0)


@pytest.mark.parametrize("name", list(ROWS))
def test_calc_pr_is_written_as_excel_writes_it(name: str, tmp_path: Path) -> None:
    row = ROWS[name]
    own = tmp_path / "book.xlsx"
    own.write_bytes((FIXTURES / "saved_values.xlsx").read_bytes())
    (tmp_path / "isdir.xlsx").mkdir()
    app = ExcelApplication.open(own, with_vba=False, sandbox=False)
    edit = row["edit"].replace("\\", os.sep)
    app.add_module(f'Public Sub Run(ByVal out As String)\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("One")\n    {edit}\nEnd Sub\n', name="E")
    app.run("E.Run", str(tmp_path))
    target = own if name in MEASURED["sequences"] else tmp_path / "saved.xlsx"
    if name not in MEASURED["sequences"]:
        app.save(target)
    assert _calc_pr(target) == row["excel_cells"]["#calcPr"]["tag"]


@pytest.mark.parametrize(("given", "written"), [
    ('<calcPr calcOnSave="0" calcCompleted="0" calcMode="manual" calcId="1"/>', '<calcPr calcId="1" calcMode="manual" calcCompleted="0" calcOnSave="0"/>'),
    ('<calcPr forceFullCalc="1" iterate="1" calcId="1"/>', '<calcPr calcId="1" iterate="1" forceFullCalc="1"/>'),
    ('<calcPr unknown="2" calcId="1"/>', '<calcPr calcId="1" unknown="2"/>'),
])
def test_the_order(given: str, written: str) -> None:
    assert _in_schema_order(given) == written


def test_a_saved_workbook_with_nothing_changed_keeps_its_tag(tmp_path: Path) -> None:
    target = tmp_path / "saved.xlsx"
    ExcelApplication.open(FIXTURES / "saved_values.xlsx", with_vba=False).save(target)
    assert _calc_pr(target) == _calc_pr(FIXTURES / "saved_values.xlsx")
