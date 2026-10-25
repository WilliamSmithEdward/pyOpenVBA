"""Renaming a sheet in a workbook that holds a table, replayed against the in-memory model.

tests/fixtures/rename_tables.json is what scripts/measure_rename_tables.py
saw in live Excel. A rename reaches every formula that names the sheet:
the cells', a defined name's, and the table's own, its calculated column
and its totals row, in the table part a save writes too. A table's
references name no sheet and stay as they were. The loaded case opens the
workbook Excel saved with such a table, tests/fixtures/
rename_tables_loaded.xlsx, renames a sheet and saves it again, as Excel
did with the same file.
"""

from __future__ import annotations

import html
import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((FIXTURES / "rename_tables.json").read_text(encoding="utf-8"))
CASES: dict[str, dict[str, Any]] = RECORD["cases"]


def saved_parts(path: Path) -> dict[str, list[str]]:
    """The saved workbook's table formulas, calculated and totals, and its defined names, as the file spells them."""
    with zipfile.ZipFile(path) as package:
        table = package.read("xl/tables/table1.xml").decode("utf-8")
        workbook = package.read("xl/workbook.xml").decode("utf-8")
    return {"calculated": [html.unescape(one) for one in
                           re.findall(r"<calculatedColumnFormula>(.*?)</calculatedColumnFormula>", table)],
            "totals": [html.unescape(one) for one in re.findall(r"<totalsRowFormula>(.*?)</totalsRowFormula>", table)],
            "names": [html.unescape(one) for one in re.findall(r"<definedName\b[^>]*>(.*?)</definedName>", workbook)]}


@pytest.mark.parametrize("name", list(CASES))
def test_a_rename_reaches_the_formulas_excel_rewrites(tmp_path: Path, name: str) -> None:
    """The sheets' and the table's names, the table's range, each cell's formula and value, and the saved table
    part's and names' formulas, as Excel had them after the rename."""
    case = CASES[name]
    path = tmp_path / f"{name}.xlsx"
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case["code"], name="Probe")
    assert app.run(f"Case_{name}", str(path)) == case["answer"]
    assert saved_parts(path) == case["saved"]


def test_a_workbook_excel_saved_with_a_table_renames_as_excel_renames_it(tmp_path: Path) -> None:
    source = tmp_path / "rename_tables_loaded.xlsx"
    shutil.copy(FIXTURES / "rename_tables_loaded.xlsx", source)
    saved = tmp_path / "renamed.xlsx"
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["loaded"]["code"], name="Probe")
    assert app.run("Rename", str(source), str(saved)) == RECORD["loaded"]["answer"]
    assert saved_parts(saved) == RECORD["loaded"]["saved"]
