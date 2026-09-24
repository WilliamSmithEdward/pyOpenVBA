"""The names a LET or a LAMBDA binds, replayed against live Excel.

tests/fixtures/bound_names.json is what scripts/measure_bound_names.py
saw: LET(x1,5,x1) with x1 a name, not the cell, which A1 cannot read and
Excel reads as R1C1, every other cell written as A1 becoming a name
there too, 'A1'; R and C as names; and LAMBDA's [optional] parameters,
a call leaving out any other being #VALUE!. Each formula is written
through Range.Formula and read back through Formula, Formula2 and Text,
the workbook saved with each <f> element as Excel's, and opened again.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE = Path(__file__).parent / "fixtures" / "bound_names.json"
CASES: list[dict[str, Any]] = json.loads(FIXTURE.read_text(encoding="utf-8"))
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>(.*?)</c>)', re.DOTALL)
_FORMULA = re.compile(r"<f\b[^>]*>.*?</f>|<f\b[^>]*/>", re.DOTALL)
_READ = ('Public Function Shown(ByVal address As String) As String\n'
         'Shown = Range(address).Formula & "~" & Range(address).Formula2 & "~" & Range(address).Text\nEnd Function\n'
         'Public Function Again(ByVal address As String) As String\n'
         'Again = Range(address).Formula & "~" & Range(address).Text\nEnd Function\n')


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Setup()\nRange("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))\nEnd Sub\n'
                   + _READ, name="Probe")
    app.run("Setup")
    sheet = app.sheet(1)
    for case in CASES:
        sheet.set_value(case["cell"], case["written"])
    return app


@pytest.fixture(scope="module")
def saved(app: ExcelApplication) -> dict[str, str]:
    """The <f> element of each formula cell the model saves."""
    with tempfile.TemporaryDirectory() as folder:
        path = app.save(Path(folder) / "bound.xlsm")
        with zipfile.ZipFile(path) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found: dict[str, str] = {}
    for address, body in _CELL.findall(xml):
        element = _FORMULA.search(body or "")
        if element is not None:
            found[address] = element.group()
    return found


@pytest.fixture(scope="module")
def opened(app: ExcelApplication) -> ExcelApplication:
    """The workbook the model saved, opened again."""
    with tempfile.TemporaryDirectory() as folder:
        again = ExcelApplication.open(app.save(Path(folder) / "bound.xlsx"), with_vba=False)
    again.add_module(_READ, name="Reading")
    return again


@pytest.mark.parametrize("case", CASES, ids=[case["written"] for case in CASES])
def test_a_bound_name_reads_back_as_in_excel(app: ExcelApplication, case: dict[str, Any]) -> None:
    assert app.run("Shown", case["cell"]) == case["read"]


@pytest.mark.parametrize("case", CASES, ids=[case["written"] for case in CASES])
def test_a_bound_name_is_saved_as_excel_saves_it(saved: dict[str, str], case: dict[str, Any]) -> None:
    assert saved[case["cell"]] == case["saved"]


@pytest.mark.parametrize("case", CASES, ids=[case["written"] for case in CASES])
def test_a_bound_name_opens_again_as_in_excel(opened: ExcelApplication, case: dict[str, Any]) -> None:
    assert opened.run("Again", case["cell"]) == case["opened"]
