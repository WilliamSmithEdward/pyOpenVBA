"""An @ written in a formula, and a call of SINGLE, replayed against live Excel.

tests/fixtures/at_sign.json is what scripts/measure_at_sign.py saw
writing formulas with an @ or a call of SINGLE through Formula, Formula2,
FormulaR1C1 and Value. Written through Formula, an @ where the formula
cuts to one value anyway is left out, =@A1:A3 reading back =A1:A3, and
one kept, =@A1, has Formula read back every @ Formula2 shows,
=@A1+A1:A3 as =@A1+@A1:A3; SINGLE(x) is written @x, its argument as it
stands. Formula2 writes a formula as Formula does where Formula would
read back the same, and as a dynamic array otherwise. Then a workbook of
such formulas was saved: a file keeps each @ as _xlfn.SINGLE, the
model's save is held to Excel's bytes cell by cell, and the workbook
Excel saved opens with each formula as Excel read it.
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
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

FIXTURE = Path(__file__).parent / "fixtures" / "at_sign.json"
WORKBOOK = Path(__file__).parent / "fixtures" / "at_sign.xlsx"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
CASES: dict[str, dict[str, str]] = {f"{case['member']} {case['formula']}": case for case in RECORD["cases"]}
FILE: dict[str, Any] = RECORD["file"]
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)

#: Cases the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {
    "FormulaR1C1 =@A:A": "R1C1 reads A:A as a range between two names, which Excel shows A:(A); the model reads a "
                         "column",
}


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Setup()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n" + RECORD["setup"]
                   + "\nEnd Sub\nPublic Sub Assign(ByVal member As String, ByVal f As String)\n"
                   'Range("C8").ClearContents\nSelect Case member\nCase "Formula": Range("C8").Formula = f\n'
                   'Case "Formula2": Range("C8").Formula2 = f\nCase "FormulaR1C1": Range("C8").FormulaR1C1 = f\n'
                   'Case "Formula2R1C1": Range("C8").Formula2R1C1 = f\nCase "Value": Range("C8").Value = f\n'
                   "End Select\nEnd Sub\nPublic Function Shown() As String\n"
                   'Shown = Range("C8").Formula & "~" & Range("C8").Formula2 & "~" & Range("C8").FormulaR1C1 & "~" & '
                   'Range("C8").Text\nEnd Function\n', name="Setup")
    app.run("Setup")
    return app


@pytest.mark.parametrize("name", [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in CASES])
def test_a_formula_with_an_at_reads_back_as_in_excel(app: ExcelApplication, name: str) -> None:
    case = CASES[name]
    try:
        app.run("Assign", case["member"], case["formula"])
        found = str(app.run("Shown"))
    except VBARuntimeError as failure:
        found = f"!{failure.number}"
    except VBAUnsupportedError as gap:
        found = f"unsupported: {gap}"
    assert found == case["answer"]


def test_every_gap_is_a_case() -> None:
    assert set(GAPS) <= set(CASES)


def _cells(xml: str) -> dict[str, str]:
    """Each <c> of a sheet's XML, by its cell."""
    return {match.group(1): match.group(0) for match in _CELL.finditer(xml)}


@pytest.fixture(scope="module")
def saved() -> dict[str, str]:
    """The parts of the workbook the model saves after the file's writes, as Excel saved them."""
    writes = "\n".join(f'ws.Range("{cell}").{member} = "{formula.replace(chr(34), chr(34) * 2)}"'
                       for cell, member, formula in FILE["writes"])
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Writes()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n" + RECORD["setup"]
                   + "\n" + writes + "\nEnd Sub\n", name="Writes")
    app.run("Writes")
    with tempfile.TemporaryDirectory() as folder:
        path = app.save(Path(folder) / "at_sign.xlsx")
        with zipfile.ZipFile(path) as package:
            return {name: package.read(name).decode("utf-8") for name in package.namelist() if name.endswith(".xml")}


@pytest.mark.parametrize("cell", list(_cells(FILE["parts"]["xl/worksheets/sheet1.xml"])))
def test_a_save_writes_each_cell_as_excel_does(saved: dict[str, str], cell: str) -> None:
    excel = _cells(FILE["parts"]["xl/worksheets/sheet1.xml"])
    assert _cells(saved["xl/worksheets/sheet1.xml"]).get(cell) == excel[cell]


def test_a_save_writes_no_cell_excel_did_not(saved: dict[str, str]) -> None:
    assert set(_cells(saved["xl/worksheets/sheet1.xml"])) == set(_cells(FILE["parts"]["xl/worksheets/sheet1.xml"]))


def test_a_save_writes_the_dynamic_array_metadata_as_excel_does(saved: dict[str, str]) -> None:
    assert saved["xl/metadata.xml"] == FILE["parts"]["xl/metadata.xml"]


@pytest.fixture(scope="module")
def opened() -> ExcelApplication:
    return ExcelApplication.open(WORKBOOK, with_vba=False)


@pytest.mark.parametrize("cell", list(FILE["reads"]))
def test_the_workbook_excel_saved_opens_with_each_formula(opened: ExcelApplication, cell: str) -> None:
    found = opened.evaluate(f'Range("{cell}").Formula & "~" & Range("{cell}").Formula2 & "~" & Range("{cell}").Text')
    assert str(found) == FILE["reads"][cell]
