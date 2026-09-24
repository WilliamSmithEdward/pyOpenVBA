"""Dynamic arrays that spill, replayed against live Excel.

tests/fixtures/dynamic_arrays.json is what
scripts/measure_dynamic_arrays.py saw: a formula written through Formula2
that answers with an array spills into the cells below and beside it,
its first cell holding the formula and the rest only values; a cell in
the way makes it #SPILL!; A1# reads what spilled from A1; and a legacy
formula is cut to one value. Each case is run as the measurement ran it
and every cell it asked about is compared: Formula, Formula2 and their
R1C1, Value, Text, HasFormula, HasSpill, HasArray, SpillParent,
SpillingToRange and CurrentArray.
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

FIXTURE = Path(__file__).parent / "fixtures" / "dynamic_arrays.json"
WORKBOOK = Path(__file__).parent / "fixtures" / "dynamic_arrays.xlsx"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
CASES: dict[str, dict[str, str]] = {case["name"]: case for case in RECORD["cases"]}
EXPRESSIONS: dict[str, dict[str, str]] = RECORD["expressions"]
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
#: The parts a file keeps its dynamic arrays in, besides the sheet.
PARTS = ("xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/rdrichvaluestructure.xml",
         "xl/richData/rdRichValueTypes.xml")
#: The cells the file's writes made, but for I2's shared string and the merged Q2:R2's format.
WRITTEN = ("E1", "E2", "E3", "F1", "F2", "F3", "G1", "H1", "I1", "J1", "K1", "L1", "K3", "M1", "N1", "O1", "P1", "O2",
           "P2", "Q1", "S1048575", "T1", "T2", "T3", "U1")

#: Cases the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {}


def _run(body: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\nPublic Function Report() As String\nDim ws As Object\n" + body
                   + "End Function\n", name="Probe")
    return str(app.run("Report"))


@pytest.mark.parametrize("name", [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in CASES])
def test_a_dynamic_array_behaves_as_in_excel(name: str) -> None:
    assert _run(CASES[name]["body"]) == CASES[name]["reported"]


@pytest.mark.parametrize(("name", "expression"), [
    (name, expression) for name, asked in EXPRESSIONS.items() for expression in asked if expression != "writes"])
def test_a_spill_reads_through_ranges_as_in_excel(name: str, expression: str) -> None:
    case = EXPRESSIONS[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Function Report() As String\nDim ws As Object, out As String, v As Variant\n"
                   "Set ws = ActiveWorkbook.Worksheets.Add\n"
                   'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n' + case["writes"] + "\n"
                   "On Error Resume Next\nv = Empty\n" + f"v = {expression}\n"
                   'If Err.Number <> 0 Then out = "!" & Err.Number Else out = TypeName(v) & ":" & CStr(v)\n'
                   "Report = out\nEnd Function\n", name="Probe")
    assert str(app.run("Report")) == case[expression]


def _cells(xml: str) -> dict[str, str]:
    """Each <c> of a sheet's XML, by its cell."""
    return {match.group(1): match.group(0) for match in _CELL.finditer(xml)}


@pytest.fixture(scope="module")
def saved() -> dict[str, str]:
    """The parts of the workbook the model saves after the file's writes, as Excel saved them."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Writes()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n"
                   'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n' + RECORD["file"]["writes"]
                   + "End Sub\n", name="Writes")
    app.run("Writes")
    with tempfile.TemporaryDirectory() as folder:
        path = app.save(Path(folder) / "spills.xlsx")
        with zipfile.ZipFile(path) as package:
            return {name: package.read(name).decode("utf-8") for name in package.namelist()
                    if name.endswith((".xml", ".rels"))}


@pytest.mark.parametrize("cell", WRITTEN)
def test_a_save_writes_each_dynamic_array_cell_as_excel_does(saved: dict[str, str], cell: str) -> None:
    excel = _cells(RECORD["file"]["parts"]["xl/worksheets/sheet1.xml"])
    assert _cells(saved["xl/worksheets/sheet1.xml"])[cell] == excel[cell]


@pytest.mark.parametrize("part", PARTS)
def test_a_save_writes_the_metadata_of_dynamic_arrays_as_excel_does(saved: dict[str, str], part: str) -> None:
    assert saved[part] == RECORD["file"]["parts"][part]


def test_a_save_of_one_spill_writes_only_the_dynamic_array_metadata() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Writes()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n"
                   'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n'
                   + RECORD["plain_file"]["writes"] + "End Sub\n", name="Writes")
    app.run("Writes")
    with tempfile.TemporaryDirectory() as folder:
        path = app.save(Path(folder) / "spill.xlsx")
        with zipfile.ZipFile(path) as package:
            names = package.namelist()
            metadata = package.read("xl/metadata.xml").decode("utf-8")
    assert metadata == RECORD["plain_file"]["parts"]["xl/metadata.xml"]
    assert not any("richData" in name for name in names)


#: What the workbook Excel saved holds, each cell's answer to one expression.
OPENED = {
    'Range("E1").Formula': "=SEQUENCE(3)", 'Range("E1").HasSpill': "True", 'Range("E2").Value': "2",
    'Range("E3").SpillParent.Address': "$E$1", 'Range("E1").SpillingToRange.Address': "$E$1:$E$3",
    'Range("E2").Formula': "", 'Range("F3").Value': "6", 'Range("G1").Value': "6", 'Range("I1").Text': "#SPILL!",
    'Range("I1").HasSpill': "False", 'Range("L1").Value': "b", 'Range("K3").Formula': "=SUM(E1#)",
    'Range("N1").Text': "#CALC!", 'Range("O1").SpillingToRange.Address': "$O$1:$P$2", 'Range("P2").Value': "4",
    'Range("Q1").Text': "#SPILL!", 'Range("T2").Text': "#DIV/0!", 'Range("U1").Formula2': "=SEQUENCE(1)/0",
    'Range("H1").Formula2': "=@A1:A3",
}


@pytest.mark.parametrize("expression", list(OPENED))
def test_a_workbook_excel_saved_opens_with_its_dynamic_arrays(expression: str) -> None:
    app = ExcelApplication.open(WORKBOOK, with_vba=False)
    assert str(app.evaluate(f"CStr({expression})")) == OPENED[expression]


def test_a_spill_the_model_saved_opens_again() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Writes()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n"
                   'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n' + RECORD["file"]["writes"]
                   + "End Sub\n", name="Writes")
    app.run("Writes")
    with tempfile.TemporaryDirectory() as folder:
        again = ExcelApplication.open(app.save(Path(folder) / "spills.xlsx"), with_vba=False)
    for expression, answer in OPENED.items():
        assert str(again.evaluate(f"CStr({expression})")) == answer, expression
