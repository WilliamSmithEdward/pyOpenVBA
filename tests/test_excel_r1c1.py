"""R1C1 formula behavior measured in native Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.formula._r1c1 import from_a1, to_a1

RECORDS = json.loads((Path(__file__).parent / "fixtures/formula_r1c1.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_r1c1(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim cell As Object, n As Long\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    assert app.run("Report") == record["reported"]


def test_r1c1_persistence_and_calculation(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Range("A1:A3").Value = 7
Range("B1:B3").FormulaR1C1 = "=RC[-1]*2"
Application.Calculate
End Sub''', name="BuildModule")
    app.run("Build")
    path = tmp_path / "r1c1.xlsx"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module('''Function Report() As String
Dim formulas As Variant
formulas = Range("B1:B4").FormulaR1C1
Report = formulas(3, 1) & "|" & CStr(Range("B3").Value) & "|" & CStr(UBound(formulas, 1)) & "|" & formulas(4, 1)
End Function''', name="Probe")
    assert reopened.run("Report") == "=RC[-1]*2|14|4|"


def test_conversion_preserves_names_strings_and_quoted_sheets() -> None:
    formula = '=SUM(\'R1C1 data\'!RC[-1],Revenue,"R1C1 ""RC""")'
    a1 = '=SUM(\'R1C1 data\'!B3,Revenue,"R1C1 ""RC""")'
    assert to_a1(formula, 3, 3) == a1
    assert from_a1(a1, 3, 3) == formula


def test_r1c1_respects_merged_cells() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Report() As String
Range("B2:C3").Merge
Range("B2:C3").FormulaR1C1 = "=RC[-1]"
Report = Range("B2").Formula & "|" & Range("C3").FormulaR1C1
End Function''', name="Probe")
    assert app.run("Report") == "=A2|"
