"""Native Excel array assignment compatibility."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures"
RECORDS = [dict(record, name=style + "_" + record["name"])
           for style, filename in (("r1c1", "formula_arrays.json"), ("a1", "formula_a1_arrays.json"))
           for record in json.loads((FIXTURES / filename).read_text())]


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_formula_array(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim cell As Object, n As Long\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("name", ["row_matrix", "tiled", "errors", "lower_bounds"])
@pytest.mark.parametrize("style", ["a1", "r1c1"])
def test_formula_array_round_trip(tmp_path: Path, name: str, style: str) -> None:
    record = next(r for r in RECORDS if r["name"] == style + "_" + name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Public Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    expected = app.run("Build")
    path = tmp_path / "formula_array.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    # Only read after reopening: never reassign the formulas being tested.
    report = body[body.index('Report = CStr(n)'):]
    reopened.add_module('Public Function Report() As String\nDim cell As Object, n As Long\n'
                        + report + 'End Function', name="Probe")
    assert reopened.run("Report") == expected


@pytest.mark.parametrize("property_name", ["Formula", "FormulaR1C1"])
def test_formula_array_read_includes_trailing_blank_cells(property_name: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Function Report() As String
Dim a As Variant
Range("B2").Formula = "=1+2"
a = Range("B2:D4").{property_name}
Report = CStr(UBound(a, 1)) & "|" & CStr(UBound(a, 2)) & "|" & a(1, 1) & "|" & a(3, 3)
End Function''', name="Probe")
    assert app.run("Report") == "3|3|=1+2|"
