"""Shared formulas in the file, replayed against the in-memory model.

tests/fixtures/formula_storage/ is what scripts/measure_formula_storage.py
saw: a workbook for each way of writing formulas and for each edit made to
the "written" one, as Excel saved it, and in formula_storage.json the
formula cells Excel wrote and every formula, with what it shows, that
Excel read back from each file.

The model reads each file's formulas as Excel reads them, works them out
again as Excel does after an edit, and saves each case with its formulas
stored as Excel stored them: which are shared, over which block, in which
order. The values a cell carries are left out of that comparison, since
the model works a formula out only when something asks.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import column_letter
from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "formula_storage"
RECORD: dict[str, Any] = json.loads((FIXTURES / "formula_storage.json").read_text(encoding="utf-8"))
CASES: dict[str, dict[str, Any]] = {case["name"]: case for case in RECORD["cases"]}
EDITS: list[str] = RECORD["edits"]
WRITTEN = [name for name in CASES if name not in EDITS]
_FORMULA_CELL = re.compile(r'<c r="([A-Z]+\d+)"[^>]*>(?:(?!</c>).)*?(<f\b[^>]*/>|<f\b[^>]*>.*?</f>)', re.DOTALL)

#: Cases the model stores differently, and why.
KNOWN: dict[str, str] = {}


def _formula_cells(xml: str) -> list[tuple[str, str]]:
    """Each formula cell's address and ``<f>`` element, in the order the sheet holds them."""
    return _FORMULA_CELL.findall(xml)


def _sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _run(app: ExcelApplication, action: str, reads: str = "") -> str:
    """Run a case's VBA on the first sheet, then hand back what ``reads`` reads, if anything."""
    body = "\n".join(["Public Function Probe() As String", "Dim ws As Object",
                      "Set ws = ActiveWorkbook.Worksheets(1)", action, f"Probe = {reads or chr(34) * 2}",
                      "End Function"])
    app.add_module(body + "\n", name="Probe")
    return str(app.run("Probe"))


def _param(name: str) -> Any:
    marks = [pytest.mark.xfail(reason=KNOWN[name], strict=True)] if name in KNOWN else []
    return pytest.param(name, marks=marks, id=name)


@pytest.mark.parametrize("name", list(CASES))
def test_the_model_reads_each_formula_excel_saved(name: str) -> None:
    app = ExcelApplication.open(FIXTURES / f"{name}.xlsx", with_vba=False)
    found = {f"{column_letter(column)}{row}": cell.formula
             for (row, column), cell in app.workbook.sheets_[0].cells_.items() if cell.formula}
    # Each read-back is the formula and the text it shows, joined by "=".
    assert found == {address: shown.rsplit("=", 1)[0] for address, shown in CASES[name]["read"].items()}


@pytest.mark.parametrize("name", EDITS)
def test_an_edit_is_worked_out_as_excel_works_it_out(name: str) -> None:
    app = ExcelApplication.open(FIXTURES / "written.xlsx", with_vba=False)
    reads = " & ".join(f'"{address}=" & ws.Range("{address}").Formula & "=" & ws.Range("{address}").Text & ";"'
                       for address in CASES[name]["read"]) or chr(34) * 2
    got = dict(one.split("=", 1) for one in _run(app, CASES[name]["action"], reads).split(";") if one)
    assert got == CASES[name]["read"]


def test_a_formula_and_its_text_keep_their_quotes_as_excel_writes_them(tmp_path: Path) -> None:
    """Between tags Excel escapes <, & and >, and leaves a quote as it is."""
    app = ExcelApplication()
    app.add_workbook()
    assert _run(app, 'ws.Range("S1").Formula = "=""a""""b<&>"""', 'ws.Range("S1").Value') == 'a"b<&>'
    saved = _sheet_xml(app.save(tmp_path / "quotes.xlsx"))
    wanted = next(cell for cell in CASES["kinds"]["cells"] if cell.startswith('<c r="S1"'))
    assert re.search(r'<c r="S1".*?</c>', saved).group() == wanted  # type: ignore[union-attr]


@pytest.mark.parametrize("name", [_param(name) for name in CASES])
def test_a_save_stores_the_formulas_as_excel_stores_them(name: str, tmp_path: Path) -> None:
    if name in EDITS:
        app = ExcelApplication.open(FIXTURES / "written.xlsx", with_vba=False)
    else:
        app = ExcelApplication()
        app.add_workbook()
    _run(app, CASES[name]["action"])
    saved = app.save(tmp_path / f"{name}.xlsx")
    assert _formula_cells(_sheet_xml(saved)) == _formula_cells(_sheet_xml(FIXTURES / f"{name}.xlsx"))
