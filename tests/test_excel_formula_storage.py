"""Shared and array formulas in the file, replayed against the in-memory model.

tests/fixtures/formula_storage/ is what scripts/measure_formula_storage.py
saw: a workbook for each way of writing formulas and for each edit made to
one of them, as Excel saved it, and in formula_storage.json the formula
cells Excel wrote and every formula, with what it shows, that Excel read
back from each file through the reader the JSON keeps.

The model reads each file's formulas as Excel reads them, works them out
again as Excel does after an edit, and saves each case with its formulas
stored as Excel stored them: which are shared or array formulas, over
which block, in which order. The values a cell carries are left out of
that comparison: a save writes every formula's value, but a cell's <v>
follows a spelling rule for numbers the model does not have yet.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "formula_storage"
RECORD: dict[str, Any] = json.loads((FIXTURES / "formula_storage.json").read_text(encoding="utf-8"))
CASES: dict[str, dict[str, Any]] = {case["name"]: case for case in RECORD["cases"]}
EDITS = [name for name, case in CASES.items() if case["base"]]
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
    body = "\n".join([RECORD["reader"], "Public Function Probe() As String", "Dim ws As Object",
                      "Set ws = ActiveWorkbook.Worksheets(1)", action, f"Probe = {reads or chr(34) * 2}",
                      "End Function"])
    app.add_module(body + "\n", name="Probe")
    return str(app.run("Probe"))


def _read(text: str) -> dict[str, str]:
    """What the reader gave: each formula cell's formula and shown text, "formula=text", by address."""
    return dict(one.split("=", 1) for one in text.split(";") if one)


def _param(name: str) -> Any:
    marks = [pytest.mark.xfail(reason=KNOWN[name], strict=True)] if name in KNOWN else []
    return pytest.param(name, marks=marks, id=name)


@pytest.mark.parametrize("name", list(CASES))
def test_the_model_reads_each_formula_excel_saved(name: str) -> None:
    app = ExcelApplication.open(FIXTURES / f"{name}.xlsx", with_vba=False)
    assert _read(_run(app, "", "Formulas(ws)")) == CASES[name]["read"]


@pytest.mark.parametrize("name", EDITS)
def test_an_edit_is_worked_out_as_excel_works_it_out(name: str) -> None:
    app = ExcelApplication.open(FIXTURES / f"{CASES[name]['base']}.xlsx", with_vba=False)
    assert _read(_run(app, CASES[name]["action"], "Formulas(ws)")) == CASES[name]["read"]


def test_a_formula_and_its_text_keep_their_quotes_as_excel_writes_them(tmp_path: Path) -> None:
    """Between tags Excel escapes <, & and >, and leaves a quote as it is."""
    app = ExcelApplication()
    app.add_workbook()
    assert _run(app, 'ws.Range("S1").Formula = "=""a""""b<&>"""', 'ws.Range("S1").Value') == 'a"b<&>'
    saved = _sheet_xml(app.save(tmp_path / "quotes.xlsx"))
    wanted = next(cell for cell in CASES["kinds"]["cells"] if cell.startswith('<c r="S1"'))
    assert re.search(r'<c r="S1".*?</c>', saved).group() == wanted  # type: ignore[union-attr]


def test_a_save_works_out_a_formula_nothing_has_read(tmp_path: Path) -> None:
    """Under automatic calculation Excel's cells are always up to date, so a formula written and never read is
    saved with its value, as the totals rows in tests/fixtures/tables/totals_row/ are."""
    app = ExcelApplication()
    app.add_workbook()
    _run(app, 'ws.Range("A1").Value = 2\nws.Range("B1").Formula = "=A1*3"', '""')
    found = re.search(r'<c r="B1".*?</c>', _sheet_xml(app.save(tmp_path / "unread.xlsx")))
    assert found is not None and found.group() == '<c r="B1"><f>A1*3</f><v>6</v></c>'


def _patched(app: ExcelApplication, tmp_path: Path, old: str, new: str) -> ExcelApplication:
    """The workbook saved, its first sheet's ``old`` replaced by ``new`` as Excel would have written it, opened."""
    path = app.save(tmp_path / "made.xlsx")
    with zipfile.ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert old in sheet
    parts["xl/worksheets/sheet1.xml"] = sheet.replace(old, new).encode("utf-8")
    patched = tmp_path / "patched.xlsx"
    with zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)
    return ExcelApplication.open(patched, with_vba=False)


def test_a_formula_the_parser_cannot_read_keeps_its_value_and_stops_no_other(tmp_path: Path) -> None:
    """A formula the parser cannot read, here a ragged array constant a file from another program might hold,
    keeps the value the file gave it, and the rest of the workbook is worked out as ever."""
    app = ExcelApplication()
    app.add_workbook()
    _run(app, 'ws.Range("A1").Formula = "=1+1"\nws.Range("A2").Value = 2\nws.Range("B1").Formula = "=A2*2"')
    opened = _patched(app, tmp_path, "<f>1+1</f><v>2</v>", "<f>SUM({1,2;3})</f><v>8</v>")
    reads = 'ws.Range("A1").Value & "|" & ws.Range("B1").Value'
    assert _run(opened, 'ws.Range("A2").Value = 5', reads) == "8|10"


def test_a_formula_whose_answer_is_empty_text_reads_back_as_text(tmp_path: Path) -> None:
    """Excel saves a formula's empty text as <v/>, as pyOfficeEditor's formula corpus has MID("hello",9,2); under
    manual calculation the cell shows what the file cached, which is text."""
    app = ExcelApplication()
    app.add_workbook()
    _run(app, 'ws.Range("A1").Formula = "=1+1"')
    opened = _patched(app, tmp_path, '<c r="A1"><f>1+1</f><v>2</v></c>',
                      '<c r="A1" t="str"><f>MID("hello",9,2)</f><v/></c>')
    reads = 'TypeName(ws.Range("A1").Value) & "|" & ws.Range("A1").Value'
    assert _run(opened, "Application.Calculation = xlCalculationManual", reads) == "String|"


#: Formulas that name a sheet, written to blocks and cut to another sheet (scripts/measure_shared_sheets.py).
SHEETS: dict[str, Any] = json.loads((FIXTURES / "sheets.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def named_sheets(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, str]:
    """The measurement's workbook made in the model, and its first sheet as the model saves it."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Make()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n" + SHEETS["setup"]
                   + "\n" + SHEETS["writes"] + "End Sub\n", name="Make")
    app.run("Make")
    return app, _sheet_xml(app.save(tmp_path_factory.mktemp("sheets") / "sheets.xlsx"))


@pytest.mark.parametrize("cell", list(SHEETS["cells"]))
def test_a_formula_naming_a_sheet_is_stored_as_excel_stores_it(named_sheets: tuple[ExcelApplication, str],
                                                               cell: str) -> None:
    """No shared formula names a sheet, not even its own, and a group a cut sends a formula of to another sheet is
    written a formula to each cell, the rest of the group's too: Excel refuses to open a file with one."""
    found = re.search(rf'<c r="{cell}".*?</c>', named_sheets[1])
    assert found is not None and found.group() == SHEETS["cells"][cell]


@pytest.mark.parametrize("cell", list(SHEETS["formulas"]))
def test_a_formula_naming_a_sheet_reads_as_excel_reads_it(named_sheets: tuple[ExcelApplication, str],
                                                          cell: str) -> None:
    assert named_sheets[0].sheet(1).formula(cell) == SHEETS["formulas"][cell]


@pytest.mark.parametrize("name", [_param(name) for name in CASES])
def test_a_save_stores_the_formulas_as_excel_stores_them(name: str, tmp_path: Path) -> None:
    if CASES[name]["base"]:
        app = ExcelApplication.open(FIXTURES / f"{CASES[name]['base']}.xlsx", with_vba=False)
    else:
        app = ExcelApplication()
        app.add_workbook()
    _run(app, CASES[name]["action"])
    saved = app.save(tmp_path / f"{name}.xlsx")
    assert _formula_cells(_sheet_xml(saved)) == _formula_cells(_sheet_xml(FIXTURES / f"{name}.xlsx"))
