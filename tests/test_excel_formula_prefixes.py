"""How a file spells newer functions, the names a LET binds, and formulas worked out whenever anything changes.

tests/fixtures/formula_prefixes/ is what scripts/measure_formula_prefixes.py
saw in live Excel: formulas written through Range.Formula, saved, and
read back; a call of each of Excel's functions, saved; and volatile
formulas, shared and array ones among them, saved beside a macro's own
function. The model writes each formula as Excel wrote it, _xlfn.,
_xlfn._xlws., _xlpm. and ca="1" included, and reads Excel's file back as
Range.Formula spells it.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "formula_prefixes"
RECORD: dict[str, Any] = json.loads((FIXTURES / "prefixes.json").read_text(encoding="utf-8"))
FORMULAS: list[dict[str, str]] = RECORD["formulas"]
FUNCTIONS: dict[str, dict[str, str]] = {name: one for name, one in RECORD["functions"].items() if one["saved"]}
VOLATILE: dict[str, Any] = RECORD["volatile"]
_FORMULA = re.compile(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", re.DOTALL)
_CELL = re.compile(r'<c r="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)', re.DOTALL)

#: Formulas the model does not read back as Range.Formula spells them, and why.
READ_GAPS: dict[str, str] = {}


def _formulas(path: Path) -> dict[str, str]:
    """The <f> element of each cell of the first sheet that has one, by address."""
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found: dict[str, str] = {}
    for address, body in _CELL.findall(xml):
        element = _FORMULA.search(body or "")
        if element is not None:
            found[address] = element.group()
    return found


def _element(cell: str) -> str:
    found = _FORMULA.search(cell)
    return found.group() if found else ""


def _written(tmp_path: Path, formulas: list[tuple[str, str]], name: str = "written.xlsx", *,
             module: str = "", writes: str = "") -> dict[str, str]:
    """The <f> elements of a workbook the model saves after writing ``formulas``, each an address and its formula."""
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Probe()", "Dim ws As Object", "Set ws = ActiveWorkbook.Worksheets(1)",
             'ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))']
    lines += [f'ws.Range("{address}").Formula = "{formula.replace(chr(34), chr(34) * 2)}"'
              for address, formula in formulas]
    app.add_module(module + "\n".join(lines) + "\n" + writes.replace("ThisWorkbook", "ActiveWorkbook") + "End Sub\n",
                   name="Probe")
    app.run("Probe")
    return _formulas(app.save(tmp_path / name))


def _cases(records: list[dict[str, str]], gaps: dict[str, str]) -> list[Any]:
    return [pytest.param(record, id=record["written"],
                         marks=[pytest.mark.xfail(reason=gaps[record["written"]], strict=True)]
                         if record["written"] in gaps else [])
            for record in records]


@pytest.fixture(scope="module")
def saved(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    """Every formula of the fixture written into its cell, as the model saves them."""
    return _written(tmp_path_factory.mktemp("prefixes"), [(record["cell"], record["written"]) for record in FORMULAS])


@pytest.mark.parametrize("record", _cases(FORMULAS, {}))
def test_a_save_spells_each_formula_as_excel_does(saved: dict[str, str], record: dict[str, str]) -> None:
    assert saved.get(record["cell"], "") == _element(record["saved"])


@pytest.fixture(scope="module")
def opened() -> ExcelApplication:
    return ExcelApplication.open(FIXTURES / "prefixes.xlsx", with_vba=False)


@pytest.mark.parametrize("record", _cases(FORMULAS, READ_GAPS))
def test_the_model_reads_each_formula_excel_saved(opened: ExcelApplication, record: dict[str, str]) -> None:
    assert opened.evaluate(f'Range("{record["cell"]}").Formula') == record["formula"]


_NARROW = "a General number is cut to the column's width in pixels, which Range.Text does not measure yet"
#: Formulas whose cell does not show what Excel's shows, and why.
TEXT_GAPS: dict[str, str] = {"=NORM.DIST(1,0,1,TRUE)": _NARROW, "=STDEV.S(A1:A4)": _NARROW}


@pytest.mark.parametrize("record", _cases(FORMULAS, TEXT_GAPS))
def test_each_formula_excel_saved_shows_what_excel_showed(opened: ExcelApplication, record: dict[str, str]) -> None:
    assert opened.evaluate(f'Range("{record["cell"]}").Text') == record["text"]


@pytest.fixture(scope="module")
def every_function(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    """A call of each of Excel's functions Excel took, one to a row of column C, as the model saves them."""
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.set_value("A1", 1)
    for row, one in enumerate(FUNCTIONS.values(), start=1):
        sheet.set_value(f"C{row}", "=" + one["call"])
    return _formulas(app.save(tmp_path_factory.mktemp("functions") / "functions.xlsx"))


@pytest.mark.parametrize("name", sorted(FUNCTIONS))
def test_a_save_spells_every_function_as_excel_does(every_function: dict[str, str], name: str) -> None:
    row = list(FUNCTIONS).index(name) + 1
    assert every_function.get(f"C{row}", "") == FUNCTIONS[name]["saved"]


def test_volatile_formulas_are_saved_as_excel_saves_them(tmp_path: Path) -> None:
    """Shared and array formulas, a macro's function, a function no one has, a named LAMBDA and INFO."""
    found = _written(tmp_path, [], "volatile.xlsm", module=VOLATILE["udf"] + "\n", writes=VOLATILE["writes"])
    assert found == VOLATILE["cells"]


def _twice(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read("xl/workbook.xml").decode("utf-8")
    found = re.search(r'<definedName name="Twice"[^>]*>(.*?)</definedName>', xml)
    return found.group(1) if found else ""


def test_a_defined_name_is_spelled_as_a_formula_is(tmp_path: Path) -> None:
    """Names.Add "Twice", "=LAMBDA(x,x*2)" saves as Excel saved it in the volatile workbook, and reads back bare."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Probe()\nActiveWorkbook.Names.Add "Twice", "=LAMBDA(x,x*2)"\nEnd Sub\n', name="Probe")
    app.run("Probe")
    saved = app.save(tmp_path / "twice.xlsx")
    assert _twice(saved) == VOLATILE["names"]["Twice"]
    assert ExcelApplication.open(saved, with_vba=False).evaluate(
        'ActiveWorkbook.Names("Twice").RefersTo') == "=LAMBDA(x,x*2)"
