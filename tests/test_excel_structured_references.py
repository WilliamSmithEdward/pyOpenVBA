"""Structured references, replayed against the in-memory model.

tests/fixtures/structured_references/ is what
scripts/measure_structured_references.py saw in live Excel: in
structured.json the VBA it ran and every answer it got -- whether a
formula was taken and what Formula, FormulaR1C1 and Value read back, a
block written, copied and filled, what Range and Evaluate make of a
reference, a sweep of the characters a column's name can hold, headers
overwritten, the table and its columns renamed -- and the two workbooks
it saved, the first before the renames and the second after.

The model runs the same VBA and has to answer the same, save the same
<f> for every formula cell, and read Excel's own files back to the
formulas Excel read.
"""

from __future__ import annotations

import json
import re
import zipfile
from functools import cache
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Range

FIXTURES = Path(__file__).parent / "fixtures" / "structured_references"
RECORD: dict[str, Any] = json.loads((FIXTURES / "structured.json").read_text(encoding="utf-8"))
ANSWERS: dict[str, str] = {key: value for key, value in RECORD["answers"].items()
                           if key != "settings" and not key.startswith("f2:")}
CASES: list[dict[str, Any]] = RECORD["cases"]
FILES: dict[str, dict[str, dict[str, str]]] = RECORD["files"]

#: Answers the model gives differently, and why.
KNOWN: dict[str, str] = {
    "char:c13": "a header holding a carriage return keeps it in the cell and in the file, but everything VBA reads "
                "of the column, its name and the formulas naming it, has a line feed there; not modelled",
    "structured/xl/worksheets/sheet1.xml/C37": "COUNTA(Big[@]) in a row of Big reads itself: Excel marks a formula "
                                               "caught in a circle ca=\"1\", and the model does not",
    "renamed/xl/worksheets/sheet1.xml/C37": "as above",
}

_PROBE = """Public Function Probe() As String
Dim wb As Object, ws As Object, ws2 As Object, ws3 As Object, out As String
Set wb = ActiveWorkbook
Set ws = wb.Worksheets(1)
Set ws2 = wb.Worksheets.Add(After:=ws)
Set ws3 = wb.Worksheets.Add(After:=ws2)
ws.Activate
Build ws, ws2
out = out & Writes(ws, ws2)
out = out & Reads(ws, ws2, "r")
out = out & Blocks(ws)
out = out & Questions(ws, ws2)
out = out & Characters(ws3)
out = out & Repeats(ws3)
Probe = out
End Function

Public Function Renaming() As String
Renaming = Renames(ActiveWorkbook.Worksheets(1), ActiveWorkbook.Worksheets(2))
End Function
"""


def _records(text: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for record in text.split("~|~"):
        if "~:~" in record:
            key, value = record.split("~:~", 1)
            out[key] = value
    return out


_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*?(?:/>|>(.*?)</c>)', re.DOTALL)
_F = re.compile(r"<f\b[^>]*?(?:/>|>.*?</f>)", re.DOTALL)


def _formulas(path: Path) -> dict[str, dict[str, str]]:
    """The <f> of every formula cell, per sheet part, as the measuring script reads them."""
    out: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith("xl/worksheets/sheet"):
                text = package.read(name).decode("utf-8")
                cells: dict[str, str] = {}
                for match in _CELL.finditer(text):
                    element = _F.search(match.group(2) or "")
                    if element is not None:
                        cells[match.group(1)] = element.group(0)
                out[name] = cells
    return out


@cache
def _model(folder: str) -> tuple[dict[str, str], dict[str, dict[str, dict[str, str]]]]:
    """What the model answers running the measured VBA, and the formulas it saves before and after the renames."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["procedures"] + _PROBE, name="Probe")
    answers = _records(str(app.run("Probe")))
    first = _formulas(app.save(Path(folder) / "structured.xlsx"))
    answers |= _records(str(app.run("Renaming")))
    second = _formulas(app.save(Path(folder) / "renamed.xlsx"))
    return answers, {"structured": first, "renamed": second}


@pytest.fixture(scope="module")
def model(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, dict[str, dict[str, str]]]]:
    return _model(str(tmp_path_factory.mktemp("structured")))


def _param(name: str) -> Any:
    marks = [pytest.mark.xfail(reason=KNOWN[name], strict=True)] if name in KNOWN else []
    return pytest.param(name, marks=marks, id=name)


@pytest.mark.parametrize("key", [_param(key) for key in ANSWERS])
def test_the_model_answers_as_excel_did(key: str, model: tuple[dict[str, str], Any]) -> None:
    assert model[0].get(key) == ANSWERS[key]


_SAVED = [(which, part, cell) for which, parts in FILES.items() for part, cells in parts.items()
          if part.startswith("xl/worksheets/") for cell in cells]


@pytest.mark.parametrize("where", [_param("/".join(one)) for one in _SAVED])
def test_a_save_writes_the_formula_excel_wrote(where: str, model: tuple[Any, dict[str, Any]]) -> None:
    which, part, cell = where.split("/", 2)[0], "/".join(where.split("/")[1:-1]), where.split("/")[-1]
    assert model[1][which].get(part, {}).get(cell) == FILES[which][part][cell]


def test_a_save_writes_no_formula_excel_did_not(model: tuple[Any, dict[str, Any]]) -> None:
    for which, parts in model[1].items():
        for part, cells in parts.items():
            extra = set(cells) - set(FILES[which].get(part, {}))
            assert not extra, f"{which} {part}: {sorted(extra)}"


def test_renaming_a_table_read_from_a_file_saves_what_excel_saved(tmp_path: Path) -> None:
    """Excel's first workbook, renamed as Excel renamed it, saves the formulas and table parts of its second."""
    app = ExcelApplication.open(FIXTURES / "structured.xlsx", with_vba=False)
    app.add_module(RECORD["procedures"] + _PROBE, name="Probe")
    answers = _records(str(app.run("Renaming")))
    assert {key: value for key, value in answers.items() if value != ANSWERS[key]} == {}
    saved = app.save(tmp_path / "renamed.xlsx")
    with zipfile.ZipFile(saved) as ours, zipfile.ZipFile(FIXTURES / "renamed.xlsx") as excels:
        parts = sorted(name for name in excels.namelist() if name.startswith("xl/tables/"))
        assert sorted(name for name in ours.namelist() if name.startswith("xl/tables/")) == parts
        for name in parts:
            assert ours.read(name).decode("utf-8") == excels.read(name).decode("utf-8"), name
    ours = _formulas(saved)
    for part, cells in FILES["renamed"].items():
        if part.startswith("xl/worksheets/"):
            mine = ours.get(part, {})
            differ = {cell for cell in set(cells) | set(mine) if cells.get(cell) != mine.get(cell)}
            # C37 is the formula caught in a circle that Excel marks ca="1" (see KNOWN).
            assert differ <= ({"C37"} if part.endswith("sheet1.xml") else set()), (part, sorted(differ))


def _formula_of(answer: str) -> str:
    return answer.split("~;~")[0]


@pytest.mark.parametrize(("which", "step"), [("structured", "r"), ("renamed", "list_column")])
def test_excels_files_read_back_to_the_formulas_excel_read(which: str, step: str) -> None:
    app = ExcelApplication.open(FIXTURES / f"{which}.xlsx", with_vba=False)
    book = app.workbook
    wrong: list[str] = []
    for case in CASES:
        sheet = book.sheets_[int(case["sheet"]) - 1]
        found = sheet.vba_get("Range", [str(case["cell"])])
        assert isinstance(found, Range)
        got = str(found.vba_get("Formula"))
        if got != _formula_of(ANSWERS[f"{step}:{case['name']}"]):
            wrong.append(f"{case['name']}: {got!r}")
    assert not wrong
