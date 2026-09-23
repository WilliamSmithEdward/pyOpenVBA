"""A table's totals row, replayed against the in-memory model.

tests/fixtures/tables/totals_row/ is what scripts/measure_totals_row.py
saw in live Excel: for each workbook, the VBA that made it -- a table,
its totals row shown, hidden, shown again, its functions set, its cells
written over -- what questions about it answered, and the table part and
cells Excel saved.

The model runs the same VBA, answers the same, and saves the same table
part and the same cells: each cell's formula and value, with shared
strings read through, since the order Excel numbers them in is its own.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "tables" / "totals_row"
RECORD: dict[str, Any] = json.loads((FIXTURES / "totals_row.json").read_text(encoding="utf-8"))
BOOKS: list[dict[str, Any]] = RECORD["books"]
_UID = re.compile(r'(xr3?:uid)="\{[0-9A-F-]+\}"')
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*?)(?:/>|>(.*?)</c>)', re.DOTALL)
_STRING = re.compile(r"<si>(.*?)</si>", re.DOTALL)


def _run(book: dict[str, Any], folder: Path) -> tuple[list[str], dict[str, str]]:
    """The answers the model gives, and the parts of the workbook it saves."""
    app = ExcelApplication()
    app.add_workbook()
    lines = [RECORD["helpers"], "Public Function Probe() As String",
             "Dim wb As Object, ws As Object, t As Object, out As String", "Set wb = ActiveWorkbook",
             "Set ws = wb.Worksheets(1)", book["making"]]
    lines += [f'out = out & Q{number}(ws, t) & "~|~"' for number in range(len(book["questions"]))]
    lines += ["Probe = out", "End Function"]
    for number, question in enumerate(book["questions"]):
        lines += [f"Private Function Q{number}(ws As Object, t As Object) As String", "On Error GoTo Bad",
                  f"Q{number} = {question}", "Exit Function", "Bad:", f'Q{number} = "!" & Err.Number',
                  "End Function"]
    app.add_module("\n".join(lines) + "\n", name="Probe")
    answers = str(app.run("Probe")).split("~|~")[:-1]
    saved = app.save(folder / f"{book['name']}.xlsx")
    with zipfile.ZipFile(saved) as package:
        parts = {name: package.read(name).decode("utf-8") for name in package.namelist()
                 if name.startswith("xl/tables/")}
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        found = re.search(r"<sheetData>.*?</sheetData>|<sheetData/>", sheet, re.DOTALL)
        parts["sheetData"] = found.group(0) if found else ""
        strings = "xl/sharedStrings.xml"
        parts["sharedStrings"] = package.read(strings).decode("utf-8") if strings in package.namelist() else ""
    return answers, parts


def _cells(parts: dict[str, str]) -> dict[str, tuple[str, str]]:
    """Each cell's formula and value, text read the same whether it is a shared string or written in the cell."""
    strings = [re.sub(r"<[^>]+>", "", item) for item in _STRING.findall(parts["sharedStrings"])]
    out: dict[str, tuple[str, str]] = {}
    for reference, attributes, body in _CELL.findall(parts["sheetData"]):
        formula = re.search(r"<f\b[^>]*?(?:/>|>.*?</f>)", body or "")
        value = re.search(r"<v>(.*?)</v>", body or "")
        inline = re.search(r"<is>(.*?)</is>", body or "", re.DOTALL)
        text = value.group(1) if value else ""
        if ' t="s"' in attributes and text:
            text = "string:" + strings[int(text)]
        elif inline is not None:
            text = "string:" + re.sub(r"<[^>]+>", "", inline.group(1))
        out[reference] = (formula.group(0) if formula else "", text)
    return out


def _loose(xml: str) -> str:
    return _UID.sub(r'\1="{...}"', xml)


@pytest.fixture(scope="module")
def model(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[list[str], dict[str, str]]]:
    folder = tmp_path_factory.mktemp("totals")
    return {book["name"]: _run(book, folder) for book in BOOKS}


@pytest.mark.parametrize("book", BOOKS, ids=[book["name"] for book in BOOKS])
def test_the_model_answers_as_excel_did(book: dict[str, Any], model: dict[str, Any]) -> None:
    assert model[book["name"]][0] == book["answers"]


@pytest.mark.parametrize("book", BOOKS, ids=[book["name"] for book in BOOKS])
def test_a_save_writes_the_table_part_excel_wrote(book: dict[str, Any], model: dict[str, Any]) -> None:
    ours = model[book["name"]][1]
    wanted = {part: _loose(text) for part, text in book["parts"].items() if part.startswith("xl/tables/")}
    assert {part: _loose(text) for part, text in ours.items() if part.startswith("xl/tables/")} == wanted


#: Cells the model saves differently, and why.
KNOWN: dict[str, str] = {
    "functions/I4": "STDEV's answer is saved as 0.7071067811865476, the shortest spelling that reads back the same "
                    "number, where Excel writes 0.70710678118654757: a cell's <v> follows a spelling rule the model "
                    "does not have yet",
}
_SAVED = [(book, cell) for book in BOOKS for cell in _cells(book["parts"])]


@pytest.mark.parametrize(("book", "cell"), [
    pytest.param(book, cell, id=f"{book['name']}/{cell}",
                 marks=[pytest.mark.xfail(reason=KNOWN[f"{book['name']}/{cell}"], strict=True)]
                 if f"{book['name']}/{cell}" in KNOWN else [])
    for book, cell in _SAVED])
def test_a_save_writes_the_cell_excel_wrote(book: dict[str, Any], cell: str, model: dict[str, Any]) -> None:
    assert _cells(model[book["name"]][1]).get(cell) == _cells(book["parts"])[cell]


@pytest.mark.parametrize("book", BOOKS, ids=[book["name"] for book in BOOKS])
def test_a_save_writes_no_cell_excel_did_not(book: dict[str, Any], model: dict[str, Any]) -> None:
    assert set(_cells(model[book["name"]][1])) - set(_cells(book["parts"])) == set()
