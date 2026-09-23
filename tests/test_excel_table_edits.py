"""Editing a table's rows and columns, replayed against the in-memory model.

tests/fixtures/tables/edits/ is what scripts/measure_table_edits.py saw
in live Excel: one sheet -- Table1 over A1:C4, cells around it, formulas
naming it -- edited a different way in each workbook with ListRows,
ListColumns, Resize, Delete, a formula written into a column or a value
written next to the table, then read back and saved.

The model runs the same VBA, answers the same, and saves the same table
part and cells: each cell's formula and value, a shared string read as
its text, since the order Excel numbers them in is its own.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "tables" / "edits"
RECORD: dict[str, Any] = json.loads((FIXTURES / "edits.json").read_text(encoding="utf-8"))
BOOKS: list[dict[str, Any]] = RECORD["books"]
_UID = re.compile(r'(xr3?:uid)="\{[0-9A-F-]+\}"')
_CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*?)(?:/>|>(.*?)</c>)', re.DOTALL)
_STRING = re.compile(r"<si>(.*?)</si>", re.DOTALL)

#: Workbooks the model edits differently, and why.
KNOWN: dict[str, str] = {
    "unlisted": "Unlist turns the table's style into each cell's own format, which the model does not do",
    "calculated_one/part": "a column made calculated from one cell brings a General dxf into styles.xml, "
                           "dataDxfId=\"0\" on the column, which the model does not add",
}


def _run(book: dict[str, Any], folder: Path) -> tuple[list[str], dict[str, str]]:
    app = ExcelApplication()
    app.add_workbook()
    lines = [RECORD["helpers"], "Public Function Probe() As String",
             "Dim wb As Object, ws As Object, t As Object, r As Object, c As Object, out As String",
             "Set wb = ActiveWorkbook", "Set ws = wb.Worksheets(1)", RECORD["setup"], book["editing"]]
    lines += [f'out = out & Q{number}(ws) & "~|~"' for number in range(len(RECORD["questions"]))]
    lines += ["Probe = out", "End Function"]
    for number, question in enumerate(RECORD["questions"]):
        lines += [f"Private Function Q{number}(ws As Object) As String", "On Error GoTo Bad",
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
        if formula or text:
            out[reference] = (formula.group(0) if formula else "", text)
    return out


def _loose(xml: str) -> str:
    return _UID.sub(r'\1="{...}"', xml)


@pytest.fixture(scope="module")
def model(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[list[str], dict[str, str]]]:
    folder = tmp_path_factory.mktemp("edits")
    return {book["name"]: _run(book, folder) for book in BOOKS if book["name"] not in KNOWN}


def _param(book: dict[str, Any], what: str = "") -> Any:
    key = f"{book['name']}/{what}" if what else book["name"]
    reason = KNOWN.get(book["name"]) or KNOWN.get(key)
    marks = [pytest.mark.xfail(reason=reason, strict=True)] if reason else []
    return pytest.param(book, marks=marks, id=book["name"])


@pytest.mark.parametrize("book", [_param(book) for book in BOOKS])
def test_the_model_answers_as_excel_did(book: dict[str, Any], model: dict[str, Any],
                                        tmp_path: Path) -> None:
    answers = model[book["name"]][0] if book["name"] in model else _run(book, tmp_path)[0]
    assert answers == book["answers"]


@pytest.mark.parametrize("book", [_param(book, "part") for book in BOOKS])
def test_a_save_writes_the_table_part_excel_wrote(book: dict[str, Any], model: dict[str, Any],
                                                   tmp_path: Path) -> None:
    ours = model[book["name"]][1] if book["name"] in model else _run(book, tmp_path)[1]
    wanted = {part: _loose(text) for part, text in book["parts"].items() if part.startswith("xl/tables/")}
    assert {part: _loose(text) for part, text in ours.items() if part.startswith("xl/tables/")} == wanted


@pytest.mark.parametrize("book", [_param(book, "cells") for book in BOOKS])
def test_a_save_writes_the_cells_excel_wrote(book: dict[str, Any], model: dict[str, Any], tmp_path: Path) -> None:
    ours = model[book["name"]][1] if book["name"] in model else _run(book, tmp_path)[1]
    assert _cells(ours) == _cells(book["parts"])
