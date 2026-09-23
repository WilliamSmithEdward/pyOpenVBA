"""Data validation, replayed against the in-memory model.

tests/fixtures/validation/ is what scripts/measure_validation.py saw in
live Excel: for each workbook, the VBA that gave it validation rules --
Add with each type, operator and argument, Modify, Delete, the settings
changed, cells inserted, deleted, copied and cleared round them, values
checked against them -- every property the Validation object of a few
cells answered, and the dataValidations element Excel saved.

The model runs the same VBA, answers the same for every cell, and saves
the same element; the xr:uid Excel gives each rule is a fresh random id
every time, so it is compared loosely. A workbook Excel saved is read
back to the same answers too.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "validation"
RECORD: dict[str, Any] = json.loads((FIXTURES / "validation.json").read_text(encoding="utf-8"))
BOOKS: list[dict[str, Any]] = RECORD["books"]
_UID = re.compile(r' xr:uid="\{[0-9A-F-]+\}"')

#: Workbooks the model answers or saves differently, and why.
KNOWN: dict[str, str] = {
    "other_sheet_list": "a list read from another sheet is kept in the worksheet's x14 extension, which the model "
                        "reads and does not write",
}


def _probe(book: dict[str, Any]) -> str:
    lines = [RECORD["helpers"], "Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveWorkbook.Worksheets(1)", book["making"],
             'out = "E1~:~" & ws.Range("E1").Value & "~|~"']
    lines += [f'out = out & "{cell}~:~" & Describe(ws.Range("{cell}")) & "~|~"' for cell in RECORD["read"]]
    return "\n".join([*lines, "Probe = out", "End Function"]) + "\n"


def _answers(text: str) -> dict[str, str]:
    return dict(part.split("~:~", 1) for part in text.split("~|~") if "~:~" in part)


def _wanted(book: dict[str, Any]) -> dict[str, str]:
    return {"E1": book["e1"], **book["cells"]}


def _param(book: dict[str, Any]) -> Any:
    marks = [pytest.mark.xfail(reason=KNOWN[book["name"]], strict=True)] if book["name"] in KNOWN else []
    return pytest.param(book, marks=marks, id=book["name"])


@pytest.mark.parametrize("book", [_param(book) for book in BOOKS])
def test_the_model_answers_as_excel_did(book: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_probe(book), name="Probe")
    assert _answers(str(app.run("Probe"))) == _wanted(book)
    saved = app.save(tmp_path / f"{book['name']}.xlsx")
    with zipfile.ZipFile(saved) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found = re.search(r"<dataValidations\b.*?</dataValidations>", sheet, re.DOTALL)
    assert _UID.sub("", found.group(0) if found else "") == _UID.sub("", book["saved"])


@pytest.mark.parametrize("book", [pytest.param(book, id=book["name"]) for book in BOOKS])
def test_a_workbook_excel_saved_reads_back_the_same(book: dict[str, Any]) -> None:
    app = ExcelApplication.open(FIXTURES / f"{book['name']}.xlsx", with_vba=False)
    lines = [RECORD["helpers"], "Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveWorkbook.Worksheets(1)"]
    lines += [f'out = out & "{cell}~:~" & Describe(ws.Range("{cell}")) & "~|~"' for cell in RECORD["read"]]
    app.add_module("\n".join([*lines, "Probe = out", "End Function"]) + "\n", name="Probe")
    assert _answers(str(app.run("Probe"))) == book["cells"]
