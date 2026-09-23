"""Excel tables, replayed against the in-memory model.

tests/fixtures/tables/ is what scripts/measure_tables.py saw: a workbook
for each way of making a table, as Excel saved it, and in tables.json
what the object model answered about every table after the file was
opened again, the parts Excel wrote for them, and the answers to a list
of questions about the "two" workbook.

The model reads each file's tables as Excel reads them, makes each
workbook's tables from the same VBA and answers the same, and saves them
in the parts Excel writes. The xr:uid Excel gives a table and each of its
columns is a fresh random id every time, so it is compared loosely.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "tables"
RECORD: dict[str, Any] = json.loads((FIXTURES / "tables.json").read_text(encoding="utf-8"))
BOOKS: dict[str, dict[str, Any]] = {book["name"]: book for book in RECORD["books"]}
QUESTIONS: list[dict[str, str]] = RECORD["questions"]
_UID = re.compile(r'(xr3?:uid)="\{[0-9A-F-]+\}"')

#: Workbooks and questions the model answers differently, and why.
KNOWN: dict[str, str] = {
    "totals": "showing a table's totals row is not implemented",
}


def _loose(xml: str) -> str:
    return _UID.sub(r'\1="{...}"', xml)


def _run(app: ExcelApplication, body: str, answer: str) -> str:
    source = "\n".join([RECORD["reader"], "Public Function Probe() As String",
                        "Dim ws As Object, t As Object", "Set ws = ActiveWorkbook.Worksheets(1)", body,
                        f"Probe = {answer}", "End Function"])
    app.add_module(source + "\n", name="Probe")
    return str(app.run("Probe"))


def _param(name: str) -> Any:
    marks = [pytest.mark.xfail(reason=KNOWN[name], strict=True)] if name in KNOWN else []
    return pytest.param(name, marks=marks, id=name)


@pytest.mark.parametrize("name", list(BOOKS))
def test_the_model_reads_each_table_excel_saved(name: str) -> None:
    app = ExcelApplication.open(FIXTURES / f"{name}.xlsx", with_vba=False)
    assert _run(app, "", "Tables(ws)") == BOOKS[name]["read"]


@pytest.mark.parametrize("name", [_param(name) for name in BOOKS])
def test_a_table_made_by_vba_answers_as_excels_does(name: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    assert _run(app, BOOKS[name]["making"], "Tables(ws)") == BOOKS[name]["read"]


@pytest.mark.parametrize("name", [_param(name) for name in BOOKS])
def test_a_save_writes_the_parts_excel_writes(name: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    _run(app, BOOKS[name]["making"], '""')
    saved = app.save(tmp_path / f"{name}.xlsx")
    with zipfile.ZipFile(saved) as package:
        names = package.namelist()
        got = {part: _loose(package.read(part).decode("utf-8")) for part in names if part.startswith("xl/tables/")}
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        types = package.read("[Content_Types].xml").decode("utf-8")
        rels = package.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    wanted = BOOKS[name]["parts"]
    assert got == {part: _loose(text) for part, text in wanted.items() if part.startswith("xl/tables/")}
    found = re.search(r"<tableParts\b.*?</tableParts>", sheet, re.DOTALL)
    assert found is not None and found.group(0) == wanted["tableParts"]
    assert " ".join(re.findall(r"<Override\b[^>]*tables/[^>]*/>", types)) == wanted["content_types"]
    assert sorted(re.findall(r"<Relationship\b[^>]*/>", rels)) == sorted(
        re.findall(r"<Relationship\b[^>]*/>", wanted["sheet_rels"]))


@pytest.mark.parametrize("name", [_param(one["name"]) for one in QUESTIONS])
def test_a_question_about_tables_answers_as_excels_does(name: str) -> None:
    case = next(one for one in QUESTIONS if one["name"] == name)
    app = ExcelApplication.open(FIXTURES / "two.xlsx", with_vba=False)
    body = f'On Error Resume Next\nDim answer As String\nanswer = {case["question"]}\n' \
           'If Err.Number <> 0 Then answer = "!" & Err.Number'
    assert _run(app, body, "answer") == case["answer"]
