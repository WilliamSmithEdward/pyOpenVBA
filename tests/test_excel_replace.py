"""Range.Replace, replayed against the in-memory model.

tests/fixtures/replace.json is what scripts/measure_replace.py saw in live
Excel: layouts that fill a few cells, replace something and dump A1:D4 --
each cell's formula, value type, number format and prefix character --
with the used range and what Replace returned, or what a Find after it
found; and 20 numbers in 26 formats, each read back as the text Replace
edits by replacing every digit in turn with q. Each layout runs on a
fresh workbook in the model, with the year a date typed without one
falls in held at the year the fixture was measured.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication, _typing

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "replace.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
EDIT_TEXTS: dict[str, Any] = RECORD["edit_texts"]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant, f As Object",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(_typing, "_this_year", lambda: RECORD["year"])
        return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_replace_leaves_what_excel_leaves(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]


@pytest.fixture(scope="module")
def edited() -> list[list[str]]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(EDIT_TEXTS["code"], name="Texts")
    rows = str(app.run("Texts")).split("|")[: len(EDIT_TEXTS["values"])]
    return [row.split("^")[: len(EDIT_TEXTS["formats"])] for row in rows]


@pytest.mark.parametrize("index", range(len(EDIT_TEXTS["values"])), ids=EDIT_TEXTS["values"])
def test_replace_edits_a_number_as_the_formula_bar_shows_it(edited: list[list[str]], index: int) -> None:
    assert dict(zip(EDIT_TEXTS["formats"], edited[index], strict=True)) == dict(
        zip(EDIT_TEXTS["formats"], EDIT_TEXTS["texts"][index], strict=True))
