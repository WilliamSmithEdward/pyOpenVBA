"""A stylesheet as Excel saves it when nothing changed it, replayed against the in-memory model.

tests/fixtures/unchanged_save/ is what scripts/measure_unchanged_save.py
saw in live Excel: a workbook with cells in a format each, a row's and a
column's format, a cell in the built-in style Good, a user's style that
nothing uses and a conditional format's dxf, and copies of it holding an
xf nothing uses, in the middle of the table and at its end; a font, a
fill and a border nothing uses; an unused xf with parts only it names; an
xf like an earlier one, with the same apply flags and with others; a font
twice; a built-in style nothing uses; an xf only a row uses; and a dxf no
rule names. Excel opened each, read its cells' formats and saved it with
nothing changed, writing what is in use and nothing else. The model opens
each, reads it the same, and saves it with nothing changed as Excel did.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "unchanged_save"
RECORD: dict[str, Any] = json.loads((FIXTURES / "unchanged_save.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
SECTIONS = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs")
_SECTION = {"numFmts": "numFmt", "fonts": "font", "fills": "fill", "borders": "border", "cellStyleXfs": "xf",
            "cellXfs": "xf", "cellStyles": "cellStyle", "dxfs": "dxf"}
#: Where the model's save keeps a slot otherwise than Excel's.
GAPS: dict[tuple[str, str], str] = {
    ("unused_parts", "fonts"): "Excel leaves out a font nothing uses; the model keeps every font no xf names, as "
                               "Excel keeps the fonts a form control's label and a note's text are in",
    ("unused_dxf", "dxfs"): "Excel leaves out a dxf no rule names; the model keeps the file's dxfs",
}
#: Each workbook's reader: the formats of its cells, a row's and a column's, and its styles.
READER = RECORD["read"].split("Public Function Read")[0] + "\n".join([
    "Public Function Read() As String", "Dim wb As Object, ws As Object, out As String, c As Object, s As Object",
    "Set wb = ActiveWorkbook", 'Set ws = wb.Worksheets("cells")',
    'For Each c In ws.Range("A1:A6,B1,A8,D1,A10")', '    out = out & Look(c) & ";"', "Next",
    "For Each s In wb.Styles", '    If Not s.BuiltIn Then out = out & "style:" & s.Name & ";"', "Next",
    'Read = out & "styles:" & wb.Styles.Count', "End Function"]) + "\n"


def _entries(styles: str, section: str) -> list[str]:
    found = re.search(rf"<{section}\b[^>]*?(?:/>|>(.*?)</{section}>)", styles, re.DOTALL)
    if found is None or found.group(1) is None:
        return []
    tag = _SECTION[section]
    listed = re.findall(rf"<{tag}\b[^>]*?/>|<{tag}\b[^>]*?>.*?</{tag}>", found.group(1), re.DOTALL)
    return [re.sub(r' xr:uid="[^"]*"', "", one) for one in listed]


def holdings(path: Path) -> dict[str, Any]:
    """What a file holds of its formats, as the measurement read it: each stylesheet section's entries, and the xf
    each of the sheet "cells"' cells, rows and columns points at."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    return {**{section: _entries(styles, section) for section in SECTIONS},
            "cells": dict(re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)),
            "rows": dict(re.findall(r'<row r="(\d+)"[^>]*? s="(\d+)"', sheet)),
            "cols": [f"{low}:{high}:{style}" for low, high, style in
                     re.findall(r'<col min="(\d+)" max="(\d+)"[^>]*? style="(\d+)"', sheet)]}


@pytest.fixture(scope="module")
def modelled(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[str, dict[str, Any]]]:
    """Each workbook opened in the model: what it reads, and what its save with nothing changed holds."""
    folder = tmp_path_factory.mktemp("unchanged_save")
    out: dict[str, tuple[str, dict[str, Any]]] = {}
    for case in CASES:
        app = ExcelApplication.open(FIXTURES / f"{case['name']}.xlsx", with_vba=False)
        app.add_module(READER, name="Reader")
        out[case["name"]] = str(app.run("Read")), holdings(app.save(folder / f"{case['name']}.xlsx"))
    return out


def test_the_record_holds_what_excels_files_hold() -> None:
    assert all(holdings(FIXTURES / f"{case['name']}.xlsx") == case["file"]
               and holdings(FIXTURES / f"{case['name']}_resaved.xlsx") == case["resaved"] for case in CASES)


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_a_file_reads_its_formats_as_excel_does(modelled: dict[str, tuple[str, Any]], case: dict[str, Any]) -> None:
    """An xf like an earlier one, an unused one beside it, or a font twice change nothing a cell shows."""
    assert modelled[case["name"]][0] == case["answer"]


@pytest.mark.parametrize(("case", "slot"), [
    pytest.param(case, slot, id=f"{case['name']}-{slot}",
                 marks=[pytest.mark.xfail(strict=True, reason=GAPS[case["name"], slot])]
                 if (case["name"], slot) in GAPS else [])
    for case in CASES for slot in (*SECTIONS, "cells", "rows", "cols")])
def test_a_file_saved_unchanged_holds_what_excel_saves(modelled: dict[str, tuple[Any, dict[str, Any]]],
                                                      case: dict[str, Any], slot: str) -> None:
    """What is in use and nothing else: an xf nothing uses left out, its parts with it, one like an earlier one
    joined to it, a built-in style nothing uses left out, and every cell, row and column pointing where its xf
    went; a font twice stays two, and an xf only a row uses stays."""
    assert modelled[case["name"]][1][slot] == case["resaved"][slot]


def _cell_xfs(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return _entries(package.read("xl/styles.xml").decode("utf-8"), "cellXfs")


def test_a_pivot_tables_button_keeps_its_format_apart(tmp_path: Path) -> None:
    """An xf alike in every part but a pivot table field's button stays apart from the one it is like, and keeps
    its button, as Excel keeps it (tests/fixtures/number_format_resave/base.xlsx)."""
    source = Path(__file__).parent / "fixtures" / "number_format_resave" / "base.xlsx"
    saved = ExcelApplication.open(source, with_vba=False).save(tmp_path / "pivot.xlsx")
    assert [one for one in _cell_xfs(saved) if "pivotButton" in one] == [
        one for one in _cell_xfs(source) if "pivotButton" in one] == [
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" pivotButton="1"/>']
