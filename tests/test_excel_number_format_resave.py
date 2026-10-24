"""A workbook's own number formats as Excel numbers them, replayed against the in-memory model.

tests/fixtures/number_format_resave/ is what
scripts/measure_number_format_resave.py saw in live Excel: a workbook
whose custom formats came after a conditional format's, four on cells, a
pivot table's data field and a rule kept in the worksheet's Excel 2010
extension, and copies of it listing the cells' formats out of order, with
gaps, with one nothing uses, one only an unused xf uses, one code twice, a
built-in id over another code, an id below 164, a spelling Excel does not
write, an xf naming a format the file lacks, and the pivot's format at an
id of its own. Excel opened each, read every format, and saved it with
nothing changed. The model opens each, reads it the same, and saves it
with nothing changed as Excel did.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "number_format_resave"
RECORD: dict[str, Any] = json.loads((FIXTURES / "number_format_resave.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
#: What a file holds of its formats, as number_format_resave.json keeps it.
SLOTS = ("numFmts", "xfs", "cells", "dxfs", "extension", "pivot")
#: Each workbook's reader: every cell's format and what it shows, the pivot's data field's format and the rules'.
READER = RECORD["read"].split("Public Function Read")[0] + "\n".join([
    "Public Function Read() As String", "Dim wb As Object, ws As Object, out As String, r As Long, f As Object",
    "Set wb = ActiveWorkbook", 'Set ws = wb.Worksheets("cells")', "For r = 1 To 4",
    '    out = out & ws.Cells(r, 1).NumberFormat & "~" & ws.Cells(r, 1).Text & "~"', "Next",
    "On Error Resume Next", 'Set f = wb.Worksheets("pivot").PivotTables(1).DataFields(1)',
    'out = out & Shown(f) & "~"', "Set f = Nothing", 'Set f = ws.Range("B2").FormatConditions(1)',
    'out = out & Shown(f) & "~"', "Set f = Nothing", 'Set f = ws.Range("C2").FormatConditions(1)',
    'If f Is Nothing Then out = out & "E" & Err.Number Else out = out & Shown(f)', "Read = out",
    "End Function"]) + "\n"


def _part(package: zipfile.ZipFile, sheet: str) -> str:
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = {found.group(1): found.group(2) for found in re.finditer(
        r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', package.read("xl/_rels/workbook.xml.rels").decode())}
    rid = re.search(rf'<sheet\b[^>]*?name="{sheet}"[^>]*?r:id="([^"]+)"', workbook)
    assert rid is not None
    target = targets[rid.group(1)].lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _within(text: str, section: str) -> str:
    found = re.search(rf"<{section}\b.*?</{section}>", text, re.DOTALL)
    return found.group() if found else ""


def numbers(path: Path) -> dict[str, Any]:
    """What a file holds of its formats, as the measurement read it: the stylesheet's numFmts, each cell xf's
    numFmtId, the one each of the sheet "cells" A1:A4 points at, the dxfs' numFmts, the numFmt of the rule in that
    sheet's extension, and every numFmtId of the pivot's parts."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read(_part(package, "cells")).decode("utf-8")
        pivots = {name: re.findall(r'<(\w+)\b[^>]*?\bnumFmtId="(\d+)"', package.read(name).decode("utf-8"))
                  for name in package.namelist() if name.startswith("xl/pivot") and name.endswith(".xml")}
    xfs = [re.search(r'numFmtId="(\d+)"', xf) for xf in re.findall(r"<xf\b[^>]*?>", _within(styles, "cellXfs"))]
    named = [found.group(1) if found else "" for found in xfs]
    return {"numFmts": re.findall(r"<numFmt [^>]*/>", _within(styles, "numFmts")),
            "xfs": named,
            "cells": {found.group(1): named[int(found.group(2))]
                      for found in re.finditer(r'<c r="(A[1-4])" s="(\d+)"', sheet)},
            "dxfs": re.findall(r"<numFmt [^>]*/>", _within(styles, "dxfs")),
            "extension": re.findall(r"<numFmt [^>]*/>", _within(sheet, "x14:dxf")),
            "pivot": {name: [f"{tag}:{identifier}" for tag, identifier in found] for name, found in pivots.items()}}


@pytest.fixture(scope="module")
def modelled(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[list[str], dict[str, Any]]]:
    """Each workbook opened in the model: what it reads, and what its save with nothing changed holds."""
    folder = tmp_path_factory.mktemp("number_format_resave")
    out: dict[str, tuple[list[str], dict[str, Any]]] = {}
    for case in CASES:
        app = ExcelApplication.open(FIXTURES / f"{case['name']}.xlsx", with_vba=False)
        app.add_module(READER, name="Reader")
        answer = str(app.run("Read")).split("~")
        out[case["name"]] = answer, numbers(app.save(folder / f"{case['name']}.xlsx"))
    return out


def test_the_record_holds_what_excels_files_hold() -> None:
    assert all(numbers(FIXTURES / f"{case['name']}.xlsx") == case["file"]
               and numbers(FIXTURES / f"{case['name']}_resaved.xlsx") == case["resaved"] for case in CASES)


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_a_file_reads_its_formats_as_excel_does(modelled: dict[str, tuple[list[str], Any]],
                                               case: dict[str, Any]) -> None:
    """A built-in id over another code reads that code; an id the file lists no format for reads General."""
    assert modelled[case["name"]][0][:10] == case["answer"].split("~")[:10]


@pytest.mark.xfail(strict=True, reason="a rule kept only in the worksheet's Excel 2010 extension, here one naming "
                                       "another sheet, is not one of the range's FormatConditions in the model")
def test_a_rule_kept_in_the_extension_reads_its_format_as_excel_does(
        modelled: dict[str, tuple[list[str], Any]]) -> None:
    assert [modelled[case["name"]][0][10] for case in CASES] == [case["answer"].split("~")[10] for case in CASES]


def test_a_second_save_keeps_what_the_first_wrote(tmp_path: Path) -> None:
    """Saved twice with nothing changed, a file holds its formats as the first save numbered them."""
    case = next(one for one in CASES if one["name"] == "gaps")
    app = ExcelApplication.open(FIXTURES / "gaps.xlsx", with_vba=False)
    first = numbers(app.save(tmp_path / "first.xlsx"))
    second = numbers(app.save(tmp_path / "second.xlsx"))
    assert first == second and first["numFmts"] == case["resaved"]["numFmts"]
    app.add_module(READER, name="Reader")
    assert str(app.run("Read")).split("~")[:10] == case["answer"].split("~")[:10]


@pytest.mark.parametrize(("case", "slot"), [
    pytest.param(case, slot, id=f"{case['name']}-{slot}") for case in CASES for slot in SLOTS])
def test_a_file_saved_unchanged_numbers_its_formats_as_excel_does(
        modelled: dict[str, tuple[Any, dict[str, Any]]], case: dict[str, Any], slot: str) -> None:
    """Each format under the id Excel gave it on opening: from 164 in the order the file lists them, one code one
    id, whatever ids the file gave them, the pivot's too, and spelled as Excel spells it; one nothing names is left
    out."""
    assert modelled[case["name"]][1][slot] == case["resaved"][slot]
