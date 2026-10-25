"""A worksheet's tab colour, Worksheet.Tab, replayed against the in-memory model.

tests/fixtures/sheet_tabs/ holds what scripts/measure_sheet_tabs.py saw in
live Excel: a sheet for each case, its tab's Color, ColorIndex, ThemeColor,
TintAndShade or Pattern set, an error's number where one was raised, every
one read back with its type, a copy of one sheet, each sheet's sheetPr as
Excel saved it, and every tab read again once the workbook was opened
again.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "sheet_tabs"
RECORD: dict[str, Any] = json.loads((FIXTURES / "sheet_tabs.json").read_text(encoding="utf-8"))


def sheet_properties(path: Path) -> dict[str, str]:
    """Each sheet's sheetPr as the file keeps it, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, str] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            found = re.search(r"<sheetPr\b[^>]*/>|<sheetPr\b.*?</sheetPr>", text, re.DOTALL)
            out[name] = found.group() if found else ""
        return out


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every case made in the model with the macro Excel ran, its dumps before the save and once the model opened
    its own file again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("sheet_tabs")
    app = ExcelApplication()
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    app.add_module(RECORD["build"].replace('folder & "\\" & name', f'folder & "{os.sep}" & name'), name="Builder")
    session, reopened = str(app.run("Build", str(folder))).split("@@", 1)

    def dumps(text: str) -> dict[str, str]:
        return dict(one.split("^", 1) for one in text.split("~") if "^" in one)

    return dumps(session), dumps(reopened), folder


@pytest.mark.parametrize("name", list(RECORD["session"]))
def test_a_tab_reads_back_what_excel_reads(made: tuple[dict[str, str], dict[str, str], Path], name: str) -> None:
    """Color, ColorIndex, ThemeColor, TintAndShade and Pattern with their types, and the errors the settings
    raise."""
    assert made[0][name] == RECORD["session"][name]


@pytest.mark.parametrize("name", list(RECORD["reopened"]))
def test_a_tab_reads_back_from_the_file_the_model_saved(made: tuple[dict[str, str], dict[str, str], Path],
                                                        name: str) -> None:
    assert made[1][name] == RECORD["reopened"][name]


def test_the_file_keeps_each_tab_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Each sheet's sheetPr: a tabColor of rgb, theme or indexed with its tint, and none for a tab with no
    colour."""
    assert sheet_properties(made[2] / "sheet_tabs.xlsx") == RECORD["sheet_properties"]


@pytest.mark.parametrize("rgb,index", RECORD["nearest"], ids=lambda value: str(value))
def test_a_tab_finds_the_palette_slot_excels_finds(rgb: list[int], index: int) -> None:
    """Tab.ColorIndex of a colour off the palette, which a tab finds otherwise than a cell: a near-grey's nearest
    grey, or the nearest slot in Windows' HLS."""
    from pyopenvba.apps.excel import _styles as S

    assert S.Colors().tab_index(S.Color.rgb(*rgb)) == index


def test_the_model_reads_excels_tabs() -> None:
    """Excel's own file, opened in the model, reads back every tab as Excel read it."""
    app = ExcelApplication.open(FIXTURES / "sheet_tabs.xlsx", with_vba=False)
    app.add_module(RECORD["build"].split("Public Function Build")[0] + (
        "Public Function Read() As String\nDim ws As Object, out As String\n"
        'For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & "^" & Dump(ws) & "~"\nNext\n'
        "Read = out\nEnd Function\n"), name="Reader")
    read = dict(one.split("^", 1) for one in str(app.run("Read")).split("~") if "^" in one)
    assert read == RECORD["reopened"]
