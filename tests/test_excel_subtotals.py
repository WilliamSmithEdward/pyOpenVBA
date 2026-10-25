"""Range.Subtotal and Range.RemoveSubtotal, replayed against the in-memory model.

tests/fixtures/subtotals/ holds what scripts/measure_subtotals.py saw in
live Excel: a sheet for each case, a list typed into it, Subtotal or
RemoveSubtotal called on it or both, an error's number where a call raised
one, and every cell, row and outline setting read back; each sheet's
outline and page breaks as Excel saved them; every sheet read again once
the workbook was opened again; and the events the calls raise.

Each case runs here with Excel's own function for it, in a workbook of its
own, saved and opened again: one workbook of every case would have each
edit work out every SUBTOTAL of them all again, as a volatile formula is.
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
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "subtotals"
RECORD: dict[str, Any] = json.loads((FIXTURES / "subtotals.json").read_text(encoding="utf-8"))
NAMES = [case["name"] for case in RECORD["cases"]]

#: The cases the model reports unsupported, and why.
UNSUPPORTED = {
    **dict.fromkeys(("breaks_below", "breaks_above", "breaks_nested", "breaks_nested_inner", "breaks_remove",
                     "breaks_replace", "breaks_inner_after_reset"),
                    "Subtotal's page breaks between groups, which Excel keeps on the sheet and a later call reuses"),
    **dict.fromkeys(("breaks_manual_kept", "breaks_default_after_manual"),
                    "manual page breaks, HPageBreaks among them"),
    "filtered": "a filtered list, which Excel subtotals over its visible rows and leaves unoutlined",
}
SUPPORTED = [name for name in NAMES if name not in UNSUPPORTED]

#: Excel's functions for the cases, before its Build, with one that makes a case alone, saves it and reads it back.
ONE = RECORD["build"].split("Public Function Build", 1)[0] + '''
Public Function One(index As Long, folder As String) As String
    Dim wb As Object, ws As Object, scratch As Object, out As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set scratch = wb.Worksheets(1)
    scratch.Name = "scratch"
    scratch.Range("A1").NumberFormat = "General"
    Select Case index
{cases}
    End Select
    wb.SaveAs folder & "{separator}" & "case.xlsx", 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "{separator}" & "case.xlsx")
    out = out & "@@"
    For Each ws In wb.Worksheets
        out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    One = out
End Function
'''.replace("{cases}", "".join(f"        Case {index}: out = Make{index}(wb, scratch)\n"
                               for index in range(len(NAMES)))).replace("{separator}", os.sep)


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("~") if "^" in one)


def parts(text: str) -> dict[str, object]:
    """A sheet's outline settings, its rows' outline attributes and its page breaks."""
    found = re.search(r"<sheetPr\b[^>]*/>|<sheetPr\b.*?</sheetPr>", text, re.DOTALL)
    outline = re.search(r"<outlinePr\b[^>]*/>", found.group()) if found else None
    format_found = re.search(r"<sheetFormatPr\b[^>]*/>", text)
    levels = {}
    for row in re.finditer(r"<row\b([^>]*?)/?>", text):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', row.group(1)))
        kept = {key: attributes[key] for key in ("outlineLevel", "hidden", "collapsed") if key in attributes}
        if kept:
            levels[attributes["r"]] = kept
    breaks = re.search(r"<rowBreaks\b.*?</rowBreaks>|<rowBreaks\b[^>]*/>", text, re.DOTALL)
    outline_level = re.search(r'\boutlineLevelRow="(\d+)"', format_found.group()) if format_found else None
    return {"outlinePr": outline.group() if outline else "",
            "outlineLevelRow": outline_level.group(1) if outline_level else "",
            "rows": levels, "rowBreaks": breaks.group() if breaks else ""}


def sheet_parts(path: Path) -> dict[str, dict[str, object]]:
    """Each sheet's outline and page breaks as the file keeps them, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, dict[str, object]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            out[name] = parts(package.read(part).decode("utf-8"))
        return out


def cell_formats(path: Path, sheet: str) -> dict[str, str]:
    """Each cell of a sheet the file gives a format, and that xf with its font, fill and number format written out
    in place of their numbers, which differ between one workbook of every case and one of a case alone."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels))
        found = re.search(rf'<sheet\b[^>]*\bname="{re.escape(sheet)}"[^>]*\br:id="([^"]*)"', workbook)
        assert found is not None
        text = package.read("xl/" + targets[found.group(1)].lstrip("/").removeprefix("xl/")).decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")

    def items(section: str, tag: str) -> list[str]:
        block = re.search(rf"<{section}\b.*?</{section}>", styles, re.DOTALL)
        assert block is not None
        return re.findall(rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>", block.group(), re.DOTALL)

    xfs, fonts, fills = items("cellXfs", "xf"), items("fonts", "font"), items("fills", "fill")
    codes = dict(re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles))
    out: dict[str, str] = {}
    for address, index in re.findall(r'<c r="([A-Z]+\d+)"[^>]*?\bs="(\d+)"', text):
        xf = xfs[int(index)]
        number = re.search(r'numFmtId="(\d+)"', xf)
        font = re.search(r'fontId="(\d+)"', xf)
        fill = re.search(r'fillId="(\d+)"', xf)
        out[address] = " ".join((re.sub(r'\b(?:numFmtId|fontId|fillId)="\d+"', "", xf),
                                 fonts[int(font.group(1))] if font else "", fills[int(fill.group(1))] if fill else "",
                                 codes.get(number.group(1), number.group(1)) if number else ""))
    return out


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_module(ONE, name="Cases")
    return made


@pytest.fixture(scope="module")
def made(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[Any, ...]]:
    """Each supported case made in the model with Excel's function for it: its dump before the save, once the
    model opened its own file again, and its sheet's outline and cells' formats as that file keeps them."""
    out: dict[str, tuple[Any, ...]] = {}
    for name in SUPPORTED:
        folder = tmp_path_factory.mktemp(name)
        try:
            session, reopened = str(app.run("One", NAMES.index(name), str(folder))).split("@@", 1)
        except Exception as exc:  # noqa: BLE001 -- one case's failure is that case's, not every case's.
            out[name] = (repr(exc),) * 4
            continue
        out[name] = (dumps(session)[name], dumps(reopened)[name], sheet_parts(folder / "case.xlsx")[name],
                     cell_formats(folder / "case.xlsx", name))
    return out


@pytest.mark.parametrize("name", SUPPORTED)
def test_a_list_subtotals_as_excels_does(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """Every cell with its formula, value, type, bold, italic, number format and fill, each row's outline level and
    whether it shows, the summary row and the errors the calls raise."""
    assert made[name][0] == RECORD["session"][name]


@pytest.mark.parametrize("name", SUPPORTED)
def test_a_subtotalled_list_reads_back_from_the_file_the_model_saved(made: dict[str, tuple[Any, ...]],
                                                                     name: str) -> None:
    assert made[name][1] == RECORD["reopened"][name]


@pytest.mark.parametrize("name", SUPPORTED)
def test_the_file_keeps_the_outline_as_excels_does(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """The sheet's outlinePr, its sheetFormatPr's outlineLevelRow and each row's outline attributes."""
    assert made[name][2] == RECORD["parts"][name]


@pytest.mark.parametrize("name", SUPPORTED)
def test_the_file_formats_each_cell_as_excels_does(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """Each cell's xf, its applied parts, font, fill and number format, labels and formulas among them."""
    assert made[name][3] == cell_formats(FIXTURES / "subtotals.xlsx", name)


@pytest.fixture(scope="module")
def excels_file_read() -> dict[str, str]:
    """Every sheet of the workbook Excel saved, as the model reads it."""
    reader = ExcelApplication()
    reader.add_module(RECORD["build"].split("Public Function Build", 1)[0] + '''
Public Function ReadBack(path As String) As String
    Dim wb As Object, ws As Object, out As String
    Set wb = Workbooks.Open(path)
    For Each ws In wb.Worksheets
        out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    ReadBack = out
End Function
''', name="Reader")
    return dumps(str(reader.run("ReadBack", str(FIXTURES / "subtotals.xlsx"))))


@pytest.mark.parametrize("name", list(RECORD["reopened"]))
def test_excels_file_reads_as_excel_reads_it(excels_file_read: dict[str, str], name: str) -> None:
    """Subtotal rows, bold labels, SUBTOTAL formulas and outline levels, the page breaks' sheets too."""
    assert excels_file_read[name] == RECORD["reopened"][name]


@pytest.mark.parametrize("dependent", [True, False], ids=["dependent", "plain"])
def test_subtotal_raises_calculate_and_no_change(dependent: bool) -> None:
    """One Calculate for each Subtotal; RemoveSubtotal raises one only where a formula reads the list."""
    watcher = ExcelApplication()
    watcher.add_module("Public Log As String\n", name="EventLog")
    watcher.add_module(RECORD["sink"], name="SubtotalSink", kind="class")
    watcher.add_module(RECORD["watch"], name="Watcher")
    assert watcher.run("Watch", dependent) == RECORD["events"]["dependent" if dependent else "plain"]


@pytest.mark.parametrize("name", sorted(UNSUPPORTED))
def test_what_the_model_does_not_do_reports_itself(app: ExcelApplication, name: str, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        app.run("One", NAMES.index(name), str(tmp_path))
