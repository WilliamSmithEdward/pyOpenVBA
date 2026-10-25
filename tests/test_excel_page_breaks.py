"""Manual page breaks, replayed against the in-memory model.

tests/fixtures/page_breaks/ holds what scripts/measure_page_breaks.py saw
in live Excel: a sheet for each case, breaks added, set, deleted or moved
on it, rows or columns edited and Subtotal called, an error's number
where a step raised one, and each collection's breaks and the PageBreak
of rows 1 to 15 and columns A to F read back; each sheet's rowBreaks and
colBreaks as Excel saved them; and every sheet read again once the
workbook was opened again.

Each case runs here with Excel's own function for it, in a workbook of
its own, saved and opened again.
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

FIXTURES = Path(__file__).parent / "fixtures" / "page_breaks"
RECORD: dict[str, Any] = json.loads((FIXTURES / "page_breaks.json").read_text(encoding="utf-8"))
NAMES = [case["name"] for case in RECORD["cases"]]

#: The cases the model reports unsupported, and why.
UNSUPPORTED = dict.fromkeys(("add_print_area", "add_print_area_edge"),
                            "breaks on a sheet with a print area, which Excel cuts to it and edges with automatic ones")
SUPPORTED = [name for name in NAMES if name not in UNSUPPORTED]

#: Excel's functions for the cases, before its Build, with one that makes a case alone, saves it and reads it back.
ONE = RECORD["build"].split("Public Function Build", 1)[0] + '''
Public Function One(index As Long, folder As String) As String
    Dim wb As Object, ws As Object, scratch As Object, out As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set scratch = wb.Worksheets(1)
    scratch.Name = "scratch"
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


def sheet_breaks(path: Path) -> dict[str, list[str]]:
    """Each sheet's rowBreaks and colBreaks as the file keeps them, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, list[str]] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            out[name] = re.findall(r"<(?:rowBreaks|colBreaks)\b.*?</(?:rowBreaks|colBreaks)>|"
                                   r"<(?:rowBreaks|colBreaks)\b[^>]*/>", text, re.DOTALL)
        return out


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_module(ONE, name="Cases")
    return made


@pytest.fixture(scope="module")
def made(app: ExcelApplication, tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[Any, ...]]:
    """Each supported case made in the model with Excel's function for it: its dump before the save, once the
    model opened its own file again, and every sheet's breaks as that file keeps them."""
    out: dict[str, tuple[Any, ...]] = {}
    for name in SUPPORTED:
        folder = tmp_path_factory.mktemp(name)
        try:
            session, reopened = str(app.run("One", NAMES.index(name), str(folder))).split("@@", 1)
        except Exception as exc:  # noqa: BLE001 -- one case's failure is that case's, not every case's.
            out[name] = (repr(exc), {}, {})
            continue
        out[name] = (dumps(session)[name], dumps(reopened), sheet_breaks(folder / "case.xlsx"))
    return out


@pytest.mark.parametrize("name", SUPPORTED)
def test_breaks_answer_as_excels_do(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """Each collection's Count and every break's Location, Type and Extent, each row's and column's PageBreak, and
    the errors the steps raise."""
    assert made[name][0] == RECORD["session"][name]


@pytest.mark.parametrize("name", SUPPORTED)
def test_breaks_read_back_from_the_file_the_model_saved(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """The sheet and any copy of it, as the model reads its own file again."""
    reopened = made[name][1]
    assert {sheet: reopened.get(sheet) for sheet in reopened if sheet.startswith(name)} == \
        {sheet: RECORD["reopened"][sheet] for sheet in RECORD["reopened"] if sheet.split(" (")[0] == name}


@pytest.mark.parametrize("name", SUPPORTED)
def test_the_file_keeps_the_breaks_as_excels_does(made: dict[str, tuple[Any, ...]], name: str) -> None:
    """The sheet's rowBreaks and colBreaks, and a copy's."""
    saved = made[name][2]
    assert {sheet: saved.get(sheet) for sheet in saved if sheet.startswith(name)} == \
        {sheet: RECORD["breaks"][sheet] for sheet in RECORD["breaks"] if sheet.split(" (")[0] == name}


@pytest.fixture(scope="module")
def excels_file_read() -> dict[str, str]:
    """Every sheet of the workbook Excel saved but those with a print area, as the model reads it."""
    reader = ExcelApplication()
    skip = " ".join(UNSUPPORTED)
    reader.add_module(RECORD["build"].split("Public Function Build", 1)[0] + f'''
Public Function ReadBack(path As String) As String
    Dim wb As Object, ws As Object, out As String
    Set wb = Workbooks.Open(path)
    For Each ws In wb.Worksheets
        If InStr(" {skip} ", " " & ws.Name & " ") = 0 Then out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    ReadBack = out
End Function
''', name="Reader")
    return dumps(str(reader.run("ReadBack", str(FIXTURES / "page_breaks.xlsx"))))


@pytest.mark.parametrize("name", [name for name in RECORD["reopened"] if name not in UNSUPPORTED])
def test_excels_file_reads_as_excel_reads_it(excels_file_read: dict[str, str], name: str) -> None:
    assert excels_file_read[name] == RECORD["reopened"][name]


@pytest.mark.parametrize("name", sorted(UNSUPPORTED))
def test_what_the_model_does_not_do_reports_itself(app: ExcelApplication, name: str, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        app.run("One", NAMES.index(name), str(tmp_path))
