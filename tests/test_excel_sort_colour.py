"""Sorting on a cell's colour or its font's, replayed against the in-memory model.

tests/fixtures/sort_colour/ is what scripts/measure_sort_colour.py saw in
live Excel: a list of eight rows, some filled or written in colour, sorted
by the Sort object one way or another, each dumped as the rows' names and
colours in order and the sort fields the sheet keeps; sort_colour.xlsx,
the workbook Excel saved of them all; each sheet as Excel read it again
from that file, and sorted again from the fields it read back; and each
sheet's sortState with the dxf each colour condition names.

The dxf ids themselves are compared through the dxfs they name: Excel
numbers them in its own order, at save.
"""

from __future__ import annotations

import json
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "sort_colour"
RECORD: dict[str, Any] = json.loads((FIXTURES / "sort_colour.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: Excel makes a Range.Sort the sheet's sort settings, its key the one field, and saves them with refs not yet
#: understood (A3:C9 for A1:C9 under a header); the model leaves the sheet's settings as they were.
RANGE_SORT = "range_sort_after"


def _gaps(layouts: list[dict[str, Any]]) -> list[Any]:
    return [pytest.param(one, marks=pytest.mark.xfail(strict=True, raises=AssertionError,
                                                      reason="Range.Sort as the sheet's sort settings"))
            if one["name"] == RANGE_SORT else one for one in layouts]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant, i As Long",
            "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{layout["name"]}"', "On Error Resume Next",
            "Err.Clear", *(RECORD["list"] + layout["setup"]).splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", _gaps(LAYOUTS), ids=lambda layout: layout["name"])
def test_a_colour_sort_orders_rows_as_excel_does(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]


def test_icons_report_themselves() -> None:
    with pytest.raises(VBAUnsupportedError, match="icons"):
        _run({"name": "icons", "setup": 'ws.Sort.SortFields.Add Key:=ws.Range("A2:A9"), SortOn:=xlSortOnIcon\n'})


@pytest.fixture(scope="module")
def excel_file_read() -> tuple[dict[str, str], dict[str, str]]:
    """Every sheet of the file Excel saved as the model reads it, and as it sorts again from what it read."""
    app = ExcelApplication.open(FIXTURES / "sort_colour.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & Dump(ws) & \"|\"\nNext\n"
                   "Read = out\nEnd Function\n"
                   "Public Function Again() As String\nDim ws As Object, out As String\nOn Error Resume Next\n"
                   "For Each ws In ActiveWorkbook.Worksheets\nErr.Clear\nws.Sort.Apply\n"
                   "out = out & ws.Name & \"^\" & IIf(Err.Number <> 0, \"S\" & Err.Number & \";\", \"\") & Dump(ws) & "
                   "\"|\"\nNext\nAgain = out\nEnd Function\n", name="Reader")

    def answered(text: object) -> dict[str, str]:
        return dict(one.split("^", 1) for one in str(text).split("|") if one)

    return answered(app.run("Read")), answered(app.run("Again"))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: tuple[dict[str, str], dict[str, str]],
                                                   layout: dict[str, Any]) -> None:
    assert excel_file_read[0][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_what_it_read_sorts_again_as_excels_did(excel_file_read: tuple[dict[str, str], dict[str, str]],
                                                 layout: dict[str, Any]) -> None:
    assert excel_file_read[1][layout["name"]] == layout["applied_again"]


def _conditions(sort_state: str, dxfs: list[str]) -> list[str]:
    """A sortState's ref and each of its conditions, with the dxf a condition names in place of its id."""
    if not sort_state:
        return []
    ref = re.search(r'<sortState\b[^>]*?\bref="([^"]*)"', sort_state)
    found = [f"ref={ref.group(1) if ref else ''}"]
    for condition in re.findall(r"<sortCondition\b[^>]*/>", sort_state):
        attributes = dict(re.findall(r'(\w+)="([^"]*)"', condition))
        dxf = attributes.pop("dxfId", None)
        found.append(" ".join(f"{key}={value}" for key, value in attributes.items()) +
                     (" " + dxfs[int(dxf)] if dxf is not None else ""))
    return found


def _saved(path: Path) -> dict[str, list[str]]:
    """Each sheet's sort conditions, or its table's, with their dxfs, as a file keeps them."""
    with zipfile.ZipFile(path) as package:

        def targets(rels: str) -> dict[str, str]:
            found: dict[str, str] = {}
            for element in re.findall(r"<Relationship\b[^>]*>", package.read(rels).decode("utf-8")):
                attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
                found[attributes["Id"]] = attributes["Target"]
            return found

        def sort_state(part: str) -> str:
            state = re.search(r"<sortState\b.*?</sortState>", package.read(part).decode("utf-8"), re.DOTALL)
            return state.group() if state else ""

        workbook = package.read("xl/workbook.xml").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
        listed = re.search(r"<dxfs\b.*?</dxfs>", styles, re.DOTALL)
        dxfs = re.findall(r"<dxf>.*?</dxf>", listed.group()) if listed else []
        sheets = targets("xl/_rels/workbook.xml.rels")
        found: dict[str, list[str]] = {}
        for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = posixpath.normpath(posixpath.join("xl", sheets[rid].lstrip("/").removeprefix("xl/")))
            text = sort_state(part)
            rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            if rels in package.namelist():
                text += "".join(sort_state(posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
                                for target in targets(rels).values() if "tables/" in target)
            found[name] = _conditions(text, dxfs)
    return found


def _excels(layout: dict[str, Any]) -> list[str]:
    return _conditions(layout["sort_state"], re.findall(r"<dxf>.*?</dxf>", RECORD["dxfs"]))


def saved_layouts(path: Path) -> None:
    """Every layout, made in one workbook as Excel made them, and saved to ``path``. Every sheet is named first:
    the model renames no sheet once the workbook has a table."""
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    for index, layout in enumerate(LAYOUTS):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"]}"']
    for index, layout in enumerate(LAYOUTS):
        build += [f'Set ws = wb.Worksheets("{layout["name"]}")', "ws.Activate", f"Case{index} ws"]
    bodies = ["\n".join([f"Private Sub Case{index}(ws As Object)", "Dim v As Variant, i As Long",
                         "On Error Resume Next", *(RECORD["list"] + layout["setup"]).splitlines(), "End Sub"]) + "\n"
              for index, layout in enumerate(LAYOUTS)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*build, "End Sub"]) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    app.save(path)


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> dict[str, list[str]]:
    path = tmp_path_factory.mktemp("sort_colour") / "sorted.xlsx"
    saved_layouts(path)
    return _saved(path)


@pytest.mark.parametrize("layout", _gaps(LAYOUTS), ids=lambda layout: layout["name"])
def test_the_file_keeps_the_sort_as_excels_does(model_file: dict[str, list[str]], layout: dict[str, Any]) -> None:
    assert model_file[layout["name"]] == _excels(layout)
