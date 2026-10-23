"""A table's own AutoFilter, replayed against the in-memory model.

tests/fixtures/table_autofilter/ is what scripts/measure_table_autofilter.py
saw in live Excel: layouts that filter a table, each dumped as which rows
are hidden, the sheet's own filter and FilterMode, and each table's filter;
table_autofilter.xlsx, the workbook Excel saved of them all; each sheet as
Excel read it again from that file; a few steps taken on the file opened
again; and what the file holds of filters.

Excel ran every layout in one workbook, so its tables are Table1 to Table77,
where each run here has a workbook of its own: the dumps are compared with
a table's number taken out. The one thing compared loosely in the file is
the xr:uid a filter takes, a fresh random id each time.
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

FIXTURES = Path(__file__).parent / "fixtures" / "table_autofilter"
RECORD: dict[str, Any] = json.loads((FIXTURES / "table_autofilter.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: Layouts the model reports unsupported: showing or hiding a table's header row.
UNSUPPORTED = {"no_headers", "no_headers_filter"}
#: Tables resized smaller over data they leave under them, whose AutoFilter.Range Excel stretches back over that
#: data once the file is opened again, even past a blank row: the model answers the block the file gives.
STRETCHED = {"resize_filtered", "resize_unfiltered", "resize_gap", "resize_one_column"}
_UID = re.compile(r'xr:uid="\{[0-9A-F-]+\}"')


def _named(dump: str) -> str:
    return re.sub(r"\[Table\d+ ", "[Table ", dump)


def _loose(xml: str) -> str:
    return _UID.sub('xr:uid="{...}"', xml)


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant, o As Object",
            "Dim p As Object", "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{layout["name"][:31]}"',
            "On Error Resume Next", "Err.Clear", *(RECORD["table"] + layout["setup"]).splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & TableDump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + RECORD["dump"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", [one for one in LAYOUTS if one["name"] not in UNSUPPORTED],
                         ids=lambda layout: layout["name"])
def test_a_table_filters_as_excel_filters_it(layout: dict[str, Any]) -> None:
    assert _named(str(_run(layout))) == _named(layout["answers"])


@pytest.mark.parametrize("name", sorted(UNSUPPORTED))
def test_what_was_not_modelled_reports_itself(name: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(next(one for one in LAYOUTS if one["name"] == name))


def _opened(code: str) -> ExcelApplication:
    """Excel's file opened by the model, with the dump and ``code`` beside it."""
    app = ExcelApplication.open(FIXTURES / "table_autofilter.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + RECORD["dump"] + code, name="Reader")
    return app


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every sheet of the file Excel saved, as the model reads it."""
    app = _opened("Public Function Read() As String\nDim ws As Object, out As String\n"
                  "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & TableDump(ws) & \"|\"\n"
                  "Next\nRead = out\nEnd Function\n")
    return dict(one.split("^", 1) for one in str(app.run("Read")).split("|") if one)


@pytest.mark.parametrize("layout", [
    pytest.param(one, marks=pytest.mark.xfail(strict=True, reason="AutoFilter.Range stretched over data left under "
                                                                   "a resized table"))
    if one["name"] in STRETCHED else one for one in LAYOUTS], ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert _named(excel_file_read[layout["name"][:31]]) == _named(layout["reopened"])


def test_steps_on_the_file_opened_again_answer_as_excels() -> None:
    steps = RECORD["after_reopening"]
    code = ["Public Function Steps() As String", "Dim wb As Object, ws As Object, out As String",
            "Set wb = ActiveWorkbook"]
    bodies: list[str] = []
    for index, step in enumerate(steps):
        code += [f'Set ws = wb.Worksheets("{step["name"]}")', *(["ws.Activate"] if step["activated"] else []),
                 f'out = out & After{index}(ws) & "|"']
        bodies.append("\n".join([f"Private Function After{index}(ws As Object) As String",
                                 "Dim failed As String, v As Variant, o As Object, p As Object", "On Error Resume Next",
                                 "Err.Clear", "Set o = ws.ListObjects(1)", *step["step"].splitlines(),
                                 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
                                 f'After{index} = failed & Show(v) & ";" & TableDump(ws)', "End Function"]) + "\n")
    app = _opened("\n".join([*code, "Steps = out", "End Function"]) + "\n" + "".join(bodies))
    answers = str(app.run("Steps")).split("|")
    assert [_named(one) for one in answers[:len(steps)]] == [_named(step["answers"]) for step in steps]


def _file_parts(path: Path) -> dict[str, dict[str, object]]:
    """Each sheet's filterMode, hidden rows and autoFilter element, and its tables' autoFilter elements."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")

        def targets(rels: str) -> dict[str, tuple[str, str]]:
            found: dict[str, tuple[str, str]] = {}
            for element in re.findall(r"<Relationship\b[^>]*>", package.read(rels).decode("utf-8")):
                one = dict(re.findall(r'(\w+)="([^"]*)"', element))
                found[one["Id"]] = (one["Type"].rsplit("/", 1)[-1], one["Target"])
            return found

        def element(text: str, tag: str) -> str:
            found = re.search(rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>", text, re.DOTALL)
            return _loose(found.group()) if found else ""

        sheets = targets("xl/_rels/workbook.xml.rels")
        parts: dict[str, dict[str, object]] = {}
        for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = posixpath.normpath(posixpath.join("xl", sheets[rid][1].lstrip("/").removeprefix("xl/")))
            text = package.read(part).decode("utf-8")
            rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            links: list[tuple[str, str]] = list(targets(rels).values()) if rels in package.namelist() else []
            tables = sorted(element(package.read(posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
                                    .decode("utf-8"), "autoFilter")
                            for kind, target in links if kind == "table")
            parts[name] = {"filterMode": 'filterMode="1"' in element(text, "sheetPr"),
                           "autoFilter": element(text, "autoFilter"),
                           "hidden": [int(row) for row in re.findall(r'<row r="(\d+)"[^>]*\bhidden="1"', text)],
                           "tables": tables}
    return parts


def _excels(layout: dict[str, Any]) -> dict[str, object]:
    found = layout["file"]
    return {"filterMode": 'filterMode="1"' in found["sheetPr"], "autoFilter": _loose(found["autoFilter"]),
            "hidden": found["hidden"], "tables": sorted(_loose(one["autoFilter"]) for one in found["tables"])}


def saved_layouts(path: Path) -> None:
    """Every layout the model has, made in one workbook as Excel made them, and saved to ``path``. Every sheet is
    named first: the model renames no sheet once the workbook has a table."""
    layouts = [one for one in LAYOUTS if one["name"] not in UNSUPPORTED]
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    for index, layout in enumerate(layouts):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"][:31]}"']
    bodies: list[str] = []
    for index, layout in enumerate(layouts):
        build += [f'Set ws = wb.Worksheets("{layout["name"][:31]}")', "ws.Activate", f"Case{index} ws"]
        bodies.append("\n".join([f"Private Sub Case{index}(ws As Object)", "Dim v As Variant, o As Object, p As Object",
                                 "On Error Resume Next", *(RECORD["table"] + layout["setup"]).splitlines(),
                                 "End Sub"]) + "\n")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*build, "End Sub"]) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    app.save(path)


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[str, object]]:
    path = tmp_path_factory.mktemp("table_autofilter") / "filters.xlsx"
    saved_layouts(path)
    return _file_parts(path)


@pytest.mark.parametrize("layout", [one for one in LAYOUTS if one["name"] not in UNSUPPORTED],
                         ids=lambda layout: layout["name"])
def test_the_file_holds_the_filters_as_excels_does(model_file: dict[str, dict[str, object]],
                                                   layout: dict[str, Any]) -> None:
    assert model_file[layout["name"][:31]] == _excels(layout)
