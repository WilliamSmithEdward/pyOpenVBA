"""AutoFilters in the file, replayed against the in-memory model.

tests/fixtures/autofilter_file/ is what scripts/measure_autofilter_file.py
saw: autofilter.xlsx as Excel saved a sheet for each kind of filter, and,
in autofilter_file.json, each filter as Excel read it before the save and
again after opening the file, its autoFilter element, and the workbook's
names; names_order.xlsx mixes the filters' hidden names with names of both
scopes, to show the order Excel writes names in.

The one thing compared loosely is the xr:uid Excel gives a filter: a fresh
random id each time, which no two saves share.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "autofilter_file"
RECORD: dict[str, Any] = json.loads((FIXTURES / "autofilter_file.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
_UID = re.compile(r'xr:uid="\{[0-9A-F-]+\}"')


def _loose(xml: str) -> str:
    return _UID.sub('xr:uid="{...}"', xml)


def _elements(path: Path) -> tuple[dict[str, str], list[str]]:
    """Each sheet's autoFilter element by sheet name, and the workbook's definedName elements."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        relationships = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {rid: target for rid, target in re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', relationships)}
        targets |= {rid: target for target, rid in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', relationships)}
        found: dict[str, str] = {}
        for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            text = package.read("xl/" + targets[rid].lstrip("/").removeprefix("xl/")).decode("utf-8")
            element = re.search(r"<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>", text, re.DOTALL)
            found[name] = element.group() if element else ""
    return found, re.findall(r"<definedName\b.*?</definedName>", workbook)


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every sheet of the file Excel saved, as the model reads the filter on it."""
    app = ExcelApplication.open(FIXTURES / "autofilter.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & Dump(ws) & \"|\"\nNext\n"
                   "Read = out\nEnd Function\n", name="Reader")
    return dict(one.split("^", 1) for one in str(app.run("Read")).split("|") if one)


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_a_filter_excel_saved_reads_back_as_excel_reads_it(excel_file_read: dict[str, str],
                                                            case: dict[str, Any]) -> None:
    assert excel_file_read[case["name"]] == case["reopened"]


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> tuple[Path, list[str]]:
    """The same filters made in the model, what it read of each, and the file it saved."""
    build = ["Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Set wb = ActiveWorkbook"]
    bodies: list[str] = []
    for index, case in enumerate(CASES):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{case["name"]}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append("\n".join([f"Private Function Case{index}(ws As Object) As String",
                                 "Dim failed As String, v As Variant", "On Error Resume Next", "Err.Clear",
                                 *(RECORD["table"] + case["setup"]).splitlines(),
                                 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
                                 f'Case{index} = failed & Show(v) & ";" & Dump(ws)', "End Function"]) + "\n")
    build += ["Build = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(build) + "\n" + "".join(bodies), name="Probe")
    readings = str(app.run("Build")).split("|")[: len(CASES)]
    path = tmp_path_factory.mktemp("autofilter") / "autofilter.xlsx"
    app.save(path)
    return path, readings


@pytest.mark.parametrize("index", range(len(CASES)), ids=[case["name"] for case in CASES])
def test_a_model_filter_reads_as_excels_did(model_file: tuple[Path, list[str]], index: int) -> None:
    assert model_file[1][index] == CASES[index]["reading"]


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_a_model_filter_is_written_as_excel_writes_it(model_file: tuple[Path, list[str]],
                                                      case: dict[str, Any]) -> None:
    written, _ = _elements(model_file[0])
    assert _loose(written[case["name"]]) == _loose(case["xml"])


def test_the_filters_hidden_names_are_written_as_excel_writes_them(model_file: tuple[Path, list[str]]) -> None:
    assert _elements(model_file[0])[1] == RECORD["names"]


def test_names_are_written_in_excels_order(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
    Dim wb As Object, ws As Object
    Set wb = ActiveWorkbook
    wb.Worksheets(1).Name = "Zed"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Abc"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Mid"
    For Each ws In wb.Worksheets
        ws.Range("A1:B3").Value = 1
        ws.Range("A1:B3").AutoFilter
    Next
    wb.Names.Add Name:="beta", RefersTo:="=Zed!$A$1"
    wb.Names.Add Name:="Alpha", RefersTo:="=Zed!$A$2"
    wb.Names.Add Name:="zulu", RefersTo:="=Zed!$A$3"
    wb.Names.Add Name:="_under", RefersTo:="=Zed!$A$4"
    wb.Worksheets("Mid").Names.Add Name:="beta", RefersTo:="=Mid!$B$1"
    wb.Worksheets("Abc").Names.Add Name:="Local", RefersTo:="=Abc!$B$2"
End Sub
''', name="Names")
    app.run("Build")
    app.save(tmp_path / "names.xlsx")
    written = _elements(tmp_path / "names.xlsx")[1]
    # The model has no print area yet, so Excel's Print_Area is left out of the comparison.
    assert written == [one for one in RECORD["names_order"] if "Print_Area" not in one]


def test_a_filter_nobody_changed_is_written_back_as_it_was(tmp_path: Path) -> None:
    app = ExcelApplication.open(FIXTURES / "autofilter.xlsx", with_vba=False)
    app.add_module("Public Sub Touch()\nDim ws As Object\nFor Each ws In ActiveWorkbook.Worksheets\n"
                   'ws.Range("Z1").Value = 1\nNext\nEnd Sub\n', name="Touch")
    app.run("Touch")
    app.save(tmp_path / "again.xlsx")
    assert _elements(tmp_path / "again.xlsx")[0] == _elements(FIXTURES / "autofilter.xlsx")[0]
