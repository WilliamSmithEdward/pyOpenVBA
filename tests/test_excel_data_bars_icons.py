"""Data bars and icon sets, replayed against the in-memory model.

tests/fixtures/data_bars_icons/ is what scripts/measure_data_bars_icons.py
saw in live Excel: layouts that put a data bar, an icon set or a colour
scale over the numbers 0 to 100, each dumped as every rule with its points,
colours, criteria and settings; data_bars_icons.xlsx, the workbook Excel
saved of them all; each sheet as Excel read it again from that file and the
conditionalFormatting and extLst it was saved with; and
data_bars_icons_resaved.xlsx, the same file changed on a few sheets once
Excel had opened it and saved again.
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

FIXTURES = Path(__file__).parent / "fixtures" / "data_bars_icons"
RECORD: dict[str, Any] = json.loads((FIXTURES / "data_bars_icons.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
RESAVED: list[dict[str, Any]] = RECORD["resaved"]
#: Layouts whose formulas name the first sheet, which a lone layout needs beside it.
NAMING = {one["name"] for one in LAYOUTS if "bar_default!" in one["setup"]}
#: What the model reports itself unsupported for: the colours of a scale whose highest point is made 0, below
#: its middle.
UNSUPPORTED = {"scale_t_max_number"}
_DECLARATIONS = re.compile(r' xmlns:(?:x14|xm)="[^"]*"')


def _code(name: str, setup: str) -> list[str]:
    return ["Public Function Report() As String",
            "Dim ws As Object, failed As String, v As Variant, i As Long, f As Object, g As Object, h As Object, "
            "k As Object", "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{name}"',
            *(['ActiveWorkbook.Worksheets.Add(After:=ws).Name = "bar_default"',
               'For i = 1 To 11: ActiveWorkbook.Worksheets("bar_default").Cells(i, 1).Value = (i - 1) * 10: Next']
              if name in NAMING else []),
            "On Error Resume Next", "Err.Clear", *(RECORD["list"] + setup).splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]


def _run(layout: dict[str, Any]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(_code(layout["name"], layout["setup"])) + "\n", name="Probe")
    return app.run("Report")


def _marked(layouts: list[dict[str, Any]], unsupported: set[str]) -> list[Any]:
    return [pytest.param(one, marks=pytest.mark.xfail(strict=True, raises=VBAUnsupportedError,
                                                     reason="reported unsupported"))
            if one["name"] in unsupported else one for one in layouts]


@pytest.mark.parametrize("layout", _marked(LAYOUTS, UNSUPPORTED), ids=lambda layout: layout["name"])
def test_a_rule_answers_as_excel_does(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]


@pytest.fixture(scope="module")
def excel_file_read() -> ExcelApplication:
    """The file Excel saved, opened in the model, with a reader for one sheet at a time."""
    app = ExcelApplication.open(FIXTURES / "data_bars_icons.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read(name As String) As String\n"
                   "Read = Dump(ActiveWorkbook.Worksheets(name))\nEnd Function\n", name="Reader")
    return app


@pytest.mark.parametrize("layout", _marked(LAYOUTS, UNSUPPORTED), ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: ExcelApplication, layout: dict[str, Any]) -> None:
    assert excel_file_read.run("Read", layout["name"]) == layout["reopened"]


def _formatting(path: Path) -> dict[str, str]:
    """Each sheet's conditionalFormatting elements and its own extLst as a file keeps them: each dxfId the dxf it
    names, since Excel numbers its dxfs in its own order at save, and each GUID by the order it first appears, but
    those ``kept`` names as they are."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        targets: dict[str, str] = {}
        for element in re.findall(r"<Relationship\b[^>]*>", package.read("xl/_rels/workbook.xml.rels").decode()):
            attributes = dict(re.findall(r'(\w+)="([^"]*)"', element))
            targets[attributes["Id"]] = attributes["Target"]
        listed = re.search(r"<dxfs\b.*?</dxfs>", package.read("xl/styles.xml").decode("utf-8"), re.DOTALL)
        dxfs = re.findall(r"<dxf>.*?</dxf>", listed.group()) if listed else []
        found: dict[str, str] = {}
        for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
            part = posixpath.normpath(posixpath.join("xl", targets[rid].lstrip("/").removeprefix("xl/")))
            text = package.read(part).decode("utf-8")
            groups = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", text, re.DOTALL)
            rest = re.sub(r"<conditionalFormatting\b.*?</conditionalFormatting>", "", text, flags=re.DOTALL)
            kept = "".join(groups) + "".join(re.findall(r"<extLst>.*?</extLst>", rest, re.DOTALL))
            found[name] = re.sub(r'dxfId="(\d+)"', lambda named: f'dxf="{dxfs[int(named.group(1))]}"', kept)
    return found


def _guids(text: str, kept: frozenset[str] = frozenset()) -> str:
    """``text`` with each GUID not in ``kept`` named by the order it first appears; the ids Excel makes for rules it
    read without one stay as they are, being Excel's own and not random."""
    seen: dict[str, str] = {}

    def name(found: re.Match[str]) -> str:
        guid = found.group()
        if guid in kept or guid.startswith("{00000000-000E-0000-"):
            return guid
        return seen.setdefault(guid, f"{{GUID-{len(seen) + 1}}}")

    return re.sub(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}", name, text)


@pytest.fixture(scope="module")
def excel_file() -> dict[str, str]:
    return _formatting(FIXTURES / "data_bars_icons.xlsx")


def saved_layouts(path: Path) -> None:
    """Every layout, made in one workbook as Excel made them, and saved to ``path``."""
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    for index, layout in enumerate(LAYOUTS):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"]}"', f"Case{index} ws"]
    bodies = ["\n".join([f"Private Sub Case{index}(ws As Object)",
                         "Dim v As Variant, i As Long, f As Object, g As Object, h As Object, k As Object",
                         "On Error Resume Next", *(RECORD["list"] + layout["setup"]).splitlines(), "End Sub"]) + "\n"
              for index, layout in enumerate(LAYOUTS)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join([*build, "End Sub"]) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    app.save(path)


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    path = tmp_path_factory.mktemp("data_bars_icons") / "bars.xlsx"
    saved_layouts(path)
    return _formatting(path)


def test_the_record_holds_what_excels_file_holds(excel_file: dict[str, str]) -> None:
    assert all(re.sub(r'dxfId="\d+"', "", one["file"]) == re.sub(r'dxf="<dxf>.*?</dxf>"', "", excel_file[one["name"]])
               for one in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_the_rules_as_excels_does(model_file: dict[str, str], excel_file: dict[str, str],
                                                 layout: dict[str, Any]) -> None:
    assert _guids(model_file[layout["name"]]) == _guids(excel_file[layout["name"]])


@pytest.fixture(scope="module")
def resaved(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str]]:
    """Excel's file opened in the model, each sheet RESAVED names changed as Excel changed it, and saved again:
    what each sheet answered then, and the file."""
    app = ExcelApplication.open(FIXTURES / "data_bars_icons.xlsx", with_vba=False)
    steps = ["\n".join([f"Private Function Resave{index}(ws As Object) As String", "Dim failed As String",
                        "On Error Resume Next", "Err.Clear", *one["step"].splitlines(),
                        'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
                        f"Resave{index} = failed & Dump(ws)", "End Function"]) + "\n"
             for index, one in enumerate(RESAVED)]
    calls = [f'out = out & "{one["name"]}" & "^" & Resave{index}(ActiveWorkbook.Worksheets("{one["name"]}")) & "|"'
             for index, one in enumerate(RESAVED)]
    app.add_module(RECORD["helper"] + "".join(steps) + "Public Function Change() As String\nDim out As String\n"
                   + "\n".join(calls) + "\nChange = out\nEnd Function\n", name="Changer")
    answers = dict(one.split("^", 1) for one in str(app.run("Change")).split("|") if one)
    path = tmp_path_factory.mktemp("resaved") / "resaved.xlsx"
    app.save(path)
    return answers, _formatting(path)


#: Every GUID Excel's first file holds, which a save of it keeps.
KEPT = frozenset(re.findall(r"\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}",
                            json.dumps([one["file"] for one in LAYOUTS])))


@pytest.mark.parametrize("change", RESAVED, ids=lambda change: change["name"])
def test_a_changed_file_answers_as_excels_does(resaved: tuple[dict[str, str], dict[str, str]],
                                               change: dict[str, Any]) -> None:
    assert resaved[0][change["name"]] == change["answers"]


@pytest.fixture(scope="module")
def excel_resaved() -> dict[str, str]:
    return _formatting(FIXTURES / "data_bars_icons_resaved.xlsx")


@pytest.mark.parametrize("change", RESAVED, ids=lambda change: change["name"])
def test_a_changed_file_keeps_the_rules_as_excels_does(resaved: tuple[dict[str, str], dict[str, str]],
                                                       excel_resaved: dict[str, str], change: dict[str, Any]) -> None:
    assert _guids(resaved[1][change["name"]], KEPT) == _guids(excel_resaved[change["name"]], KEPT)
