"""Colour scales, replayed against the in-memory model.

tests/fixtures/colour_scales/ is what scripts/measure_colour_scales.py saw
in live Excel: layouts that put a colour scale over the numbers 0 to 100,
each dumped as the colour every cell shows and every rule with its
criteria; two gradients of 101 steps; colour_scales.xlsx, the workbook
Excel saved of them all; each sheet as Excel read it again from that file;
and the conditionalFormatting each sheet was saved with.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest
from _conditional_files import formatting

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "colour_scales"
RECORD: dict[str, Any] = json.loads((FIXTURES / "colour_scales.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String",
            "Dim ws As Object, failed As String, v As Variant, i As Long, f As Object, g As Object",
            "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{layout["name"]}"', "On Error Resume Next",
            "Err.Clear", *(RECORD["list"] + layout["setup"]).splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_colour_scale_colours_cells_as_excel_does(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every sheet of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "colour_scales.xlsx", with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & Dump(ws) & \"|\"\nNext\n"
                   "Read = out\nEnd Function\n", name="Reader")
    return dict(one.split("^", 1) for one in str(app.run("Read")).split("|") if one)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    """A lowest or highest criterion reads 0 until the scale colours a cell: the gradients' rules lie outside the
    cells the dump reads."""
    assert excel_file_read[layout["name"]] == layout["reopened"]


@pytest.fixture(scope="module")
def excel_file() -> dict[str, str]:
    return formatting(FIXTURES / "colour_scales.xlsx")


def saved_layouts(path: Path) -> None:
    """Every layout, made in one workbook as Excel made them, and saved to ``path``."""
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    for index, layout in enumerate(LAYOUTS):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"]}"', f"Case{index} ws"]
    bodies = ["\n".join([f"Private Sub Case{index}(ws As Object)", "Dim v As Variant, i As Long, f As Object, g As Object",
                         "On Error Resume Next", *(RECORD["list"] + layout["setup"]).splitlines(), "End Sub"]) + "\n"
              for index, layout in enumerate(LAYOUTS)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join([*build, "End Sub"]) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    app.save(path)


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    path = tmp_path_factory.mktemp("colour_scales") / "scales.xlsx"
    saved_layouts(path)
    return formatting(path)


def test_the_record_holds_what_excels_file_holds(excel_file: dict[str, str]) -> None:
    assert all(re.sub(r'dxfId="\d+"', "", one["file"]) == re.sub(r'dxf="<dxf>.*?</dxf>"', "", excel_file[one["name"]])
               for one in LAYOUTS)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_the_scale_as_excels_does(model_file: dict[str, str], excel_file: dict[str, str],
                                                 layout: dict[str, Any]) -> None:
    assert model_file[layout["name"]] == excel_file[layout["name"]]
