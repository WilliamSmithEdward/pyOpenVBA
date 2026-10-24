"""Conditional formats Excel joins as it reads a file, replayed against the in-memory model.

tests/fixtures/conditional_joins/ is what scripts/measure_conditional_joins.py
saw in live Excel: layouts of rules a macro made over the numbers r * 10 + c
in A1:J20, each dumped as every rule with its priority, range and formulas,
then the cells of B1:B16 and A2:J2 that show red; the workbook saved as an
.xlsx, as an .xlsm with no code and as an .xlsm with a module, each sheet as
Excel read it again from each file, and each file as Excel saved it again.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest
from _conditional_files import formatting, guids

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "conditional_joins"
RECORD: dict[str, Any] = json.loads((FIXTURES / "conditional_joins.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
FILES: dict[str, str] = RECORD["files"]


def _code(name: str, setup: str) -> str:
    lines = ["Public Function Report() As String",
             "Dim failed As String, r As Long, c As Long, f As Object, g As Object, ws As Object",
             "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{name}"', "On Error Resume Next", "Err.Clear",
             *(RECORD["grid"] + setup).splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", "ws.Activate", "Report = failed & Dump(ws)", "End Function"]
    return RECORD["helper"] + "\n".join(lines) + "\n"


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_macro_makes_the_rules_as_excel_does(layout: dict[str, Any]) -> None:
    """Made in a session, alike rules stay apart."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_code(layout["name"], layout["setup"]), name="Probe")
    assert app.run("Report") == layout["answers"]


@pytest.fixture(scope="module", params=list(FILES))
def excel_file_read(request: pytest.FixtureRequest) -> tuple[str, ExcelApplication]:
    """Each file Excel saved, opened in the model, with a reader for one sheet at a time."""
    file: str = request.param
    app = ExcelApplication.open(FIXTURES / FILES[file], with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read(name As String) As String\n"
                   "Read = Dump(ActiveWorkbook.Worksheets(name))\nEnd Function\n", name="Reader")
    return file, app


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_joins_the_rules_of_a_file_as_excel_does(excel_file_read: tuple[str, ExcelApplication],
                                                            layout: dict[str, Any]) -> None:
    """Read from a file without a VBA project, alike rules over ranges making one block join; with one, none do."""
    file, app = excel_file_read
    assert app.run("Read", layout["name"]) == layout["reopened"][file]


@pytest.fixture(scope="module")
def resaved(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[str, str]]:
    """Each file Excel saved, opened in the model and saved again unchanged, as Excel saved it again."""
    folder = tmp_path_factory.mktemp("joins")
    out: dict[str, dict[str, str]] = {}
    for file, name in FILES.items():
        app = ExcelApplication.open(FIXTURES / name, with_vba=False)
        out[file] = formatting(app.save(folder / name))
    return out


@pytest.mark.parametrize("file", list(FILES))
@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_file_saved_again_keeps_the_rules_as_excel_does(resaved: dict[str, dict[str, str]], file: str,
                                                          layout: dict[str, Any]) -> None:
    """Saved again, a file read without a VBA project holds its rules joined and in Excel's order."""
    excel = formatting(FIXTURES / FILES[file].replace(".", "_resaved.", 1))
    assert guids(resaved[file][layout["name"]]) == guids(excel[layout["name"]])


@pytest.fixture(scope="module")
def excel_files() -> dict[str, dict[str, str]]:
    return {file: formatting(FIXTURES / name) for file, name in FILES.items()}


def test_the_record_holds_what_excels_files_hold(excel_files: dict[str, dict[str, str]]) -> None:
    for file in FILES:
        assert all(re.sub(r'dxfId="\d+"', "", one["file"][file])
                   == re.sub(r'dxf="<dxf>.*?</dxf>"', "", excel_files[file][one["name"]]) for one in LAYOUTS)


def made_layouts(path: Path) -> None:
    """Every layout, made in one workbook as Excel made them, and saved to ``path``."""
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    for index, layout in enumerate(LAYOUTS):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"]}"', f"Case{index} ws"]
    bodies = ["\n".join([f"Private Sub Case{index}(ws As Object)",
                         "Dim r As Long, c As Long, f As Object, g As Object", "On Error Resume Next",
                         *(RECORD["grid"] + layout["setup"]).splitlines(), "End Sub"]) + "\n"
              for index, layout in enumerate(LAYOUTS)]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*build, "End Sub"]) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    app.save(path)


@pytest.fixture(scope="module")
def model_file(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    path = tmp_path_factory.mktemp("joins_made") / "joins.xlsx"
    made_layouts(path)
    return formatting(path)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_saves_the_rules_as_excel_does(model_file: dict[str, str], excel_files: dict[str, dict[str, str]],
                                                 layout: dict[str, Any]) -> None:
    assert guids(model_file[layout["name"]]) == guids(excel_files["xlsx"][layout["name"]])
