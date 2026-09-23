"""Events and document modules, replayed against the in-memory model.

tests/fixtures/events.json is what scripts/measure_events.py saw: the
handlers each sheet's module and ThisWorkbook got, the log every action
left when it ran on its own with events on, and what code reaches through
a document module. The model gets the same modules as document modules,
runs the same actions in the same order, and must leave the same log.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "events.json").read_text(encoding="utf-8"))
ACTIONS: list[dict[str, str]] = RECORD["actions"]
REACH: list[dict[str, str]] = RECORD["reach"]


def _probe() -> str:
    """The measuring script's actions module: each action on its own, the log emptied before it."""
    lines = ["Public Function Probe() As String", "Dim out As String", "Application.EnableEvents = True"]
    lines += [f'out = out & "{one["name"]}=" & A{index}() & vbLf' for index, one in enumerate(ACTIONS)]
    lines += [f'out = out & "{one["name"]}=" & R{index}() & vbLf' for index, one in enumerate(REACH)]
    lines += ["Probe = out", "End Function"]
    for index, one in enumerate(ACTIONS):
        lines += [f"Private Function A{index}() As String", "On Error GoTo Bad", 'Events = ""', one["action"],
                  f"A{index} = Events", "Exit Function", "Bad:", f'A{index} = Events & "!" & Err.Number',
                  "End Function"]
    for index, one in enumerate(REACH):
        body = one["code"] if "Reach =" in one["code"] else f"Reach = {one['code']}"
        lines += [f"Private Function R{index}() As String", "Dim Reach As Variant", "On Error GoTo Bad", body,
                  f"R{index} = Reach", "Exit Function", "Bad:", f'R{index} = "!" & Err.Number', "End Function"]
    return "\n".join(lines) + "\n"


@pytest.fixture(scope="module")
def answers() -> dict[str, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["log"], name="EventLog")
    # As the measuring script set the workbook up, with no handler yet to hear it.
    app.add_module("Public Sub Setup()\nWorksheets.Add After:=Worksheets(1)\nWorksheets(1).Activate\n"
                   'Worksheets(1).Range("D1").Formula = "=A1*2"\nEnd Sub\n', name="Setup")
    app.run("Setup")
    # Reaching the project through its CodeModule gave the added sheet its code name.
    app.workbook.sheets_[1].code_name = "Sheet2"
    app.add_module(RECORD["sheet1"], name="Sheet1", kind="document")
    app.add_module(RECORD["sheet2"], name="Sheet2", kind="document")
    app.add_module(RECORD["book"], name="ThisWorkbook", kind="document")
    app.add_module(_probe(), name="Actions")
    return dict(line.split("=", 1) for line in str(app.run("Probe")).split("\n") if "=" in line)


@pytest.mark.parametrize("index", range(len(ACTIONS)), ids=[one["name"] for one in ACTIONS])
def test_an_action_raises_the_events_excel_raises(answers: dict[str, str], index: int) -> None:
    case = ACTIONS[index]
    assert answers[case["name"]] == case["events"], case["action"]


@pytest.mark.parametrize("index", range(len(REACH)), ids=[one["name"] for one in REACH])
def test_a_document_module_reaches_what_excel_reaches(answers: dict[str, str], index: int) -> None:
    case = REACH[index]
    assert answers[case["name"]] == case["answer"], case["code"]


def test_the_sheet_module_names_the_sheet() -> None:
    assert RECORD["code_names"] == "Sheet1|Sheet2"


def test_a_workbook_file_brings_its_sheet_modules_as_the_sheets_code(tmp_path: Path) -> None:
    """A sheet module read from an xlsm is the sheet's code: its handler hears the sheet's changes."""
    import shutil

    from pyopenvba.excel import ExcelFile

    path = tmp_path / "handled.xlsm"
    shutil.copy(Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm", path)
    with ExcelFile(path) as host:
        host.set_module("Sheet1", "Private Sub Worksheet_Change(ByVal Target As Range)\n"
                                  "    Heard = Heard & Me.CodeName & \":\" & Target.Address(False, False) & \"|\"\n"
                                  "End Sub\n")
        host.set_module("Module1", "Public Heard As String\n\nPublic Function Probe() As String\n"
                                   "    Sheet1.Range(\"B2:C3\").Value = 1\n    Probe = Heard\nEnd Function\n")
        host.save()
    app = ExcelApplication.open(path)
    assert app.run("Module1.Probe") == "Sheet1:B2:C3|"
