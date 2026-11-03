"""What Application hands a macro with no screen, read as Excel reads it (scripts/measure_application_objects.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/application_objects.json").read_text(encoding="utf-8"))
ROWS: dict[str, str] = MEASURED["rows"]
IMPLEMENTED = {"AutoRecover", "MultiThreadedCalculation", "SpellingOptions", "AutoCorrect", "DefaultWebOptions",
               "DefaultPivotTableLayoutOptions", "FileValidation", "FileValidationPivot", "WindowsForPens",
               "ClusterConnector", "DDEAppReturnCode", "AlwaysUseClearType", "FormulaBarHeight", "RegisteredFunctions",
               "MailSession", "FileConverters", "ActiveEncryptionSession", "SensitivityLabelPolicy"}
#: Not a member Excel's CallByName finds, or a part of Application this slice leaves out.
LEFT_OUT = {"MultiThreadedCalculation.Mode"}
CHECKED = sorted(name for name in ROWS if name.split(".")[0] in IMPLEMENTED and name not in LEFT_OUT)


def _read(name: str) -> str:
    owner, _, part = name.rpartition(".")
    fetch = MEASURED["probe"].split("Public Function Probe")[0]
    target = f'CallByName(Application, "{owner}", VbGet)' if owner else "Application"
    source = fetch + f'Public Function Answer() As String\nAnswer = Fetch({target}, "{part}")\nEnd Function\n'
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name="Probe")
    return str(app.run("Answer"))


@pytest.mark.parametrize("name", CHECKED)
def test_reads_as_excel_does(name: str) -> None:
    assert _read(name) == ROWS[name]


def test_build_this_workbook_and_display_alerts() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["probe"], name="Probe")
    assert app.run("Direct") == MEASURED["direct"]


def test_setting_an_option_reports_itself() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Sub Go()\nApplication.AutoRecover.Time = 5\nEnd Sub\n", name="M")
    with pytest.raises(VBAUnsupportedError):
        app.run("Go")
