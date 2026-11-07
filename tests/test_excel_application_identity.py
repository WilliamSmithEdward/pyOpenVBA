"""Application's paths and names, typed and late bound (scripts/measure_application_identity.py).

The folders are chosen, as Path is, since real ones name the machine and its user, so their reads are held to the type and the
folder's name in the profile. Everything else is held to Excel's answer.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/application_identity.json").read_text(encoding="utf-8"))
ROWS: dict[str, dict[str, str]] = RECORD["rows"]
SOURCE = """Function Go(ByVal folder As String) As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    {action}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
BINDINGS = {"typed": "Dim a As Application", "late": "Dim a As Object"}
#: Read by the machine: what holds is the type and the folder's place in the profile.
CHOSEN = {"read StartupPath", "read TemplatesPath", "read UserLibraryPath"}
#: Sets of folders, which need the folders on disk and a rule for each kind of value not yet split: these stay MISSING.
UNBUILT = ("set DefaultFilePath", "set AltStartupPath", "read DefaultFilePath", "read AltStartupPath")


def _action(key: str) -> str:
    kind, _, rest = key.partition(" ")
    if kind == "read":
        return f"v = a.{rest}"
    name, _, value = rest.partition(" = ")
    if not value:
        return f'a.{name} = "x"'
    return f"a.{name} = {value}: If Err.Number = 0 Then v = a.{name}"


CASES = [(key, binding) for key, row in ROWS.items() if not key.startswith(UNBUILT) for binding in row]


@pytest.mark.parametrize(("key", "binding"), CASES, ids=[f"{key}-{binding}" for key, binding in CASES])
def test_answers_as_excel_does(key: str, binding: str, tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE.format(dim=BINDINGS[binding], action=_action(key)), name="M")
    answer = str(app.run("Go", tmp_path.as_posix()))
    measured = ROWS[key][binding]
    if key in CHOSEN:
        assert answer.partition(":")[0] == measured.partition(":")[0]
        assert re.sub(r"^String:C:\\Users\\[^\\]+", "String:<home>", answer) == measured
    else:
        assert answer == measured
