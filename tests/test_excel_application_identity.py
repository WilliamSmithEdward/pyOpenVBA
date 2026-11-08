"""Application's paths and names, typed and late bound (scripts/measure_application_identity.py).

The folders are chosen, as Path is, since real ones name the machine and its user, so their reads are held to the type and the
folder's name in the profile. Everything else is held to Excel's answer.
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError

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
#: Sets of folders that ran one after another in one Excel, so each reads what the one before left: application_folders.json
#: measures them again, each from empty, and is replayed below. DefaultFilePath read what an earlier run had left, a folder
#: since deleted; the model gives the profile's Documents, where Excel keeps it as it ships.
UNBUILT = ("set DefaultFilePath", "set AltStartupPath", "read DefaultFilePath")


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


def test_default_file_path_starts_in_the_profile_documents() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE.format(dim=BINDINGS["typed"], action="v = a.DefaultFilePath"), name="M")
    assert app.run("Go", "") == "String:C:\\Users\\User\\Documents"


#: Each set from empty, with folder a folder holding a folder Sub and a file f.txt (scripts/measure_application_folders.py).
FOLDERS: dict[str, dict[str, str]] = json.loads((Path(__file__).parent / "fixtures/application_folders.json").read_text(encoding="utf-8"))["rows"]
FOLDER_SOURCE = """Function Go(ByVal folder As String) As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    a.{name} = ""
    Err.Clear
    a.{name} = {value}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        v = a.{name}
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
FOLDER_CASES = [(key, binding) for key, row in FOLDERS.items() if key.startswith("set ") for binding in row]


@pytest.mark.parametrize(("key", "binding"), FOLDER_CASES, ids=[f"{key}-{binding}" for key, binding in FOLDER_CASES])
def test_a_folder_set_keeps_what_excel_kept(key: str, binding: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    name, _, value = key.removeprefix("set ").partition(" = ")
    measured = FOLDERS[key][binding]
    if os.name != "nt" and ("folder" in value or "\\" in measured.replace("Office16\\", "")):
        pytest.skip("a Windows path, which Excel's folders are")
    folder = tmp_path / "Mixed"
    (folder / "Sub").mkdir(parents=True)
    (folder / "f.txt").write_text("x", encoding="utf-8")
    # The current folder, as the harness's was, holds no folder Sub.
    monkeypatch.chdir(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(FOLDER_SOURCE.format(dim=BINDINGS[binding], name=name, value=value), name="M")
    given = str(folder)
    wanted = measured.replace("<folder:lower>", given.lower()).replace("<FOLDER>", given.upper()).replace("<folder>", given)
    assert app.run("Go", given) == wanted


def test_the_sandbox_refuses_a_default_file_path_it_would_look_for() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Go()\n    Application.DefaultFilePath = ""\n    Application.DefaultFilePath = "C:\\"\nEnd Function\n', name="M")
    with pytest.raises(VBASandboxError, match="DefaultFilePath"):
        app.run("Go")
