"""Names.Add and Names(...) with a table's structured references and with names of the same name on sheets,
replayed against the in-memory model.

tests/fixtures/names_add/ holds what scripts/measure_names_add.py saw in
live Excel: a workbook for each case, a sheet Data with a table Sales and a
sheet Other, the sheet the case names active; names added to the
workbook's Names or a sheet's, looked up by their text, an error's number
and description where a step raised one, and every name read back with
its RefersTo, RefersToR1C1 and Visible, before the save and once the file
was opened again.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "names_add"
RECORD: dict[str, Any] = json.loads((FIXTURES / "names_add.json").read_text(encoding="utf-8"))
NAMES = [case["name"] for case in RECORD["cases"]]

#: Excel's functions for the cases, each folder joined to its file by this system's separator, not Windows' alone.
ONE = RECORD["build"].split("Public Function Build", 1)[0].replace('folder & "\\', 'folder & "' + os.sep) + '''
Public Function One(index As Long, folder As String) As String
    Application.DisplayAlerts = False
    Select Case index
{cases}
    End Select
End Function
'''.replace("{cases}", "".join(f"        Case {index}: One = Make{index}(folder)\n" for index in range(len(NAMES))))


def carried(text: str) -> str:
    """Text as the harness carries a macro's answer home: a line break as a line feed, a tab as itself."""
    return text.replace("\r\n", "\n").replace("\r", "\n")


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> dict[str, tuple[str, str, str]]:
    """Each case made in the model with Excel's function for it: the errors, the names, and the names read again
    from the file the model saved."""
    folder = tmp_path_factory.mktemp("names_add")
    out: dict[str, tuple[str, str, str]] = {}
    for index, name in enumerate(NAMES):
        app = ExcelApplication()
        app.add_module(ONE, name="Cases")
        try:
            answer = carried(str(app.run("One", index, str(folder))))
        except Exception as exc:  # noqa: BLE001 -- one case's failure is that case's, not every case's.
            out[name] = (f"failed: {exc!r}", "", "")
            continue
        label, rest = answer.split("{{name}}", 1)
        session, reopened = rest.removesuffix("{{case}}").split("{{reopened}}")
        notes, listed = session.split("{{names}}", 1)
        out[label] = (notes, listed, reopened)
    return out


@pytest.mark.parametrize("name", NAMES)
def test_names_add_and_find_names_as_excel_does(made: dict[str, tuple[str, str, str]], name: str) -> None:
    case = next(case for case in RECORD["cases"] if case["name"] == name)
    assert made[name] == (case["notes"], case["names"], case["reopened"])
