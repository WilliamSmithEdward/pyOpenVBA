"""Application.LanguageSettings and Application.VBE, typed and late bound, against live Excel
(scripts/measure_application_language_vbe.py, tests/fixtures/application_language_vbe.json).

Each read runs with two new workbooks open beside the macro's, the second active, as in Excel. The editor's windows, panes,
add-ins, command bars and events showed the measuring machine's editor, and the model refuses them.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/application_language_vbe.json").read_text(encoding="utf-8"))
ROWS: dict[str, dict[str, str]] = {key: row for key, row in RECORD["rows"].items() if key != "opened"}
OPEN = 'Function Opened() As String\n    Workbooks.Add\n    Workbooks.Add\n    Opened = Workbooks.Count & ":" & ActiveWorkbook.Name\nEnd Function\n'
SOURCE = """Function Walk(a As Object) As String
    Dim p As Object, s As String
    For Each p In a.VBE.VBProjects
        s = s & p.Name & "/" & (p Is ThisWorkbook.VBProject) & "/" & (p Is ActiveWorkbook.VBProject) & ";"
    Next
    Walk = s
End Function

Function Go() As String
    {dim}
    Dim v As Variant
    Set a = Application
    On Error Resume Next
    v = {read}
    If Err.Number <> 0 Then
        Go = "E" & Err.Number
    Else
        Go = TypeName(v) & ":" & CStr(v)
    End If
End Function
"""
BINDINGS = {"typed": "Dim a As Application", "late": "Dim a As Object"}
#: The measuring machine's editor: refused.
REFUSED = ("a.VBE.MainWindow", "a.VBE.Windows", "a.VBE.CodePanes", "a.VBE.Addins", "a.VBE.Events", "a.VBE.CommandBars")
CASES = [(read, binding) for read, row in ROWS.items() for binding in row]


def _run(read: str, binding: str) -> object:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(OPEN, name="Opener")
    assert app.run("Opened") == RECORD["rows"]["opened"]["late"]
    app.add_module(SOURCE.format(dim=BINDINGS[binding], read=read), name="M")
    return app.run("Go")


@pytest.mark.parametrize(("read", "binding"), CASES, ids=[f"{read}-{binding}" for read, binding in CASES])
def test_reads_as_excel_does(read: str, binding: str) -> None:
    if any(read.removeprefix("TypeName(").startswith(refused) for refused in REFUSED):
        with pytest.raises(VBAUnsupportedError, match="VBE"):
            _run(read, binding)
    else:
        assert _run(read, binding) == ROWS[read][binding]
