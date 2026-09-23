"""Range.RemoveDuplicates, replayed against the in-memory model.

tests/fixtures/remove_duplicates.json is what
scripts/measure_remove_duplicates.py saw in live Excel: layouts that fill
a few cells, remove duplicates from them and dump A1:E8 afterwards --
each cell's formula, the type of its value, its number format and bold --
with the used range and what the call returned or the error it raised.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "remove_duplicates.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]

#: Layouts the model reports it cannot do, and why.
UNSUPPORTED: dict[str, str] = {
    "accents": "text beyond ASCII compares by Windows' linguistic rules",
    "ligature": "text beyond ASCII compares by Windows' linguistic rules",
}


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_duplicates_go_as_excel_removes_them(layout: dict[str, Any]) -> None:
    if layout["name"] in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            _run(layout)
        return
    assert _run(layout) == layout["answers"]
