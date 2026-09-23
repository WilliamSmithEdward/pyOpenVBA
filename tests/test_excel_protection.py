"""Sheet protection, replayed against the in-memory model.

tests/fixtures/protection.json is what scripts/measure_protection.py saw
in live Excel: each case protects a fresh sheet some way, tries an edit
under On Error, and records the error number and description it raised
and the protection flags and A1 it left. The model runs the same code.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "protection.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]

#: Cases the model does not answer as Excel does, and why.
KNOWN = {"autofit_allowed": "AutoFit of a column holding text measures the text, which the model does not"}


def _module(case: dict[str, Any]) -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, out As String, ws As Object",
             "Set wb = ActiveWorkbook", "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
             f'out = out & "{case["name"]}" & "="', "On Error Resume Next", "Err.Clear"]
    if case["setup"]:
        lines += [case["setup"], 'If Err.Number <> 0 Then out = out & "setup " & Err.Number & ";"', "Err.Clear"]
    lines += [case["action"], 'out = out & "E" & Err.Number & ":" & Err.Description & ";"', "Err.Clear",
              f"out = out & {RECORD['state']}", "On Error GoTo 0", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def _run(case: dict[str, Any]) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(case), name="Probe")
    return str(app.run("Probe")).split("=", 1)[1]


@pytest.mark.parametrize("case", [
    pytest.param(case, marks=[pytest.mark.xfail(reason=KNOWN[case["name"]], strict=True)] if case["name"] in KNOWN
                 else [], id=case["name"]) for case in CASES])
def test_protection_answers_as_excels_does(case: dict[str, Any]) -> None:
    assert _run(case) == case["answer"]


@pytest.mark.parametrize("action", ["ws.Unprotect", "ws.Protect Contents:=False"])
def test_a_password_excel_would_ask_for_says_so(action: str) -> None:
    """Excel prompts for the password in a dialog, which a macro here cannot answer."""
    case = {"name": "prompt", "setup": 'ws.Protect "pw"', "action": action}
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(case), name="Probe")
    with pytest.raises(VBAUnsupportedError, match="dialog"):
        app.run("Probe")
