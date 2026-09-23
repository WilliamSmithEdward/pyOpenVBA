"""What a number format shows, replayed against the in-memory model.

tests/fixtures/number_formats.json is what scripts/measure_number_formats.py
saw in live Excel: 156 format codes, each set on a cell in a column 60
characters wide, with 28 values written to it in turn and read back
through Range.Text and through WorksheetFunction.Text. The model runs the
same macro. Where the section showing a value has a ``*`` fill, what the
cell shows depends on its width in pixels, which the model does not
measure, so Range.Text says so instead.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import UndisplayableError, format_value, shown

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "number_formats.json").read_text(encoding="utf-8"))
CODES: list[dict[str, Any]] = RECORD["codes"]
VALUES: list[str] = RECORD["values"]
STEP, ITEM = "<^>", "<|>"


def _value(expression: str) -> object:
    if expression.startswith('"'):
        return expression[1:-1]
    return True if expression == "True" else float(expression)


def _fills(code: str, expression: str) -> bool:
    try:
        return shown(_value(expression), code)[1]
    except UndisplayableError:
        return False


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


#: The codes whose every value the model can read back; the rest have a fill somewhere.
PLAIN = [entry for entry in CODES if not any(_fills(entry["code"], value) for value in VALUES)]

SHOW = f'''Private Function Shown(code As String, values As Variant) As String
    Dim out As String, c As Object, v As Variant, i As Long
    Set c = Range("A1")
    On Error Resume Next
    c.Clear
    Err.Clear
    c.NumberFormat = code
    If Err.Number <> 0 Then out = out & "S" & Err.Number & "{ITEM}"
    For i = LBound(values) To UBound(values)
        Err.Clear
        c.Value = values(i)
        v = Empty
        v = c.Text
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "{ITEM}" Else out = out & v & "{ITEM}"
        Err.Clear
        v = Empty
        v = Application.WorksheetFunction.Text(values(i), code)
        If Err.Number <> 0 Then out = out & "E" & Err.Number & "{ITEM}" Else out = out & v & "{ITEM}"
    Next
    On Error GoTo 0
    Shown = out & "{STEP}"
End Function
'''


@pytest.fixture(scope="module")
def replayed() -> dict[str, tuple[list[str], list[str]]]:
    """Every plain code's reads in the model: Range.Text of each value, and TEXT of it."""
    lines = [SHOW, "Public Function Probe() As String", "Dim out As String, values As Variant",
             "Columns(1).ColumnWidth = 60", f"values = Array({', '.join(VALUES)})",
             *(f"out = out & Shown({_vba_text(entry['code'])}, values)" for entry in PLAIN),
             "Probe = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(lines) + "\n", name="Probe")
    answers = str(app.run("Probe")).split(STEP)[: len(PLAIN)]
    out: dict[str, tuple[list[str], list[str]]] = {}
    for entry, answer in zip(PLAIN, answers, strict=True):
        items = answer.split(ITEM)[:-1]
        out[entry["code"]] = (items[0::2], items[1::2])
    return out


@pytest.mark.parametrize("entry", PLAIN, ids=[entry["code"] for entry in PLAIN])
def test_a_cell_shows_what_excel_shows(replayed: dict[str, tuple[list[str], list[str]]],
                                       entry: dict[str, Any]) -> None:
    assert replayed[entry["code"]][0] == entry["text"]


@pytest.mark.parametrize("entry", PLAIN, ids=[entry["code"] for entry in PLAIN])
def test_text_shows_what_excel_shows(replayed: dict[str, tuple[list[str], list[str]]], entry: dict[str, Any]) -> None:
    assert replayed[entry["code"]][1] == entry["function"]


@pytest.mark.parametrize("entry", CODES, ids=[entry["code"] for entry in CODES])
def test_the_engine_answers_what_text_answers(entry: dict[str, Any]) -> None:
    """Every code, fills included: TEXT leaves a fill out, and fails where a cell would fill with #."""
    for expression, want in zip(VALUES, entry["function"], strict=True):
        try:
            got = format_value(_value(expression), entry["code"])
        except UndisplayableError:
            got = "E1004"
        assert got == want, expression


FILLED = [(entry, expression) for entry in CODES if entry not in PLAIN for expression in VALUES]


@pytest.mark.parametrize("entry,expression", FILLED,
                         ids=[f"{entry['code']} {expression}" for entry, expression in FILLED])
def test_a_filled_cell_says_what_it_cannot_show(entry: dict[str, Any], expression: str) -> None:
    """A value shown through a * fill reports itself; one shown without the fill reads as Excel's does."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Public Function Probe() As String\nColumns(1).ColumnWidth = 60\n"
                   f"Range(\"A1\").NumberFormat = {_vba_text(entry['code'])}\nRange(\"A1\").Value = {expression}\n"
                   "Probe = Range(\"A1\").Text\nEnd Function\n", name="Probe")
    if _fills(entry["code"], expression):
        with pytest.raises(VBAUnsupportedError, match=r"\* fill"):
            app.run("Probe")
    else:
        assert app.run("Probe") == entry["text"][VALUES.index(expression)]
