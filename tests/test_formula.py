"""The calculation engine against what Excel computed.

Every probe in ``fixtures/formula/probes.txt`` was written into a cell
of a freshly added worksheet in live Excel, and what that cell showed is
in ``measured.json``.  This replays each one through the engine and
requires the same answer, including which error came out.

The probes are chosen for the places a formula engine goes wrong:
whether a range contributes its text, which way a half rounds, what an
empty cell equals, how an error travels, and where a relative reference
lands when a formula is written to a block.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue, VBADate, to_text, type_name

FIXTURES = Path(__file__).parent / "fixtures" / "formula"

#: Where the measuring script puts each probe's formula.
CELL = "H1"


def split_probe(line: str) -> tuple[str, str]:
    setup, marker, formula = line.partition(";;")
    return (setup.strip(), formula.strip()) if marker else ("", line.strip())


def read_probes() -> list[str]:
    out: list[str] = []
    for line in (FIXTURES / "probes.txt").read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            out.append(stripped)
    return out


def read_measured() -> dict[str, str]:
    return json.loads((FIXTURES / "measured.json").read_text(encoding="utf-8"))["probes"]


def describe(app: ExcelApplication, value: object) -> str:
    """The same description the VBA side of the measurement writes."""
    if isinstance(value, VBAErrorValue):
        shown = app.evaluate('Range("' + CELL + '").Text')
        return f"Error|{to_text(shown)}"
    if isinstance(value, VBADate):
        return f"Date|{to_text(value)}"
    if value is EMPTY:
        return "Empty|"
    return f"{type_name(value)}|{to_text(value)}"


def evaluate(probe: str) -> str:
    setup, formula = split_probe(probe)
    app = ExcelApplication()
    app.add_workbook()
    source = "Sub Probe()\n"
    if setup:
        source += f"    {setup}\n"
    quoted = formula.replace('"', '""')
    source += f'    Range("{CELL}").Formula = "{quoted}"\n'
    source += "End Sub\n"
    try:
        app.add_module(source, name="Probe")
        app.run("Probe")
        target = app.workbook.sheets_[0].vba_get("Range", [CELL])
        return describe(app, target.vba_get("Value"))  # type: ignore[union-attr]
    except VBAUnsupportedError as gap:
        return f"!unsupported: {gap}"
    except VBARuntimeError as failure:
        # As the measuring macro records it: a formula Excel will not take is error 1004.
        return f"!{failure.number}"


PROBES = read_probes()
MEASURED = read_measured()

#: Probes the engine does not answer as Excel does yet, and why.
GAPS: dict[str, str] = {}


def test_every_probe_was_measured() -> None:
    missing = [probe for probe in PROBES if probe not in MEASURED]
    assert not missing, f"{len(missing)} probes have no measured answer, first {missing[0]!r}"
    assert set(GAPS) <= set(PROBES)


def _cases() -> list[object]:
    return [pytest.param(probe, id=str(index),
                         marks=[pytest.mark.xfail(reason=GAPS[probe], strict=True)] if probe in GAPS else [])
            for index, probe in enumerate(PROBES)]


@pytest.mark.parametrize("probe", _cases())
def test_matches_excel(probe: str) -> None:
    want = MEASURED[probe]
    got = evaluate(probe)
    assert got == want, f"{probe}\n  Excel: {want}\n  ours : {got}"


def test_an_e_in_a_number_format_is_the_year() -> None:
    """In pyOfficeEditor's formula corpus Excel shows TEXT(0,"0;-0;zero") as z1900ro: the e is the era year."""
    assert evaluate('=TEXT(0,"0;-0;zero")') == "String|z1900ro"


def test_an_unimplemented_function_says_so_rather_than_answering_name() -> None:
    """A gap here has to look different from a function Excel lacks."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Main()\n    Range("A1").Formula = "=BAHTTEXT(1)"\nEnd Sub\n', name="Module1")
    app.run("Main")
    with pytest.raises(VBAUnsupportedError) as raised:
        app.sheet(1).value("A1")
    assert "BAHTTEXT" in str(raised.value)


def test_a_function_excel_has_not_got_either_is_a_name_error() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Main()\n    Range("A1").Formula = "=NOTAFUNCTION(1)"\nEnd Sub\n', name="Module1")
    app.run("Main")
    assert app.evaluate('Range("A1").Text') == "#NAME?"
    assert isinstance(app.workbook.sheets_[0].cells_[(1, 1)].value, ExcelError)


#: Every function the engine has, written on its own and into COUNTIF's range in live Excel
#: (scripts/measure_cells_functions.py).
CELLS_FUNCTIONS: dict[str, dict[str, object]] = json.loads(
    (FIXTURES / "cells_functions.json").read_text(encoding="utf-8"))

#: The same macro the measuring script runs for each function.
_CELLS_PROBE = """Public Function Probe(call As String) As String
    Dim v As Variant, shown As String
    Range("A1").Value = 1
    On Error Resume Next
    Err.Clear
    Range("H1").Formula = "=" & call
    If Err.Number <> 0 Then
        Probe = "!" & Err.Number & "|"
    Else
        v = Range("H1").Value
        If IsError(v) Then shown = CStr(v) Else shown = "ok"
        Err.Clear
        Range("H2").Formula = "=COUNTIF(" & call & ",1)"
        Probe = shown & "|" & Err.Number
    End If
    Range("H1:H2").ClearContents
End Function
"""


_ONE_CELL_DATABASE = "a database function over one cell is #VALUE! in Excel"
#: Functions whose call here is not answered as Excel answers it, and why.
CELLS_GAPS: dict[str, str] = {
    "ACCRINTM": "an issue date on the settlement date is answered in Excel and #NUM! here",
    "AMORDEGRC": "a life of one year is #NUM! in Excel",
    "ANCHORARRAY": "Excel does not take ANCHORARRAY written as a call, the file's spelling of A1#",
    "BINOM.INV": "one trial with a chance of 1 and an alpha of 1 is #NUM! in Excel; which of the two it refuses is not "
                 "measured",
    "CRITBINOM": "one trial with a chance of 1 and an alpha of 1 is #NUM! in Excel; which of the two it refuses is not "
                 "measured",
    "CELL": "an info type that is not text is #VALUE! in Excel before CELL reads a cell",
    **dict.fromkeys(("DAVERAGE", "DCOUNT", "DCOUNTA", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR",
                     "DVARP"), _ONE_CELL_DATABASE),
    "DROP": "dropping every row is #CALC! in Excel",
    "LAMBDA": "the model does not read a LAMBDA called where it is written",
    "XIRR": "one value on one date is #N/A in Excel",
}


@pytest.fixture(scope="module")
def cells_app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_CELLS_PROBE, name="Probe")
    return app


def _cells_cases() -> list[object]:
    return [pytest.param(name, id=name, marks=[pytest.mark.xfail(reason=CELLS_GAPS[name], strict=True)]
                         if name in CELLS_GAPS else []) for name in sorted(CELLS_FUNCTIONS)]


def test_every_cells_gap_is_a_measured_function() -> None:
    assert set(CELLS_GAPS) <= set(CELLS_FUNCTIONS)


@pytest.mark.parametrize("name", _cells_cases())
def test_a_function_is_taken_where_cells_are_wanted_as_excel_takes_it(cells_app: ExcelApplication, name: str) -> None:
    measured = CELLS_FUNCTIONS[name]
    want = f"{measured['alone']}|{'' if measured['in_cells'] is None else measured['in_cells']}"
    try:
        got = str(cells_app.run("Probe", measured["call"]))
    except VBAUnsupportedError as gap:
        got = f"unsupported: {gap}"
    assert got == want, f"{measured['call']}: Excel {want}, ours {got}"
