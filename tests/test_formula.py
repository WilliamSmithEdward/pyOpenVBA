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

_AMORDEGRC = "Excel's AMORDEGRC has other coefficients and ends an asset's life otherwise, not yet worked out"

#: Probes the engine does not answer as Excel does yet, and why.
GAPS: dict[str, str] = {
    **dict.fromkeys((f"=AMORDEGRC({arguments})" for arguments in (
        "1,1,1,1,1,1", "1000,1,100,100,1,2", "1000,1,100,100,1,1", "1000,1,100,100,1,0.5", "1000,1,100,100,1,0.4",
        "1000,1,100,100,1,0.22")), _AMORDEGRC),
}


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


def _irr(flows: list[float], guess: str = "") -> tuple[object, object]:
    """IRR over flows written to row 1, as a cell works it out: its value and its text."""
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    for column, flow in enumerate(flows):
        sheet.set_value(f"{chr(ord('A') + column)}1", flow)
    last = chr(ord("A") + len(flows) - 1)
    sheet.set_value("A3", f"=IRR(A1:{last}1{',' + guess if guess else ''})")
    return sheet.value("A3"), app.evaluate('Range("A3").Text')


# IRR follows Excel's own secant iteration, stopping where it stops: measured to the bit in live Excel by
# pyOfficeEditor (its commit ed27bf2, tests/test_excel_calc.py), whose engine this is.


def test_irr_stops_short_of_the_root_where_excel_does() -> None:
    flows = [-23518.90552495122, 5897.322375982933, 8074.228465462295, 3418.6542983744653, 6167.091710440405]
    assert _irr(flows, "0.3")[0] == 0.0006747819308212666


def test_irr_starts_again_from_a_tenth_when_its_guess_fails() -> None:
    flows = [-24518.42305940892, 1937.6740514260678, 1767.627855347036, 1621.8383666307395, 3423.8749530313867,
             4005.481458699994, 3303.67042464501, 4648.740185700573, 2305.081428575526, 7245.145640629922,
             5949.791802306586]
    assert _irr(flows, "1")[0] == 0.06297805183414074


def test_irr_needs_a_residual_under_its_tolerance() -> None:
    assert _irr([-1069510457926.0002, 1099511627776.0])[1] == "#NUM!"
    assert _irr([-977338919173.8009, 1099511627776.0], "0.05")[0] == 0.12500546760736642


def _bond(function: str, arguments: list[float]) -> object:
    """A bond function over arguments written to row 1, where no formula literal cuts them to fifteen digits."""
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    cells = [f"{chr(ord('A') + column)}1" for column in range(len(arguments))]
    for cell, argument in zip(cells, arguments, strict=True):
        sheet.set_value(cell, argument)
    sheet.set_value("A3", f"={function}({','.join(cells)})")
    return sheet.value("A3")


# PRICE, DURATION and MDURATION give Excel's bits on every bond probed in live Excel by pyOfficeEditor (its commit
# 81481d2, tests/test_excel_calc.py): 1,500 PRICE probes and 1,800 each for DURATION and MDURATION.


def test_duration_times_coupons_as_price_does_and_the_redemption_its_own_way() -> None:
    # A coupon's time is index + DSC/E; the redemption's is DSC/E + N - 1, which rounds another way.
    assert _bond("DURATION", [44468.0, 44647.0, 0.035429952036797184, 0.16067767704236355, 4.0, 2.0]) \
        == 0.492182006415991
    # Settled on a coupon date: each coupon weighs its time times its present value.
    assert _bond("DURATION", [49383.0, 50844.0, 0.13249040076812385, 0.030546357927438164, 2.0, 1.0]) \
        == 3.3581234390639882


def test_mduration_divides_by_the_growth() -> None:
    assert _bond("MDURATION", [41018.0, 41334.0, 0.10018741153059099, 0.11925232538933145, 4.0, 0.0]) \
        == 0.806202009349328


def test_price_takes_accrued_interest_from_the_rate() -> None:
    # A/E * rate * 100 / frequency with coupons to come, not coupon * A/E.
    assert _bond("PRICE", [45098.0, 45403.0, 0.13431405895368478, 0.18167305658967808, 100.0, 2.0, 2.0]) \
        == 96.42924034785027
    assert _bond("PRICE", [40129.0, 48923.0, 0.015269366925158684, 0.08153059843154511, 98.78907894617276, 2.0,
                           0.0]) == 30.415592647768403
    # With one coupon left it is coupon * (A/E).
    assert _bond("PRICE", [44139.0, 44198.0, 0.0858656433341273, 0.015639045918274874, 98.14162125781019, 1.0,
                           3.0]) == 99.24148797292862


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


#: Functions whose call here is not answered as Excel answers it, and why.
CELLS_GAPS: dict[str, str] = {"AMORDEGRC": _AMORDEGRC}


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
