"""The interpreter against what real VBA answered.

Every probe in ``fixtures/vba_semantics/probes.txt`` was evaluated in
live Excel by ``scripts/measure_vba_semantics.py`` and its answer
recorded in ``measured.json``.  This test replays them through the
interpreter and requires the same answer, which is the type VBA gave
the value and the string it turns into, or the error number it raised.

Both halves of the answer matter.  An interpreter that says Long where
VBA says Integer looks right until the day a macro multiplies two of
them and VBA raises an overflow that the interpreter never would.

To re-measure after changing a probe, run the script on a Windows
machine with Excel installed.  Nothing here needs Office.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter import Interpreter
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import NULL, VBAArray, VBAErrorValue, to_text, type_name

FIXTURES = Path(__file__).parent / "fixtures" / "vba_semantics"


def split_probe(line: str) -> tuple[str, str]:
    """A probe is an expression, or setup and an expression around ``;;``."""
    setup, marker, expression = line.partition(";;")
    return (setup.strip(), expression.strip()) if marker else ("", line.strip())


def read_probes() -> list[str]:
    out: list[str] = []
    for line in (FIXTURES / "probes.txt").read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            out.append(stripped)
    return out


def read_measured() -> dict[str, str]:
    return json.loads((FIXTURES / "measured.json").read_text(encoding="utf-8"))["probes"]


def describe(value: object) -> str:
    """The same description the VBA side of the measurement writes."""
    if value is NULL:
        return "Null|Null"
    if isinstance(value, VBAObject):
        return f"{type_name(value)}|<object>"
    if isinstance(value, VBAArray):
        return f"{type_name(value)}|{value.bounds[0][0]}..{value.bounds[0][1]}"
    if isinstance(value, VBAErrorValue):
        return "Error|<error>"
    return f"{type_name(value)}|{to_text(value)}"


def evaluate(probe: str) -> str:
    setup, expression = split_probe(probe)
    source = "Function Probe() As Variant\n"
    if setup:
        source += f"    {setup}\n"
    source += f"    Probe = ({expression})\nEnd Function\n"
    vba = Interpreter()
    try:
        vba.add_module(source, name="Probes")
        return describe(vba.run("Probe"))
    except VBARuntimeError as failure:
        return f"!{failure.number}"
    except VBACompileError:
        return "!compile"
    except VBAUnsupportedError as gap:
        return f"!unsupported: {gap}"


PROBES = read_probes()
MEASURED = read_measured()


def test_every_probe_was_measured() -> None:
    missing = [probe for probe in PROBES if probe not in MEASURED]
    assert not missing, f"{len(missing)} probes have no measured answer, first {missing[0]!r}"


@pytest.mark.parametrize("probe", PROBES, ids=range(len(PROBES)))
def test_matches_excel(probe: str) -> None:
    want = MEASURED[probe]
    got = evaluate(probe)
    assert got == want, f"{probe}\n  Excel: {want}\n  ours : {got}"
