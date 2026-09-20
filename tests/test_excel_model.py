"""The in-memory Excel against what real Excel answered.

Every probe in ``fixtures/excel_model/probes.txt`` ran in live Excel on
a freshly added worksheet, and its answer is in ``measured.json``.  This
replays each one against the model and requires the same answer: the
type Excel gave the value and the string it turns into, or the error
number Excel raised.

The probes are the places a plausible object model goes wrong: what
Address spells, what type a cell gives back what was put into it, where
End lands, and what UsedRange covers.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import NULL, VBAArray, VBAErrorValue, to_text, type_name

FIXTURES = Path(__file__).parent / "fixtures" / "excel_model"

#: Probes measured against nothing stable.  Kept in the file because
#: they are worth re-reading after a re-measure, and skipped here
#: because the answer is the measuring session's, not Excel's.
UNCOMPARABLE: dict[str, str] = {
    ";; Len(Range(\"A1\").Font.Name) > 0": "the default font is the Excel version's, not a behaviour",
}


def split_probe(line: str) -> tuple[str, str]:
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
    app = ExcelApplication()
    app.add_workbook()
    source = "Function Probe() As Variant\n"
    if setup:
        source += f"    {setup}\n"
    source += f"    Probe = ({expression})\nEnd Function\n"
    try:
        app.add_module(source, name="Probes")
        return describe(app.interpreter.run("Probe"))
    except VBARuntimeError as failure:
        return f"!{failure.number}"
    except VBACompileError:
        return "!compile"
    except VBAUnsupportedError as gap:
        return f"!unsupported: {gap}"


PROBES = [probe for probe in read_probes() if probe not in UNCOMPARABLE]
MEASURED = read_measured()


def test_every_probe_was_measured() -> None:
    missing = [probe for probe in PROBES if probe not in MEASURED]
    assert not missing, f"{len(missing)} probes have no measured answer, first {missing[0]!r}"


@pytest.mark.parametrize("probe", PROBES, ids=range(len(PROBES)))
def test_matches_excel(probe: str) -> None:
    want = MEASURED[probe]
    got = evaluate(probe)
    assert got == want, f"{probe}\n  Excel: {want}\n  ours : {got}"
