"""The in-memory Word against what real Word answered.

Every probe in ``fixtures/word_model/probes.txt`` ran in live Word on a
document emptied of what the probe before it left, and its answer is in
``measured.json``.  This replays each against the model and requires
the same answer.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.word import WordApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import NULL, VBAArray, VBAErrorValue, to_text, type_name

FIXTURES = Path(__file__).parent / "fixtures" / "word_model"


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
    # The measuring wrote each answer on its own line and stripped the
    # carriage return off the end, so a Word paragraph's trailing mark
    # is not in the file; the same is stripped here.
    return f"{type_name(value)}|{to_text(value).rstrip(chr(13))}"


def evaluate(probe: str) -> str:
    setup, expression = split_probe(probe)
    app = WordApplication()
    app.add_document()
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


PROBES = read_probes()
MEASURED = read_measured()


def test_every_probe_was_measured() -> None:
    missing = [probe for probe in PROBES if probe not in MEASURED]
    assert not missing, f"{len(missing)} probes have no measured answer, first {missing[0]!r}"


@pytest.mark.parametrize("probe", PROBES, ids=range(len(PROBES)))
def test_matches_word(probe: str) -> None:
    want = MEASURED[probe]
    got = evaluate(probe)
    assert got == want, f"{probe}\n  Word: {want}\n  ours: {got}"
