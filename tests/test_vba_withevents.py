"""Class event delivery held to isolated real Excel measurements."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.interpreter import Interpreter
from pyopenvba.exceptions import VBACompileError

MEASURED = json.loads((Path(__file__).parent / "fixtures/vba_semantics/withevents.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", MEASURED["probes"], ids=lambda probe: probe["name"])
def test_excel_class_events(probe: dict[str, str]) -> None:
    runtime = Interpreter()
    for name, source in MEASURED["modules"].items():
        runtime.add_module(source, name=name, kind="standard" if name == "EventLog" else "class")
    runtime.add_module(probe["source"], name="Probes")
    assert runtime.run("Probe") == probe["result"]


@pytest.mark.parametrize("source,kind", [
    ("Private WithEvents Source As EventSource", "standard"),
    ("Sub Main()\nDim WithEvents Source As EventSource\nEnd Sub", "class"),
    ("Private WithEvents Source() As EventSource", "class"),
    ("Private WithEvents Source As New EventSource", "class"),
    ("Private WithEvents Source As Object", "class"),
    ("Private WithEvents Source As Long", "class"),
])
def test_invalid_event_variable_declaration(source: str, kind: str) -> None:
    with pytest.raises(VBACompileError, match="WithEvents"):
        Interpreter().add_module(source, kind=kind)
