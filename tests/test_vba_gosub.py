"""Replay real Excel's same-procedure subroutine and computed-jump answers."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.exceptions import VBACompileError, VBARuntimeError
from pyopenvba.interpreter import Interpreter

MEASURED = json.loads((Path(__file__).parent / "fixtures/vba_semantics/gosub.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", MEASURED["probes"], ids=lambda probe: probe["name"])
def test_excel_gosub_conformance(probe: dict[str, str]) -> None:
    runtime = Interpreter()
    runtime.add_module(probe["source"])
    assert runtime.run("Probe") == probe["result"]


def test_return_cannot_use_callers_gosub() -> None:
    runtime = Interpreter()
    runtime.add_module("Sub Main()\nGoSub Worker\nExit Sub\nWorker:\nOther\nReturn\nEnd Sub\nSub Other()\nReturn\nEnd Sub")
    with pytest.raises(VBARuntimeError) as caught:
        runtime.run("Main")
    assert caught.value.number == 3


def test_gosub_unknown_label() -> None:
    runtime = Interpreter()
    runtime.add_module("Sub Main()\nGoSub MissingLabel\nEnd Sub")
    with pytest.raises(VBACompileError, match="MissingLabel"):
        runtime.run("Main")


def test_gosub_stack_exhaustion_is_vba_error() -> None:
    runtime = Interpreter()
    runtime.add_module("Sub Main()\nAgain:\nGoSub Again\nReturn\nEnd Sub")
    with pytest.raises(VBARuntimeError) as caught:
        runtime.run("Main")
    assert caught.value.number == 28
