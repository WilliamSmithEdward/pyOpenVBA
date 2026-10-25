"""Optional parameters of every declared kind, replayed from native VBA (scripts/measure_optional_types.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures/vba_semantics/optional_types.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("case", RECORD["probes"], ids=lambda case: case["name"])
def test_an_optional_parameter_compiles_and_binds_as_native_vba(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    if case.get("compile_error"):
        message = case["compiler_message"].rsplit("\n", 1)[-1]
        with pytest.raises(VBACompileError, match=message):
            app.add_module(case["source"], name="Harness")
            app.run("Harness.Probe")
    else:
        app.add_module(case["source"], name="Harness")
        assert app.run("Harness.Probe") == case["result"]
