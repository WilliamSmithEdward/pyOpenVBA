"""When the Err object clears, replayed against the interpreter.

tests/fixtures/vba_semantics/err_lifetime.json is what
scripts/measure_err_lifetime.py saw in live Excel: small programs that
raise, handle and call, and Err.Number at the points that matter. Err
clears when a procedure of the macro starts, on every On Error statement,
on Resume and on leaving a procedure from its error handler; it keeps
what a procedure left in it after the procedure returns, and a call to a
built-in function or an object's member leaves it alone.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "vba_semantics" / "err_lifetime.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("case", RECORD["cases"], ids=[case["name"] for case in RECORD["cases"]])
def test_err_holds_what_vba_holds(case: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = "\n".join([RECORD["helpers"], "Public Function Probe() As String", "Dim out As String", case["body"],
                        "Probe = out", "End Function"])
    app.add_module(source + "\n", name="Probe")
    assert str(app.run("Probe")) == case["answer"]
