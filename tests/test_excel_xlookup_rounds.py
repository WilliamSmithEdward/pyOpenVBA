"""WorksheetFunction.XLookup and XMatch replayed from held-out rounds against live Excel (scripts/measure_xlookup_rounds.py,
tests/fixtures/xlookup_rounds.json): random keys and results, looked up by numbers, text, wildcards, blanks, logicals and
arrays, in ranges and VBA arrays, down and across, every match and search mode and bad ones, typed and late.

Every round is kept, the failed ones too, each predicted before Excel answered. A case Excel never answered is left out:
those ended a call on an argument left empty, which Excel's VBA refuses to compile. The model refuses one kind of case,
XLookup over an array whose items each find several cells, where Excel answers the first cell or #VALUE! by what was
calculated before.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location("measure_xlookup_rounds", ROOT / "scripts/measure_xlookup_rounds.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
RECORD: dict[str, Any] = json.loads((ROOT / "tests/fixtures/xlookup_rounds.json").read_text(encoding="utf-8"))
REFUSAL = "REFUSED WorksheetFunction.XLookup over an array whose items each find several cells is not implemented"


def model_answer(case: dict[str, Any]) -> str:
    return str(PROBE.model(case))


@pytest.mark.parametrize("seed", sorted(RECORD["rounds"]))
def test_each_case_as_excel_answered(seed: str) -> None:
    wrong: list[tuple[Any, ...]] = []
    refused: list[dict[str, Any]] = []
    for case in RECORD["rounds"][seed]:
        answer = model_answer(case)
        if answer == REFUSAL:
            refused.append(case)
        elif answer != case["excel"]:
            wrong.append((case["function"], case["value"], case["lookup"], case["keys"], case["excel"][:120], answer[:120]))
    assert wrong == []
    assert all(case["function"] == "XLookup" and case["result"] == "two" for case in refused)


def test_the_last_two_rounds_were_clean_as_run() -> None:
    last = sorted(RECORD["scores"])[-2:]
    assert [(RECORD["scores"][seed]["kind"], RECORD["scores"][seed]["differ_when_run"]) for seed in last] == [("counted", 0)] * 2
