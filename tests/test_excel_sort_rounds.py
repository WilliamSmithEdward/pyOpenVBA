"""The worksheet Sort rows replayed from held-out rounds against live Excel (scripts/measure_sort_rounds.py,
tests/fixtures/sort_rounds.json): random cells, fields added by Add and Add2, the Sort's settings and a field's set to good
and bad values, typed and late, every setting read back with its type, and Apply once or twice with the cells after.

Every round is kept, the ones that found differences too; each was predicted before Excel answered. Two cases the model
refuses: Apply with a field on a data option past xlSortTextAsNumbers, where Excel can stop for a dialog of its own, and
which once did.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location("measure_sort_rounds", ROOT / "scripts/measure_sort_rounds.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
RECORD: dict[str, Any] = json.loads((ROOT / "tests/fixtures/sort_rounds.json").read_text(encoding="utf-8"))
REFUSAL = "REFUSED sorting with a DataOption other than xlSortNormal or xlSortTextAsNumbers is not implemented"


def model_answer(case: dict[str, Any]) -> str:
    return str(PROBE.model(case))


def _data_option_past_text_as_numbers(case: dict[str, Any]) -> bool:
    """Whether a field Apply sorts by ends on such a data option: one Add gave or a set left."""
    options = {n + 1: field["option"] for n, field in enumerate(case["fields"])}
    for at, name, value in case.get("sets", []):
        if name == "DataOption" and value in ("2", "1.5"):
            options[at] = value
    return case["apply"] and any(value in ("2", "1.5") for value in options.values())


@pytest.mark.parametrize("seed", sorted(RECORD["rounds"]))
def test_each_case_as_excel_answered(seed: str) -> None:
    wrong: list[tuple[Any, ...]] = []
    refused: list[tuple[dict[str, Any], str]] = []
    for case in RECORD["rounds"][seed]:
        answer = model_answer(case)
        if answer.startswith("REFUSED"):
            refused.append((case, answer))
        elif answer != case["excel"]:
            wrong.append((case["settings"], case["range"], case["fields"], case["excel"][:160], answer[:160]))
    assert wrong == []
    assert all(answer == REFUSAL and _data_option_past_text_as_numbers(case) for case, answer in refused)


def test_where_excel_stopped_for_a_dialog_the_model_refuses() -> None:
    stopped = [case for cases in RECORD["rounds"].values() for case in cases if case["excel"].startswith("no answer")]
    assert stopped and all(model_answer(case) == REFUSAL for case in stopped)


def test_the_last_two_rounds_were_clean_as_run() -> None:
    last = sorted(RECORD["scores"])[-2:]
    assert [(RECORD["scores"][seed]["kind"], RECORD["scores"][seed]["differ_when_run"]) for seed in last] == [("counted", 0)] * 2
