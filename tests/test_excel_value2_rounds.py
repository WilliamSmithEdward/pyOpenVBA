"""Range.Value2 replayed from held-out rounds against live Excel (scripts/measure_value2_rounds.py,
tests/fixtures/value2_rounds.json): numbers, Currency, dates, logicals, errors, Empty, text Excel types as numbers, dates or
times, text it keeps, formulas, and formulas naming another workbook, open, closed and missing, written to one cell or
several, typed and late, and read back cell by cell.

Every round is kept, the first look too, each predicted before Excel answered.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location("measure_value2_rounds", ROOT / "scripts/measure_value2_rounds.py")
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
RECORD: dict[str, Any] = json.loads((ROOT / "tests/fixtures/value2_rounds.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("seed", sorted(RECORD["rounds"]))
def test_each_case_as_excel_answered(seed: str, tmp_path: Path) -> None:
    PROBE.books(tmp_path)
    wrong = [(case["target"], case["value"], case["excel"][:120], answer[:120])
             for case in RECORD["rounds"][seed] if (answer := str(PROBE.model(case, tmp_path))) != case["excel"]]
    assert wrong == []


def test_the_last_two_rounds_were_clean_as_run() -> None:
    last = sorted(RECORD["scores"])[-2:]
    assert [(RECORD["scores"][seed]["kind"], RECORD["scores"][seed]["differ_when_run"]) for seed in last] == [("counted", 0)] * 2
