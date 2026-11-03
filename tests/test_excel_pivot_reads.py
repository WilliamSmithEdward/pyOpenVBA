"""Members a macro reads off a pivot table and its fields, filters, cells and cache (scripts/measure_pivot_reads.py)."""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/pivot_reads.json").read_text(encoding="utf-8"))
#: MemoryUsed has no rule to measure, so it is left reporting itself and not read here.
UNREAD = [index for index, read in enumerate(MEASURED["reads"]) if "MemoryUsed" in read]
SOURCE = re.sub(r"    Err\.Clear\n    v = \"\"\n    v = [^\n]*MemoryUsed[^\n]*\n[^\n]*\n[^\n]*\n", "", MEASURED["source"])


def _without_unread(answer: str) -> str:
    parts = answer.split(",")
    return ",".join(part for place, part in enumerate(parts) if place - 1 not in UNREAD)


def _case(index: int) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE.replace("Private Function Tried", "Public Function Tried"), name="Measure")
    return str(app.run("Tried", index))


@pytest.mark.parametrize("index", range(len(MEASURED["cases"])), ids=lambda index: str(index))
def test_pivot_reads(index: int) -> None:
    assert _case(index) == _without_unread(MEASURED["cases"][index][1])
