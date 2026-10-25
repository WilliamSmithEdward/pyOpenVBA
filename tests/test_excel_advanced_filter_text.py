"""AdvancedFilter over text beyond ASCII, replayed against the in-memory model.

tests/fixtures/advanced_filter_text.json is what
scripts/measure_advanced_filter_text.py saw in live Excel, the list and
the criteria of tests/fixtures/filter_text.json filtered in place through
a criteria range. AdvancedFilter matches text as AutoFilter does: an =
with no wildcard keeps the text Windows' word sort cannot tell from it,
case aside, > and < order text by word sort, and a wildcard pattern
matches one character at a time by each character's word-sort key. A
plain criterion is a beginning: it keeps text whose first characters, as
many as it has, word sort cannot tell from it, so aesop begins Aesop
spelt with the ash while ae does not.

The record keeps each probe's module and its answer: a 1 for each value's
row shown and a 0 for one hidden, after each criterion.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "advanced_filter_text.json")
                                    .read_text(encoding="utf-8"))
PROBES = {"criteria": ("code", "Measure", "raw"), "letters": ("letters_code", "Letters", "letters_raw"),
          "beginnings": ("beginnings_code", "Beginnings", "beginnings_raw")}


@pytest.mark.parametrize("probe", list(PROBES))
def test_each_criterion_keeps_the_text_excel_keeps(probe: str) -> None:
    """51 criteria over 36 spellings, 26 wildcard patterns over 13 pairs of letters, and 20 plain beginnings."""
    code, macro, answer = PROBES[probe]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD[code], name="Probe")
    assert app.run(macro) == RECORD[answer]
