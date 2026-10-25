"""AutoFilter over text beyond ASCII, replayed against the in-memory model.

tests/fixtures/filter_text.json is what scripts/measure_filter_text.py saw
in live Excel. An = with no wildcard, and each value of a list, keeps the
text Windows' word sort cannot tell from it, case aside: =aesop keeps
the spelling with the ash and =strasse the one with a sharp s, while
=resume leaves out the accented spelling and =co-op the one with a
Unicode hyphen. > and < order text by word sort. A wildcard pattern
matches one character at a time, each by its word-sort key: ae* leaves
out the ash's spelling, *ss* the sharp s's, and k* keeps the Kelvin
sign's, while a dotless i is no I. A saved list keeps its values in word
sort's order.

The record keeps each probe's module and its answer as it came back: a 1
for each value's row shown and a 0 for one hidden, after each filter.
"""

from __future__ import annotations

import html
import json
import re
import zipfile
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "filter_text.json")
                                    .read_text(encoding="utf-8"))


def _run(code: str, macro: str, *args: object) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(code, name="Probe")
    return app.run(macro, *args)


def test_each_criterion_and_list_keeps_the_text_excel_keeps() -> None:
    """51 criteria, each =, <>, a wildcard pattern, > or <, and 6 lists, over 36 spellings."""
    assert _run(RECORD["code"], "Measure") == RECORD["raw"]


def test_a_wildcard_matches_each_letter_as_excel_does_and_a_save_lists_values_in_word_sort_order(
        tmp_path: Path) -> None:
    """Each of 13 pairs of letters, each looked for by a pattern: one letter for the other where Excel's
    wildcards take them for one; then a list the saved sheet keeps in word sort's order."""
    saved = tmp_path / "filter_text.xlsx"
    assert _run(RECORD["letters_code"], "Letters", str(saved)) == RECORD["letters_raw"]
    with zipfile.ZipFile(saved) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert [html.unescape(value) for value in re.findall(r'<filter val="([^"]*)"', sheet)] == RECORD["saved_values"]
