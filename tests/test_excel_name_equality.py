"""Which sheet names and defined names Excel takes for one, replayed against the in-memory model.

tests/fixtures/name_equality.json is what scripts/measure_name_equality.py
saw in live Excel. Excel compares names two ways. Where it keeps names
apart -- a sheet's name against the others, a formula's sheet, and the
spelling Names.Add keeps -- two names are one when they are as long as
each other and Windows' word sort cannot tell them apart, case aside: a
Kelvin sign is a K, a final sigma a sigma, and ash-ss is ae with a sharp
s, while ae is not the ash. Where it finds a name -- Worksheets(...),
Names(...), Names.Add and a formula's names -- word sort alone decides,
so AESOP finds the sheet spelt with the ash, the first in tab order, and
aesop redefines the name spelt so, which keeps its spelling.

The record keeps each probe's module and its answer as it came back, each
name as its characters' codes in hex: "measure" names sheets and adds
names spelt apart by case, accents, ligatures and a sharp s; "rules" adds
pairs of spellings one candidate rule or another takes for one; "marks"
the pairs with a combining accent.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._names import same_name, same_sheet_name

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "name_equality.json")
                                    .read_text(encoding="utf-8"))
#: Each probe: its module, the macro to run and what Excel's run of it answered.
PROBES = {"measure": (RECORD["code"], "Measure", RECORD["raw"]),
          "rules": (RECORD["rules"]["code"], "Rules", RECORD["rules"]["raw"]),
          "marks": (RECORD["marks"]["code"], "Marks", RECORD["marks"]["raw"])}


@pytest.mark.parametrize("probe", list(PROBES))
def test_names_are_kept_apart_and_found_as_excel_does(probe: str) -> None:
    """Every sheet named or refused, every sheet and name looked up, every name added, the Names collection and
    every formula's value and spelling, as Excel answered them."""
    code, macro, answer = PROBES[probe]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(code, name="Probe")
    assert app.run(macro) == answer


@pytest.mark.parametrize("probe", ["rules", "marks"])
def test_each_pair_is_one_name_where_excel_takes_it_for_one(probe: str) -> None:
    """A pair's second sheet is refused where its names are one sheet's, and its two names are one where
    Names(...) finds one by both spellings."""
    for (first, second, label), seen in zip(RECORD[probe]["pairs"], RECORD[probe]["rules"], strict=True):
        assert same_sheet_name("s" + first, "s" + second) == (seen["sheet_second"] == 1004), label
        assert same_name("n" + first, "n" + second) == (seen["name_found_first"] == seen["name_found_second"]), label
