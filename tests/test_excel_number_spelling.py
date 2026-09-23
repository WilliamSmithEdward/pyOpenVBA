"""How a stored number is spelt, replayed against the in-memory model.

tests/fixtures/number_spelling.json is what scripts/measure_number_spelling.py
saw in live Excel: numbers from 1E-25 to 1E+300 with one to fifteen
significant digits, each stored in A1 and read back through A1's Formula,
through ``=A1&""`` in B1, and through A1's Text under General in a column
wide enough not to cut it short. The model runs the same macro.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.formula._values import number_text

RECORD: list[dict[str, str]] = json.loads(
    (Path(__file__).parent / "fixtures" / "number_spelling.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def spelt() -> list[tuple[str, ...]]:
    lines = ["Public Function Probe() As String", "Dim out As String", 'Range("B1").Formula = "=A1&"""""',
             "Columns(1).ColumnWidth = 60"]
    for entry in RECORD:
        lines += [f'Range("A1").Value = CDbl("{entry["number"]}")',
                  'out = out & Range("A1").Formula & "~" & Range("B1").Value & "~" & Range("A1").Text & "|"']
    lines += ["Probe = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(lines) + "\n", name="Probe")
    return [tuple(answer.split("~")) for answer in str(app.run("Probe")).split("|")[: len(RECORD)]]


@pytest.mark.parametrize("index", range(len(RECORD)), ids=[entry["number"] for entry in RECORD])
def test_a_number_is_spelt_as_excel_spells_it(spelt: list[tuple[str, ...]], index: int) -> None:
    entry = RECORD[index]
    assert spelt[index] == (entry["formula"], entry["text"], entry["general"])


def test_the_two_spellings_are_one_rule_with_two_widths() -> None:
    """Twenty characters for text and twenty-one for a Formula, the sign apart."""
    assert number_text(1e20) == "1E+20"
    assert number_text(1e20, formula=True) == "100000000000000000000"
    assert number_text(-1.23456789012345e100) == "-1.2345678901235E+100"
    assert number_text(-1.23456789012345e100, formula=True) == "-1.23456789012345E+100"
