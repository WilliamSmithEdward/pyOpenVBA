"""Calculation state and controls, replayed from live Excel (scripts/measure_calculation_state.py,
tests/fixtures/calculation_state.json): CalculationState, CalculationVersion, CalculationInterruptKey,
CalculateUntilAsyncQueriesDone, ForceFullCalculation, Range.Dirty and CalculateRowMajorOrder, and the files Excel
works out again as it opens them."""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "calculation_state.json").read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


def session(code: str) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(portable(code), name="Probe")
    return app


def crafted(made: Path, target: Path, calc: str | None) -> None:
    """A copy of ``made`` with ``calc`` for its calcPr, and the cached values 99 and 77 its formulas would not
    give, as scripts/measure_calculation_state.py writes them."""
    with zipfile.ZipFile(made) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as copy:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "xl/workbook.xml" and calc is not None:
                text = re.sub(r"<calcPr\b[^>]*?(?:/>|>.*?</calcPr>)", calc, data.decode("utf-8"), flags=re.DOTALL)
                if calc and "<calcPr" not in text:
                    text = text.replace("</workbook>", calc + "</workbook>")
                data = text.encode("utf-8")
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                text = re.sub(r'(<c r="A2"[^>]*>\s*<f>A1\*10</f>\s*<v>)20(</v>)', r"\g<1>99\g<2>", text)
                text = re.sub(r'(<c r="A3"[^>]*>\s*<f>A2\+1</f>\s*<v>)21(</v>)', r"\g<1>77\g<2>", text)
                data = text.encode("utf-8")
            copy.writestr(item, data)


def saved_calc(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        found = re.search(r"<calcPr\b[^>]*>", package.read("xl/workbook.xml").decode("utf-8"))
    return found.group(0) if found else ""


def test_calculation_state_follows_excel() -> None:
    assert str(session(RECORD["rules"]).run("Probe.Probe")).split("|") == RECORD["rules_answer"].split("|")


@pytest.mark.parametrize("name", list(RECORD["variants"]))
def test_a_file_is_worked_out_as_it_opens_as_excel_works_it_out(name: str, tmp_path: Path) -> None:
    app = session(RECORD["opens"])
    made = tmp_path / "made.xlsx"
    app.run("Probe.MakeFile", str(made))
    target, copy = tmp_path / f"{name}.xlsx", tmp_path / f"{name}_copy.xlsx"
    crafted(made, target, RECORD["variants"][name])
    answer = str(app.run("Probe.Opened", str(target), str(copy)))
    assert [answer, saved_calc(copy)] == RECORD["opens_answer"][name]


def test_force_full_calculation_holds_every_workbook_to_full_calculation(tmp_path: Path) -> None:
    answer = str(session(RECORD["forced"]).run("Probe.Probe", str(tmp_path)))
    assert answer.split("|") == RECORD["forced_answer"].split("|")
    assert saved_calc(tmp_path / "forced.xlsx") == RECORD["forced_file"]
    assert saved_calc(tmp_path / "unforced.xlsx") == RECORD["unforced_file"]
    # A file that says forceFullCalc, opened first, holds the next session to it.
    assert str(session(RECORD["first"]).run("Probe.Probe", str(tmp_path / "forced.xlsx"))) == RECORD["first_answer"]
