"""Everyday Application members, replayed from live Excel (scripts/measure_application_members.py,
tests/fixtures/application_members.json): PathSeparator, Value, OperatingSystem, MeasurementUnit, the length
conversions to the bit, SheetsInNewWorkbook and what Workbooks.Add makes of it, and ThisCell beside Caller."""

from __future__ import annotations

import json
import struct
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "application_members.json").read_text(encoding="utf-8"))
CONVERT = '''Public Function Cm(x As Double) As Double
    Cm = Application.CentimetersToPoints(x)
End Function

Public Function Inches(x As Double) As Double
    Inches = Application.InchesToPoints(x)
End Function
'''


@pytest.fixture(scope="module")
def replayed() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")
    return str(app.run("Probe.Probe")).split("|")


@pytest.mark.parametrize("index", range(len(RECORD["answers"])),
                         ids=lambda index: RECORD["answers"][index].split("=")[0][:30])
def test_each_answer_is_excels(replayed: list[str], index: int) -> None:
    assert replayed[index] == RECORD["answers"][index]


def _bits(value: object) -> str:
    assert isinstance(value, float)
    return struct.pack(">d", value).hex().upper()


@pytest.mark.parametrize("value", list(RECORD["conversions"]))
def test_the_conversions_are_excels_to_the_bit(value: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(CONVERT, name="Convert")
    centimetres, inches = RECORD["conversions"][value].split(",")
    assert _bits(app.run("Convert.Cm", float(value))) == centimetres
    assert _bits(app.run("Convert.Inches", float(value))) == inches
