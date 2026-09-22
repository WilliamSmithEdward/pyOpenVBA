"""Cell formatting as Excel's object model answers it, replayed against the in-memory model.

tests/fixtures/range_format.json is written by scripts/measure_range_format.py
in live Excel: every record is the probe module Excel ran and what it
reported. tests/fixtures/format/ holds a workbook Excel formatted, saved and
reopened (scripts/measure_format_file.py), with what Excel read back from
it, and scripts/measure_border_storage.py's record of where Excel stores
each border it is asked to set.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._styles import parse_border

FIXTURES = Path(__file__).parent / "fixtures"
RECORDS: list[dict[str, str]] = json.loads((FIXTURES / "range_format.json").read_text())
ANSWERS: dict[str, Any] = json.loads((FIXTURES / "format" / "formats_answers.json").read_text())
STORAGE: dict[str, list[dict[str, Any]]] = json.loads((FIXTURES / "format" / "border_storage.json").read_text())
CASES: list[str] = ANSWERS["cases"]
ROWS = list(enumerate(CASES, start=1))


@pytest.mark.parametrize("record", RECORDS, ids=[record["name"] for record in RECORDS])
def test_measured_format(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record["code"], name="Probe")
    assert app.run("Report") == record["reported"]


def _described(app: ExcelApplication, row: int) -> dict[str, str]:
    reads: list[str] = ANSWERS["reads"]
    got = str(app.run("ReadCase", row)).split(";")[:-1]
    return dict(zip(reads, got, strict=True))


def _reader(app: ExcelApplication) -> None:
    app.add_module(ANSWERS["describe"] + 'Public Function ReadCase(row As Long) As String\n'
                   'ReadCase = Describe(ActiveSheet.Cells(row, 1))\nEnd Function\n', name="Reader")


@pytest.fixture(scope="module")
def excel_authored() -> ExcelApplication:
    app = ExcelApplication.open(FIXTURES / "format" / "formats.xlsx", with_vba=False)
    _reader(app)
    return app


@pytest.mark.parametrize("row,name", ROWS, ids=CASES)
def test_an_excel_authored_format_reads_back(excel_authored: ExcelApplication, row: int, name: str) -> None:
    assert _described(excel_authored, row) == ANSWERS["answers"][name]


@pytest.fixture(scope="module")
def model_authored(tmp_path_factory: pytest.TempPathFactory) -> ExcelApplication:
    """The same cases set by the model, saved, and opened again."""
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim c As Object"]
    for row, name in enumerate(ANSWERS["cases"], start=1):
        lines.append(f'Set c = ActiveSheet.Range("A{row}")')
        lines.append(f'ActiveSheet.Range("B{row}").Value = "{name}"')
        lines += [line.strip() for line in str(ANSWERS["setups"][name]).splitlines()]
    lines.append("End Sub")
    app.add_module("\n".join(lines) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path_factory.mktemp("formats") / "formats.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    _reader(reopened)
    return reopened


@pytest.mark.parametrize("row,name", ROWS, ids=CASES)
def test_a_model_authored_format_survives_a_save(model_authored: ExcelApplication, row: int, name: str) -> None:
    assert _described(model_authored, row) == ANSWERS["answers"][name]


@pytest.mark.parametrize("record", STORAGE["cases"], ids=[record["name"] for record in STORAGE["cases"]])
def test_a_border_is_stored_where_excel_stores_it(tmp_path: Path, record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    cells: dict[str, dict[str, str]] = record["cells"]  # type: ignore[assignment]
    reads = ['Public Function Report() As String', 'Dim e As Variant']
    for cell in cells:
        reads += [f'Report = Report & "{cell}="', 'For Each e In Array(5, 6, 7, 8, 9, 10)',
                  f'Report = Report & Range("{cell}").Borders(e).LineStyle & ":" & '
                  f'Range("{cell}").Borders(e).Weight & ","', 'Next e', 'Report = Report & "|"']
    reads.append("End Function")
    app.add_module("Public Sub Build()\n" + str(record["setup"]) + "\nEnd Sub\n" + "\n".join(reads) + "\n",
                   name="Probe")
    app.run("Build")
    reported = dict(item.split("=", 1) for item in str(app.run("Report")).split("|") if item)
    assert reported == {cell: found["reads"] for cell, found in cells.items()}
    path = tmp_path / "borders.xlsx"
    app.save(path)
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", re.search(r"<cellXfs.*?</cellXfs>", styles, re.S).group(0))  # type: ignore[union-attr]
    borders = re.findall(r"<border\b[^>]*>.*?</border>", re.search(r"<borders.*?</borders>", styles, re.S).group(0))  # type: ignore[union-attr]
    style_of = {ref: int(index) for ref, index in re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)}
    for cell, found in cells.items():
        border_id = int(re.search(r'borderId="(\d+)"', xfs[style_of.get(cell, 0)]).group(1))  # type: ignore[union-attr]
        assert parse_border(borders[border_id]) == parse_border(found["stored"]), cell


@pytest.mark.parametrize("item", STORAGE["patterns"], ids=[str(item["pattern"]) for item in STORAGE["patterns"]])
def test_each_fill_pattern_writes_excels_pattern_type(tmp_path: Path, item: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Function Report() As Long\nRange("B2").Interior.Pattern = {item["pattern"]}\n'
                   'Report = Range("B2").Interior.Pattern\nEnd Function\n', name="Probe")
    assert str(app.run("Report")) == item["reads"]
    path = tmp_path / "pattern.xlsx"
    app.save(path)
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    wanted = re.search(r'patternType="([^"]*)"', str(item["stored"])).group(1)  # type: ignore[union-attr]
    fills = re.findall(r"<fill>.*?</fill>", styles, re.S)
    assert any(f'patternType="{wanted}"' in fill for fill in fills)
