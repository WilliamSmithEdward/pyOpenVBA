"""Escaped characters in Range.NumberFormat, replayed against the in-memory model.

tests/fixtures/format_escapes.json is what scripts/measure_format_escapes.py
saw in live Excel: formats that escape each printable character after a
digit, before one, alone and twice, and a few words, as NumberFormat read
them back once set.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

CODES: dict[str, str] = json.loads(
    (Path(__file__).parent / "fixtures" / "format_escapes.json").read_text(encoding="utf-8"))["codes"]


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


@pytest.fixture(scope="module")
def read_back() -> dict[str, str]:
    """Every code set on a cell of the model and read back, as the probe did it in Excel."""
    codes = list(CODES)
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String, codes As Variant, i As Long",
             "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next"]
    for index in range(0, len(codes), 40):
        lines += ["codes = Array(" + ", ".join(_vba_text(code) for code in codes[index:index + 40]) + ")",
                  "For i = 0 To UBound(codes)", "    Err.Clear", '    ws.Range("A1").NumberFormat = "General"',
                  '    ws.Range("A1").NumberFormat = codes(i)',
                  '    If Err.Number <> 0 Then out = out & "E" & Err.Number & "|~|" Else '
                  'out = out & ws.Range("A1").NumberFormat & "|~|"', "Next"]
    lines += ["Probe = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(lines) + "\n", name="Probe")
    return dict(zip(codes, str(app.run("Probe")).split("|~|"), strict=False))


@pytest.mark.parametrize("code", list(CODES))
def test_number_format_keeps_an_escape_where_excel_does(code: str, read_back: dict[str, str]) -> None:
    assert read_back[code] == CODES[code]
