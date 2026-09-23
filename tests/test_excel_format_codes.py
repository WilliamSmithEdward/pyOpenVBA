"""Number format codes as NumberFormat spells them and as the file does, replayed against the in-memory model.

tests/fixtures/format_codes/ is what scripts/measure_format_codes.py saw
in live Excel: each built-in and custom code set on a cell of its own and
read back through NumberFormat, the workbook Excel saved with them, and
the same workbook with codes planted in its stylesheet the way another
program might spell them, with what NumberFormat read of each.
"""

from __future__ import annotations

import html
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._number_format import from_file, normalized, to_file

FIXTURES = Path(__file__).parent / "fixtures" / "format_codes"
RECORD: dict[str, Any] = json.loads((FIXTURES / "format_codes.json").read_text(encoding="utf-8"))
CODES: list[dict[str, str]] = RECORD["codes"]


def _vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


@pytest.fixture(scope="module")
def model_book(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    """The probe's cells set in the model, what NumberFormat read back, and the model's save."""
    lines = ["Public Function Build() As String", "Dim out As String", "On Error Resume Next"]
    for row, entry in enumerate(CODES, start=1):
        lines += ["Err.Clear", f"Cells({row}, 1).Value = 1", f"Cells({row}, 1).NumberFormat = {_vba_text(entry['code'])}",
                  f'If Err.Number <> 0 Then out = out & "S" & Err.Number & "^" Else out = out & '
                  f'Cells({row}, 1).NumberFormat & "^"']
    lines += ["On Error GoTo 0", "Build = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(lines) + "\n", name="Probe")
    read = str(app.run("Build")).split("^")[: len(CODES)]
    path = tmp_path_factory.mktemp("format_codes") / "format_codes.xlsx"
    app.save(path)
    return read, path


def _written(path: Path) -> list[tuple[int, str]]:
    """Each row's cell in column A: the numFmtId of its xf, and the code numFmts spells out for it, if any."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    codes = {int(one): html.unescape(code)
             for one, code in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles)}
    block = re.search(r"<cellXfs.*?</cellXfs>", styles, re.DOTALL)
    assert block is not None
    xfs = re.findall(r"<xf\b[^>]*>", block.group(0))
    cells = dict(re.findall(r'<c r="A(\d+)"(?: s="(\d+)")?', sheet))
    out: list[tuple[int, str]] = []
    for row in range(1, len(CODES) + 1):
        found = re.search(r'numFmtId="(\d+)"', xfs[int(cells.get(str(row)) or 0)])
        number = int(found.group(1)) if found else 0
        out.append((number, codes.get(number, "")))
    return out


@pytest.mark.parametrize("index", range(len(CODES)), ids=[entry["code"] for entry in CODES])
def test_number_format_reads_back_as_excel_keeps_it(model_book: tuple[list[str], Path], index: int) -> None:
    assert model_book[0][index] == CODES[index]["read"]


@pytest.fixture(scope="module")
def excel_written() -> list[tuple[int, str]]:
    return _written(FIXTURES / "format_codes.xlsx")


@pytest.mark.parametrize("index", range(len(CODES)), ids=[entry["code"] for entry in CODES])
def test_a_format_is_written_with_excels_id_and_spelling(model_book: tuple[list[str], Path],
                                                         excel_written: list[tuple[int, str]], index: int) -> None:
    assert _written(model_book[1])[index] == excel_written[index]


@pytest.fixture(scope="module")
def excel_file_read() -> list[str]:
    """NumberFormat of every cell of a workbook Excel saved, and of every planted one, as the model reads them."""
    count = len(CODES) + len(RECORD["planted"]) + len(RECORD["planted_ids"])
    app = ExcelApplication.open(FIXTURES / "planted_source.xlsx", with_vba=False)
    app.add_module("Public Function Read() As String\nDim out As String, i As Long\n"
                   f"For i = 1 To {count}\nout = out & Cells(i, 1).NumberFormat & \"^\"\nNext\n"
                   "Read = out\nEnd Function\n", name="Reader")
    return str(app.run("Read")).split("^")[:count]


@pytest.mark.parametrize("index", range(len(CODES)), ids=[entry["code"] for entry in CODES])
def test_an_excel_written_format_reads_back(excel_file_read: list[str], index: int) -> None:
    assert excel_file_read[index] == CODES[index]["read"]


@pytest.mark.parametrize("index", range(len(RECORD["planted"])), ids=[entry["code"] for entry in RECORD["planted"]])
def test_a_planted_spelling_reads_as_excel_reads_it(excel_file_read: list[str], index: int) -> None:
    assert excel_file_read[len(CODES) + index] == RECORD["planted"][index]["read"]


@pytest.mark.parametrize("index", range(len(RECORD["planted_ids"])),
                         ids=[str(entry["id"]) for entry in RECORD["planted_ids"]])
def test_a_bare_builtin_id_reads_as_excel_knows_it(excel_file_read: list[str], index: int) -> None:
    offset = len(CODES) + len(RECORD["planted"])
    assert excel_file_read[offset + index] == RECORD["planted_ids"][index]["read"]


def test_the_two_spellings_go_back_and_forth() -> None:
    for entry in CODES:
        if entry["read"].startswith("S"):
            continue
        assert from_file(to_file(entry["read"])) == entry["read"]
        assert normalized(entry["read"]) == entry["read"]
