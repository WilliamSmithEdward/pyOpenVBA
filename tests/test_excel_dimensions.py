"""Row heights, column widths and hidden rows and columns, replayed against the in-memory model.

tests/fixtures/dimensions.json is written by scripts/measure_dimensions.py in
live Excel on a 96-DPI display: every record is the probe module Excel ran
and what it reported. tests/fixtures/dimensions/ holds what
scripts/measure_dimension_file.py saw of the file side: a workbook Excel
sized case by case and saved, with what it reads back and the XML it wrote;
a workbook with sizes planted the way another program might write them;
and the model's own new-workbook template as Excel writes it back.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
RECORDS: list[dict[str, str]] = json.loads((FIXTURES / "dimensions.json").read_text(encoding="utf-8"))
FILE: dict[str, Any] = json.loads((FIXTURES / "dimensions" / "dimensions_answers.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = FILE["cases"]
#: The cases the model can write as Excel does, with where each sits among
#: the sheets: a recorded_ case rests on heights only Excel works out, which
#: the model reads back from Excel's file but cannot write itself.
WRITTEN = [(index, case) for index, case in enumerate(CASES) if not case["name"].startswith("recorded_")]


#: Probes whose answer rests on something the model has not measured -- a
#: mix of fonts, a font outside the tables, super- or subscript, wrapped or
#: turned text. Excel's answers are recorded; the model says it cannot tell.
UNMEASURED = [record for record in RECORDS if record["name"].startswith("unmeasured_")]


@pytest.mark.parametrize("record", [record for record in RECORDS if record not in UNMEASURED],
                         ids=[record["name"] for record in RECORDS if record not in UNMEASURED])
def test_measured_dimension(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record["code"], name="Probe")
    assert app.run("Report") == record["reported"]


def test_the_baked_font_table_is_the_measurement() -> None:
    """src/pyopenvba/apps/excel/_font_rows.py holds exactly what tests/fixtures/font_rows.json measured."""
    from pyopenvba.apps.excel import _font_rows

    tables: list[dict[str, Any]] = json.loads((FIXTURES / "font_rows.json").read_text(encoding="utf-8"))
    for table in tables:
        index = _font_rows.table_of(table["font"].upper(), table["bold"], table["italic"])
        assert index is not None, table["font"]
        measured = list(zip(table["rows"], table["descent"], strict=True))
        baked = [_font_rows.row_of(index, pixels) for pixels in range(1, _font_rows.MOST_PIXELS + 1)]
        assert baked == measured, (table["font"], table["bold"], table["italic"])


@pytest.mark.parametrize("record", UNMEASURED, ids=[record["name"] for record in UNMEASURED])
def test_an_unmeasured_row_height_reports_itself(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record["code"], name="Probe")
    with pytest.raises(VBAUnsupportedError, match="not measured"):
        app.run("Report")


def _sheet_parts(path: Path) -> list[str]:
    """Each worksheet's XML, in tab order."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets: dict[str, str] = {}
        for element in re.findall(r"<Relationship\b[^>]*/>", rels):
            found = dict(re.findall(r'(\w+)="([^"]*)"', element))
            targets[found["Id"]] = found["Target"]
        out: list[str] = []
        for rid in re.findall(r'<sheet\b[^>]*r:id="([^"]+)"', workbook):
            target = targets[rid]
            part = target.lstrip("/") if target.startswith("/xl/") else f"xl/{target}"
            out.append(package.read(part).decode("utf-8"))
        return out


def _sizes(xml: str) -> dict[str, object]:
    head = re.search(r"<sheetFormatPr\b[^>]*/>", xml)
    cols = re.search(r"<cols>.*?</cols>", xml, re.DOTALL)
    dimension = re.search(r"<dimension\b[^>]*/>", xml)
    return {"sheetFormatPr": head.group(0) if head else "", "cols": cols.group(0) if cols else "",
            "rows": re.findall(r"<row\b[^>]*>", xml), "dimension": dimension.group(0) if dimension else ""}


def _reader(app: ExcelApplication) -> None:
    app.add_module(FILE["describe"] + "Public Function ReadAll() As String\nDim ws As Object, out As String\n"
                   'For Each ws In ActiveWorkbook.Worksheets\nout = out & Describe(ws) & "|"\nNext ws\n'
                   "ReadAll = out\nEnd Function\n", name="Reader")


def _answers(app: ExcelApplication) -> dict[str, str]:
    blocks = [block for block in str(app.run("ReadAll")).split("|") if block]
    return dict(zip([case["name"] for case in CASES], blocks, strict=True))


#: One case Excel reads differently once reopened, and the model does not
#: follow: with every column hidden, Excel takes the sheet's used block
#: from the file's dimension (C1, left by the column hidden while it kept
#: a width) and reads the columns before the last, open-ended <col> as
#: having widths of their own. The model reads what the sheet holds.
#: These are the fields apart: UsedRange and those columns'
#: UseStandardWidth; everything else must agree.
_REOPENED_APART = {"hide_every_column": {2: ("$C$1", "$A$1"), 15: ("0:0:True:False", "0:0:True:True"),
                                         16: ("0:0:True:False", "0:0:True:True"),
                                         17: ("0:0:True:False", "0:0:True:True")}}


def _agrees(name: str, got: str, want: str) -> bool:
    apart = _REOPENED_APART.get(name)
    if apart is None:
        return got == want
    got_fields, want_fields = got.split(";"), want.split(";")
    for index, (excel, model) in apart.items():
        if want_fields[index] != excel or got_fields[index] != model:
            return False
        got_fields[index] = want_fields[index]
    return got_fields == want_fields


@pytest.fixture(scope="module")
def excel_authored() -> dict[str, str]:
    app = ExcelApplication.open(FIXTURES / "dimensions" / "dimensions.xlsx", with_vba=False)
    _reader(app)
    return _answers(app)


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_an_excel_authored_size_reads_back(excel_authored: dict[str, str], case: dict[str, Any]) -> None:
    assert _agrees(case["name"], excel_authored[case["name"]], case["answers"])


@pytest.fixture(scope="module")
def model_authored(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """The same cases set by the model, one sheet each, and saved."""
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim ws As Object"]
    for index, case in enumerate(CASES):
        lines.append("Set ws = ActiveWorkbook.Worksheets(1)" if index == 0 else
                     "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets"
                     "(ActiveWorkbook.Worksheets.Count))")
        lines.append(f'ws.Name = "{case["name"]}"')
        lines += case["setup"].splitlines()
    lines += ["ActiveWorkbook.Worksheets(1).Activate", "End Sub"]
    app.add_module("\n".join(lines) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path_factory.mktemp("dimensions") / "dimensions.xlsx"
    app.save(path)
    return path


#: After inserting a row Excel marks rows customFormat="1" until something
#: reads the used range; insert_rows_read is the same case with that read
#: before the save, which is what the model writes.
_AS_READ = {"insert_rows": "insert_rows_read"}
_BY_NAME = {case["name"]: case for case in CASES}


@pytest.mark.parametrize("index,case", WRITTEN, ids=[case["name"] for _, case in WRITTEN])
def test_a_model_authored_size_is_written_as_excel_writes_it(model_authored: Path, index: int,
                                                             case: dict[str, Any]) -> None:
    expected = _BY_NAME[_AS_READ.get(case["name"], case["name"])]["sizes"]
    assert _sizes(_sheet_parts(model_authored)[index]) == expected


def test_an_insert_leaves_custom_format_on_rows_until_the_used_range_is_read() -> None:
    """The one thing the model does not write: a flag that says nothing, and that Excel drops on its own."""
    unread, read = _BY_NAME["insert_rows"]["sizes"], _BY_NAME["insert_rows_read"]["sizes"]
    assert [row.replace(' customFormat="1"', "") for row in unread["rows"]] == read["rows"]
    assert all('customFormat="1"' in row for row in unread["rows"])
    assert _BY_NAME["insert_rows"]["answers"] == _BY_NAME["insert_rows_read"]["answers"]


@pytest.fixture(scope="module")
def model_reopened(model_authored: Path) -> dict[str, str]:
    app = ExcelApplication.open(model_authored, with_vba=False)
    _reader(app)
    return _answers(app)


@pytest.mark.parametrize("case", [case for _, case in WRITTEN], ids=[case["name"] for _, case in WRITTEN])
def test_a_model_authored_size_survives_a_save(model_reopened: dict[str, str], case: dict[str, Any]) -> None:
    assert _agrees(case["name"], model_reopened[case["name"]], case["answers"])


def test_sizes_another_program_wrote_read_as_excel_reads_them() -> None:
    foreign = FILE["foreign"]
    app = ExcelApplication.open(FIXTURES / "dimensions" / "foreign.xlsx", with_vba=False)
    app.add_module("Public Function Read() As String\nDim ws As Object, out As String, r As Long, c As Long\n"
                   "Set ws = ActiveSheet\n"
                   f"For r = 1 To {foreign['read_rows']}\n"
                   'out = out & ws.Rows(r).RowHeight & ":" & ws.Rows(r).Height & ":" & ws.Rows(r).Hidden & ":" '
                   '& ws.Rows(r).UseStandardHeight & ";"\nNext r\n'
                   f"For c = 1 To {foreign['read_columns']}\n"
                   'out = out & ws.Columns(c).ColumnWidth & ":" & ws.Columns(c).Width & ":" & ws.Columns(c).Hidden '
                   '& ":" & ws.Columns(c).UseStandardWidth & ";"\nNext c\nRead = out\nEnd Function\n', name="Reader")
    assert app.run("Read") == foreign["answers"]


def test_a_workbook_in_calibri_is_sized_as_excel_sizes_it(tmp_path: Path) -> None:
    """Calibri 11, the Normal font before Aptos, measures the same on this display."""
    calibri = FILE["calibri"]
    app = ExcelApplication.open(FIXTURES / "dimensions" / "calibri.xlsx", with_vba=False)
    app.add_module("Public Sub Edit()\nDim ws As Object\nSet ws = ActiveSheet\n" + calibri["edit"] + "\nEnd Sub\n",
                   name="Editor")
    app.run("Edit")
    path = tmp_path / "calibri.xlsx"
    app.save(path)
    assert _sizes(_sheet_parts(path)[0]) == calibri["sizes"]
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module(FILE["describe"] + "Public Function Read() As String\nRead = Describe(ActiveSheet)\n"
                        "End Function\n", name="Reader")
    assert reopened.run("Read") == calibri["answers"]


def test_a_new_workbook_is_written_as_excel_on_this_display_writes_it(tmp_path: Path) -> None:
    """The template came from a 144-DPI display; Excel at 96 DPI writes 15pt rows and a 0.25 descent."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\nActiveSheet.Range("A1").Value = 1\nActiveSheet.Rows(3).RowHeight = 20\nEnd Sub\n',
                   name="Go")
    app.run("Go")
    path = tmp_path / "new.xlsx"
    app.save(path)
    assert _sizes(_sheet_parts(path)[0]) == FILE["template"]
