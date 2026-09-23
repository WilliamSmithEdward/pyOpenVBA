"""Row, column and sheet-wide formats, replayed against the in-memory model.

tests/fixtures/row_formats/ is what scripts/measure_row_formats.py saw in
live Excel on a 96-DPI display: every case's setup and reads, what Excel
read before saving and after opening its file again, the workbook it
saved, and workbooks with row and column styles planted in the XML the
way another program might write them, as planted and as Excel saved them
again. The model must read the same, and write the same sheet XML, with
each style compared by what its xf says rather than by its index, since
Excel numbers xfs in the order they were made.
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

FIXTURES = Path(__file__).parent / "fixtures" / "row_formats"
RECORD: dict[str, Any] = json.loads((FIXTURES / "row_formats.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
PLANTED: list[dict[str, Any]] = RECORD["planted"]
#: Cases whose answer rests on a mix of fonts the model has not measured: it says so.
UNMEASURED = [case for case in CASES if case["name"].startswith("unmeasured_")]
MEASURED = [case for case in CASES if case not in UNMEASURED]


def _reads_function(name: str, reads: list[str]) -> str:
    """The same reads the probe made: the used range first, then each read, into one line."""
    lines = [f"Private Function {name}(ws As Object) As String", "Dim out As String, v As Variant",
             "On Error Resume Next"]
    for expression in ["ws.UsedRange.Address", *reads]:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return "\n".join([*lines, "On Error GoTo 0", f"{name} = out", "End Function"]) + "\n"


def _case_function(index: int, setup: str, reads: list[str]) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String", "Dim failed As String", "On Error Resume Next",
             "Err.Clear", *setup.splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
             "On Error GoTo 0", f"Case{index} = failed & Reads{index}(ws)", "End Function"]
    return "\n".join(lines) + "\n" + _reads_function(f"Reads{index}", reads)


def _replay(case: dict[str, Any]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + _case_function(0, case["setup"], case["reads"])
                   + "Public Function Report() As String\nReport = Case0(ActiveWorkbook.Worksheets(1))\n"
                     "End Function\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("case", MEASURED, ids=[case["name"] for case in MEASURED])
def test_measured_row_format(case: dict[str, Any]) -> None:
    assert _replay(case) == case["before"]


@pytest.mark.parametrize("case", UNMEASURED, ids=[case["name"] for case in UNMEASURED])
def test_an_unmeasured_mix_of_fonts_reports_itself(case: dict[str, Any]) -> None:
    with pytest.raises(VBAUnsupportedError, match="not measured"):
        _replay(case)


# --- the file side ----------------------------------------------------------------------------


def _parts(path: Path) -> tuple[list[str], str]:
    """Each worksheet's XML in tab order, and the stylesheet."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        relationships = (dict(re.findall(r'(\w+)="([^"]*)"', element))
                         for element in re.findall(r"<Relationship\b[^>]*/>", rels))
        targets = {found["Id"]: found["Target"] for found in relationships}
        sheets: list[str] = []
        for rid in re.findall(r'<sheet\b[^>]*r:id="([^"]+)"', workbook):
            target = targets[rid]
            sheets.append(package.read(target.lstrip("/") if target.startswith("/xl/") else f"xl/{target}")
                          .decode("utf-8"))
        return sheets, package.read("xl/styles.xml").decode("utf-8")


def _children(styles: str, section: str) -> list[str]:
    block = re.search(rf"<{section}\b[^>]*>(.*?)</{section}>", styles, re.DOTALL)
    if block is None:
        return []
    tag = {"fonts": "font", "fills": "fill", "borders": "border", "cellXfs": "xf"}[section]
    return re.findall(rf"<{tag}\b[^>]*/>|<{tag}\b[^>]*>.*?</{tag}>", block.group(1), re.DOTALL)


def _xfs(styles: str) -> list[str]:
    """Each cell xf spelled out: its number format, font, fill and border, its flags and its alignment."""
    formats = dict(re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles))
    fonts, fills, borders = (_children(styles, section) for section in ("fonts", "fills", "borders"))
    out: list[str] = []
    for xf in _children(styles, "cellXfs"):
        head = dict(re.findall(r'(\w+)="([^"]*)"', xf[: xf.index(">")]))
        number = head.pop("numFmtId", "0")
        parts = [f"numFmt={formats.get(number, number)}", fonts[int(head.pop("fontId", "0"))],
                 fills[int(head.pop("fillId", "0"))], borders[int(head.pop("borderId", "0"))]]
        parts += [f"{key}={value}" for key, value in sorted(head.items())]
        inner = xf[xf.index(">") + 1: xf.rindex("<")] if not xf.endswith("/>") else ""
        out.append(";".join([*parts, inner]))
    return out


def _described(sheet: str, xfs: list[str]) -> dict[str, object]:
    """The parts of a sheet that say how it is formatted, each style by what its xf says."""

    def styled(tag: str, name: str) -> str:
        return re.sub(rf'\b{name}="(\d+)"', lambda match: f'{name}="{xfs[int(match.group(1))]}"', tag)

    rows: list[tuple[str, list[str]]] = []
    for row in re.findall(r"<row\b[^>]*?(?:/>|>.*?</row>)", sheet, re.DOTALL):
        start = re.match(r"<row\b[^>]*?/?>", row)
        assert start is not None
        head = styled(start.group(0).rstrip("/>").rstrip(">"), "s")
        cells = [styled(" ".join(re.findall(r'\b(?:r|s)="[^"]*"', cell)), "s")
                 for cell in re.findall(r"<c\b[^>]*?/?>", row)]
        rows.append((head, cells))
    cols = re.search(r"<cols>.*?</cols>", sheet, re.DOTALL)
    head = re.search(r"<sheetFormatPr\b[^>]*/>", sheet)
    dimension = re.search(r"<dimension\b[^>]*/>", sheet)
    return {"dimension": dimension.group(0) if dimension else "", "sheetFormatPr": head.group(0) if head else "",
            "cols": styled(cols.group(0), "style") if cols else "", "rows": rows}


def _read_all(app: ExcelApplication, cases: list[dict[str, Any]], first: int = 1) -> list[str]:
    """Every case's reads, each on the sheet that holds it."""
    functions = "".join(_reads_function(f"Reads{index}", case["reads"]) for index, case in enumerate(cases))
    calls = "".join(f'out = out & Reads{index}(ActiveWorkbook.Worksheets({index + first})) & "|"\n'
                    for index in range(len(cases)))
    app.add_module(RECORD["helper"] + functions + "Public Function ReadAll() As String\nDim out As String\n"
                   + calls + "ReadAll = out\nEnd Function\n", name="Reader")
    return str(app.run("ReadAll")).split("|")[: len(cases)]


@pytest.fixture(scope="module")
def excel_authored() -> list[str]:
    return _read_all(ExcelApplication.open(FIXTURES / "row_formats.xlsx", with_vba=False), CASES)


@pytest.mark.parametrize("index,case", list(enumerate(MEASURED)), ids=[case["name"] for case in MEASURED])
def test_an_excel_authored_format_reads_back(excel_authored: list[str], index: int, case: dict[str, Any]) -> None:
    assert excel_authored[CASES.index(case)] == case["after"]


@pytest.fixture(scope="module")
def model_authored(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Every measured case set by the model on a sheet of its own, and saved."""
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim ws As Object, report As String"]
    for index, case in enumerate(MEASURED):
        lines.append("Set ws = ActiveWorkbook.Worksheets(1)" if index == 0 else
                     "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets"
                     "(ActiveWorkbook.Worksheets.Count))")
        lines += [f'ws.Name = "{case["name"]}"', f"report = Case{index}(ws)"]
    lines += ["ActiveWorkbook.Worksheets(1).Activate", "End Sub"]
    functions = "".join(_case_function(index, case["setup"], case["reads"]) for index, case in enumerate(MEASURED))
    app.add_module(RECORD["helper"] + functions + "\n".join(lines) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path_factory.mktemp("row_formats") / "row_formats.xlsx"
    app.save(path)
    return path


@pytest.fixture(scope="module")
def excel_sheets() -> dict[str, dict[str, object]]:
    sheets, styles = _parts(FIXTURES / "row_formats.xlsx")
    xfs = _xfs(styles)
    return {case["name"]: _described(sheet, xfs) for case, sheet in zip(CASES, sheets, strict=True)}


@pytest.mark.parametrize("index,case", list(enumerate(MEASURED)), ids=[case["name"] for case in MEASURED])
def test_a_model_authored_format_is_written_as_excel_writes_it(model_authored: Path, excel_sheets: dict[str, Any],
                                                               index: int, case: dict[str, Any]) -> None:
    sheets, styles = _parts(model_authored)
    assert _described(sheets[index], _xfs(styles)) == excel_sheets[case["name"]]


@pytest.fixture(scope="module")
def model_reopened(model_authored: Path) -> list[str]:
    return _read_all(ExcelApplication.open(model_authored, with_vba=False), MEASURED)


@pytest.mark.parametrize("index,case", list(enumerate(MEASURED)), ids=[case["name"] for case in MEASURED])
def test_a_model_authored_format_survives_a_save(model_reopened: list[str], index: int, case: dict[str, Any]) -> None:
    assert model_reopened[index] == case["after"]


# --- formats another program wrote -----------------------------------------------------------------


@pytest.fixture(scope="module")
def planted(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    """The planted workbook opened by the model, each macro run and each sheet read, then saved."""
    app = ExcelApplication.open(FIXTURES / "planted_source.xlsx", with_vba=False)
    macros = "".join(f"Set ws = ActiveWorkbook.Worksheets({index + 2})\n{item['macro']}\n"
                     for index, item in enumerate(PLANTED) if item["macro"])
    app.add_module("Public Sub Edit()\nDim ws As Object\n" + macros + "End Sub\n", name="Editor")
    app.run("Edit")
    answers = _read_all(app, PLANTED, first=2)
    path = tmp_path_factory.mktemp("planted") / "planted.xlsx"
    app.save(path)
    return answers, path


@pytest.mark.parametrize("index,item", list(enumerate(PLANTED)), ids=[item["name"] for item in PLANTED])
def test_a_planted_format_reads_as_excel_reads_it(planted: tuple[list[str], Path], index: int,
                                                  item: dict[str, Any]) -> None:
    assert planted[0][index] == item["answers"]


@pytest.mark.parametrize("index,item", list(enumerate(PLANTED)), ids=[item["name"] for item in PLANTED])
def test_a_planted_format_is_written_back_as_excel_writes_it(planted: tuple[list[str], Path], index: int,
                                                             item: dict[str, Any]) -> None:
    ours, our_styles = _parts(planted[1])
    theirs, their_styles = _parts(FIXTURES / "planted.xlsx")
    assert _described(ours[index + 1], _xfs(our_styles)) == _described(theirs[index + 1], _xfs(their_styles))
