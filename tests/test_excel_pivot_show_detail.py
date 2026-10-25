"""A pivot table's items collapsed and expanded, ShowDetail, replayed against the in-memory model.

tests/fixtures/pivot_show_detail/ holds what
scripts/measure_pivot_show_detail.py saw in live Excel: a pivot table on
a sheet of its own for each layout, over one list of regions, products,
channels, amounts and units, with an item or a whole field collapsed or
expanded on the rows or the columns, in compact, outline and tabular
form, and the errors that raises; each table dumped as its cells with
their indents and the ShowDetail of each field and item; the workbook
Excel saved of them all; and each table as Excel read it again from that
file.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_pivot_layouts import column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_show_detail"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_show_detail.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: What Excel writes afresh each time: revision ids and cache numbering.
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(part).decode("utf-8").replace(" />", "/>"))
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            out[name.group(1) if name else part] = REVISIONS.sub("", text)
    return out


def recorded(layout: dict[str, Any]) -> str:
    return REVISIONS.sub("", layout["file"].replace(" />", "/>"))


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it."""
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{step}", str(index))
                   for index, layout in enumerate(LAYOUTS, 1))
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{source}", RECORD["source"])
            .replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep))
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    return app


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model, its dumps in the session and once the model opened its own file again, and
    the folder it saved to."""
    folder = tmp_path_factory.mktemp("pivot_show_detail")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_show_detail.xlsx") == {layout["name"]: recorded(layout) for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_details_show_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                         layout: dict[str, Any]) -> None:
    """Every cell of the table with its indent, and each field's and item's ShowDetail, with the errors setting it
    raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_details_read_back_from_the_file_the_model_saved(made: tuple[dict[str, str], dict[str, str], Path],
                                                           layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_detail_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                 layout: dict[str, Any]) -> None:
    """The table's definition whole: each item's sd, its subtotals' too, and the lines a collapsed item leaves."""
    assert tables(made[2] / "pivot_show_detail.xlsx")[layout["name"]] == recorded(layout)


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return sheet_cells(made[2] / "pivot_show_detail.xlsx", names), sheet_cells(FIXTURES / "pivot_show_detail.xlsx", names)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_show_detail.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's columns as wide as Excel fits them each time it lays the table out, bestFit beside them (see
    tests/test_excel_pivot_column_fit.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    assert column_widths(made[2] / "pivot_show_detail.xlsx", names) == column_widths(FIXTURES / "pivot_show_detail.xlsx", names)
