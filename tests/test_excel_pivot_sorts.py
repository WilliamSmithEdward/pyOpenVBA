"""Sorting a pivot table's items, replayed against the in-memory model.

tests/fixtures/pivot_sorts/ holds what scripts/measure_pivot_sorts.py saw
in live Excel: a pivot table on a sheet of its own for each layout, over
one list of regions, products and amounts, its items sorted by
PivotField.AutoSort by their names or a data field's values, up, down or
back to manual, and moved by PivotItem.Position; each table dumped as its
cells, each field's items with their positions, and each field's
AutoSortOrder and AutoSortField; the workbook Excel saved of them all; and
each table as Excel read it again from that file.
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
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_sorts"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_sorts.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's pivotFields, rowItems and colItems as a file keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            out[name.group(1) if name else part] = "".join(
                found.group() for tag in ("pivotFields", "rowItems", "colItems")
                if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL)))
    return out


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its dumps in the session and once the model opened
    its own file again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("pivot_sorts")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{page}", layout["page"])
                   .replace("{source}", layout["own_source"]).replace("{step}", str(index))
                   for index, layout in enumerate(LAYOUTS, 1))
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{source}", RECORD["source"])
            .replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep))
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_sorts.xlsx") == {layout["name"]: layout["file"] for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_sorted_pivot_table_shows_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                     layout: dict[str, Any]) -> None:
    """The cells, each field's items with their positions and its sort: by names or values, inside each outer
    item, ties in the list's order and descending the ascending turned round; an order or field AutoSort does not
    know moves nothing, and Position counts the items showing."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_sorted_pivot_table_reads_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_sort_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                 layout: dict[str, Any]) -> None:
    """The sortType, the data field a value sort's scope names, the items in their order, and the lines of the
    table as Excel writes them."""
    assert tables(made[2] / "pivot_sorts.xlsx")[layout["name"]] == layout["file"]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_sorts.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("statement", ['pt.PivotFields("Region").AutoSort xlAscending, "Region", 1',
                                       'pt.PivotFields("Region").PivotItems("A").Position = 1',
                                       'pt.DataFields(1).AutoSort xlAscending, "Region"'])
def test_what_is_not_measured_reports_itself(statement: str) -> None:
    """A PivotLine, an item of a page field, and a data field."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(HELPER + "Public Sub Probe()\n"
                   "Dim wb As Object, data As Object, cache As Object, r As Long, regions As Variant, "
                   "products As Variant, pt As Object\nSet wb = ActiveWorkbook\n" + RECORD["source"] +
                   'Set pt = Made(wb, wb.Worksheets.Add, cache, "Region", "", "Product")\n'
                   + statement.replace('PivotFields("Region").PivotItems("A")', 'PivotFields("Product").PivotItems("A")')
                   + "\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError):
        app.run("Probe")
