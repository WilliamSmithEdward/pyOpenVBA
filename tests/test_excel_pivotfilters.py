"""A pivot table's label, value and top 10 filters, replayed against the in-memory model.

tests/fixtures/pivotfilters/ holds what scripts/measure_pivotfilters.py saw
in live Excel: a pivot table on a sheet of its own for each layout, over
one list of regions, products and amounts, with a field's PivotFilters
given a label filter, a value filter or a top 10 filter, cleared, deleted,
mixed with hidden items and refreshed; each table dumped as its cells,
each field's items and whether each shows, and each field's filters with
their types, values, data fields and whether active; the workbook Excel
saved of them all; and each table as Excel read it again from that file.
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

FIXTURES = Path(__file__).parent / "fixtures" / "pivotfilters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivotfilters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])


def tables(path: Path) -> dict[str, tuple[str, str]]:
    """Each pivot table's multipleFieldFilters, and its pivotFields, rowItems, colItems and filters, as a file keeps
    them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, tuple[str, str]] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            head = re.search(r"<pivotTableDefinition\b[^>]*>", text)
            multiple = re.search(r'multipleFieldFilters="[^"]*"', head.group() if head else "")
            kept = [found.group() for tag in ("pivotFields", "rowItems", "colItems")
                    if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL))]
            filters = re.search(r'<filters count="\d+">.*?</filter></filters>', text, re.DOTALL)
            out[name.group(1) if name else part] = (multiple.group() if multiple else "",
                                                    "".join(kept) + (filters.group() if filters else ""))
    return out


def held(layout: dict[str, Any]) -> tuple[str, str]:
    """What the record says Excel's file keeps of a layout's table, as tables() reads it."""
    text = layout["file"]
    head = re.match(r"<pivotTableDefinition\b[^>]*>", text)
    multiple = re.search(r'multipleFieldFilters="[^"]*"', head.group() if head else "")
    return (multiple.group() if multiple else "", text[head.end():] if head else text)


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its dumps in the session and once the model opened
    its own file again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("pivotfilters")
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
    assert tables(FIXTURES / "pivotfilters.xlsx") == {layout["name"]: held(layout) for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_filtered_pivot_table_shows_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                       layout: dict[str, Any]) -> None:
    """The cells, each field's items and whether each shows, and each field's filters read back: an item's caption
    against a label filter, its total against a value filter inside its outer items, top 10 filters with ties
    kept, label filters before value filters, and the errors Add raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_filtered_pivot_table_reads_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_filter_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                   layout: dict[str, Any]) -> None:
    """Each filter's type, ids, values and the autoFilter spelling its test, a value filter's field marked, and
    multipleFieldFilters, as Excel writes them."""
    assert tables(made[2] / "pivotfilters.xlsx")[layout["name"]] == held(layout)


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivotfilters.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("statement", [
    'pt.PivotFields("Region").PivotFilters.Add Type:=xlCaptionEquals, Value1:="N", Name:="mine"',
    'pt.PivotFields("Region").PivotFilters.Add Type:=xlTopPercent, DataField:=pt.DataFields(1), Value1:=0',
    'pt.PivotFields("Product").Orientation = xlPageField'])
def test_what_is_not_measured_reports_itself(statement: str) -> None:
    """Add's other arguments, a top percent of 0, and a filtered field moved to the page area."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(HELPER + "Public Sub Probe()\n"
                   "Dim wb As Object, data As Object, cache As Object, r As Long, regions As Variant, "
                   "products As Variant, pt As Object\nSet wb = ActiveWorkbook\n" + RECORD["source"] +
                   'Set pt = Made(wb, wb.Worksheets.Add, cache, "Region", "Product", "")\n'
                   'pt.PivotFields("Product").PivotFilters.Add Type:=xlCaptionEquals, Value1:="B"\n'
                   + statement + "\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError):
        app.run("Probe")


def test_add2_is_add() -> None:
    """The type library's name for Add since Excel 2013."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(HELPER + "Public Function Probe() As String\n"
                   "Dim wb As Object, data As Object, cache As Object, r As Long, regions As Variant, "
                   "products As Variant, pt As Object\nSet wb = ActiveWorkbook\n" + RECORD["source"] +
                   'Set pt = Made(wb, wb.Worksheets.Add, cache, "Region", "", "")\n'
                   'pt.PivotFields("Region").PivotFilters.Add2 Type:=xlCaptionContains, Value1:="o"\n'
                   "Probe = Dump(pt)\nEnd Function\n", name="Probe")
    contains = next(layout for layout in LAYOUTS if layout["name"] == "contains")
    assert app.run("Probe") == contains["dump"]
