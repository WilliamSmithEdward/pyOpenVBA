"""A pivot table's data fields shown as Excel's calculations, Show Values As, replayed against the in-memory model.

tests/fixtures/pivot_calculations/ holds what scripts/measure_pivot_calculations.py
saw in live Excel: a pivot table on a sheet of its own for each layout,
over one list of regions, products, channels, amounts and units, a data
field shown as a difference from an item or from the previous or next one,
as a percent of an item, of the row, column, total or a parent, as a
running total, an index or a rank; each table dumped as its cells with
their number formats and every data field's calculation, base field, base
item and number format; the workbook Excel saved of them all; and each
table as Excel read it again from that file.
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
from test_excel_pivot_layouts import REVISIONS, caches, column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_calculations"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_calculations.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: The parts of a table's definition its calculations show in: its data fields and its own extensions.
PARTS = (r"<dataFields\b.*?</dataFields>", r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)")


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's data fields and extensions, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            out[name.group(1) if name else part] = "".join(
                REVISIONS.sub("", found.group()) for pattern in PARTS if (found := re.search(pattern, text, re.DOTALL)))
    return out


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
    folder = tmp_path_factory.mktemp("pivot_calculations")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_calculations.xlsx") == {
        layout["name"]: REVISIONS.sub("", layout["file"].replace(" />", "/>")) for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_calculations_show_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                layout: dict[str, Any]) -> None:
    """Every value cell of the table with its number format, and each data field's calculation, base field, base
    item and number format, with the errors setting them raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_calculations_read_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_calculation_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                        layout: dict[str, Any]) -> None:
    """Each data field's showDataAs, base field, base item and number format, and Excel 2010's pivotShowAs."""
    assert tables(made[2] / "pivot_calculations.xlsx")[layout["name"]] == REVISIONS.sub(
        "", layout["file"].replace(" />", "/>"))


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return sheet_cells(made[2] / "pivot_calculations.xlsx", names), sheet_cells(FIXTURES / "pivot_calculations.xlsx",
                                                                                 names)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_calculations.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_a_calculation_beside_an_inner_fields_subtotals_by_another_function() -> None:
    """The innermost field's own subtotals by another function, each over every item outside it, are not measured
    under a calculation, which reports itself unsupported there."""
    app = ExcelApplication()
    app.add_module(
        "Public Sub Run()\nDim wb As Object, ws As Object, pt As Object\n"
        "Set wb = Workbooks.Add\nSet ws = wb.Worksheets(1)\n"
        'ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")\n'
        'ws.Range("A2:C2").Value = Array("East", "Apple", 10)\nws.Range("A3:C3").Value = Array("West", "Pear", 20)\n'
        'Set pt = wb.PivotCaches.Create(1, ws.Name & "!R1C1:R3C3").CreatePivotTable(ws.Range("E3"), "T")\n'
        'pt.PivotFields("Region").Orientation = 1\npt.PivotFields("Product").Orientation = 1\n'
        'pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", -4157\n'
        'pt.PivotFields("Product").Subtotals(3) = True\n'
        "pt.DataFields(1).Calculation = 8\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError, match="inner field's subtotals"):
        app.run("Run")


def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's columns as wide as Excel fits them each time it lays the table out, bestFit beside them (see
    tests/test_excel_pivot_column_fit.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    assert column_widths(made[2] / "pivot_calculations.xlsx", names) == column_widths(FIXTURES / "pivot_calculations.xlsx", names)


def test_the_caches_are_excels(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's cache as Excel saves it: shared items listed for the fields a table lists items for, the
    other fields' values in the records, the sharedItems' attributes and the numbers spelled as Excel's (see
    src/pyopenvba/apps/excel/_pivot_cache_file.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    mine, theirs = caches(made[2] / "pivot_calculations.xlsx"), caches(FIXTURES / "pivot_calculations.xlsx")
    assert {name: mine.get(name) for name in names} == {name: theirs.get(name) for name in names}
