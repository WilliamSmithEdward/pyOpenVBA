"""A pivot table's options, replayed against the in-memory model.

tests/fixtures/pivot_options/ holds what scripts/measure_pivot_options.py
saw in live Excel: a pivot table on a sheet of its own for each layout,
over one list of regions, products, channels, amounts and units, with its
style and style flags, the text its empty and error cells show, its page
fields' order and wrap, its flags, tag and alternative text set, or laid
out again by AddFields, ClearTable, Location and ManualUpdate; each table
dumped as its cells, its ranges, every option and the fields it shows;
the workbook Excel saved of them all; and each table as Excel read it
again from that file.
"""
from __future__ import annotations

import json
import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_pivot_layouts import REVISIONS, caches, column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_options"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_options.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: The parts of a table's definition its options show in.
PARTS = (r"<pivotTableDefinition\b[^>]*>", r"<location\b[^>]*/>", r"<pivotTableStyleInfo\b[^>]*/>",
         r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)")


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's attributes, location, style and extensions, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            out[name.group(1) if name else part] = "".join(
                REVISIONS.sub("", found.group()) for pattern in PARTS if (found := re.search(pattern, text, re.DOTALL)))
    return out


def recorded(layout: dict[str, Any]) -> str:
    """The table's parts as the record keeps them, the cache's attributes after them left out: the record holds the
    table's extensions without the end of the table after them."""
    text = layout["file"].replace(" />", "/>")
    patterns = (*PARTS[:3], r"<extLst>(?:(?!</extLst>).)*</extLst>")
    return "".join(REVISIONS.sub("", found.group()) for pattern in patterns
                   if (found := re.search(pattern, text, re.DOTALL)))


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it."""
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{page}", layout["page"])
                   .replace("{source}", layout["own_source"]).replace("{step}", str(index))
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
    folder = tmp_path_factory.mktemp("pivot_options")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_options.xlsx") == {layout["name"]: recorded(layout) for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_options_show_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                           layout: dict[str, Any]) -> None:
    """The cells, the ranges, every option and the fields shown: the style, the null and error strings in the
    value cells, the page fields' order, AddFields, ClearTable, Location, ManualUpdate's captions alone, and the
    errors the options raise."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_options_read_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_option_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                   layout: dict[str, Any]) -> None:
    """The table's attributes and their order, its location, its style and its Excel 2010 options."""
    assert tables(made[2] / "pivot_options.xlsx")[layout["name"]] == recorded(layout)


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return sheet_cells(made[2] / "pivot_options.xlsx", names), sheet_cells(FIXTURES / "pivot_options.xlsx", names)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_options.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_a_cache_saved_without_its_records() -> None:
    """SaveData off on a cache of its own: the file keeps saveData="0" and no records, as Excel's does (a probe in
    Excel), and another table made from the cache turns it back on."""
    app = ExcelApplication()
    app.add_module(
        "Public Function Run(path As String) As String\nDim wb As Object, ws As Object, pt As Object, out As String\n"
        "Set wb = Workbooks.Add\nSet ws = wb.Worksheets(1)\n"
        'ws.Range("A1:B1").Value = Array("Region", "Amount")\nws.Range("A2:B2").Value = Array("East", 10)\n'
        'ws.Range("A3:B3").Value = Array("West", 20)\n'
        'Set pt = wb.PivotCaches.Create(1, ws.Name & "!R1C1:R3C2").CreatePivotTable(ws.Range("D3"), "T")\n'
        'pt.PivotFields("Region").Orientation = 1\npt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", -4157\n'
        "pt.SaveData = False\nout = CStr(pt.SaveData)\nwb.SaveAs path, 51\n"
        'wb.PivotCaches(1).CreatePivotTable ws.Range("H3"), "T2"\nout = out & "/" & CStr(pt.SaveData)\n'
        "Run = out\nEnd Function\n", name="Probe")
    with tempfile.TemporaryDirectory() as place:
        path = Path(place) / "save_data.xlsx"
        assert app.run("Run", str(path)) == "False/True"
        with zipfile.ZipFile(path) as package:
            parts = [name for name in package.namelist() if "pivotCache" in name]
            cache = package.read("xl/pivotCache/pivotCacheDefinition1.xml").decode("utf-8")
        assert parts == ["xl/pivotCache/pivotCacheDefinition1.xml"]
        assert 'saveData="0"' in cache


def test_a_table_in_no_style() -> None:
    """A table left in no style reads TableStyle2 as a NUL, as Excel's does (a probe in Excel); putting one in no
    style, which Excel draws with borders of its cells, reports itself unsupported."""
    app = ExcelApplication()
    app.add_module(
        "Public Function Run(name As String) As String\nDim wb As Object, ws As Object, pt As Object\n"
        "Set wb = Workbooks.Add\nSet ws = wb.Worksheets(1)\n"
        'ws.Range("A1:B1").Value = Array("Region", "Amount")\nws.Range("A2:B2").Value = Array("East", 10)\n'
        'Set pt = wb.PivotCaches.Create(1, ws.Name & "!R1C1:R2C2").CreatePivotTable(ws.Range("D3"), "T")\n'
        "pt.TableStyle2 = name\nRun = pt.TableStyle2\nEnd Function\n", name="Style")
    assert app.run("Run", "pivotstylemedium4") == "PivotStyleMedium4"
    table = app.workbook.sheets_[0].pivot_tables[0]
    info = table.root.find(f"{NS}pivotTableStyleInfo")
    assert info is not None
    del info.attrib["name"]
    assert table.vba_get("TableStyle2") == "\x00"
    for name in ("", "NoSuchStyle", "PivotStyleLight29"):
        with pytest.raises(VBAUnsupportedError, match="no style"):
            app.run("Run", name)


def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's columns as wide as Excel fits them each time it lays the table out, bestFit beside them (see
    tests/test_excel_pivot_column_fit.py): the style layout's fit to its style flags, which wait for a layout, as
    the clear_table layout's ClearTable lays out every table of the cache they share."""
    names = {layout["name"] for layout in LAYOUTS}
    assert column_widths(made[2] / "pivot_options.xlsx", names) == column_widths(FIXTURES / "pivot_options.xlsx", names)


def test_the_caches_are_excels(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's cache as Excel saves it: shared items listed for the fields a table lists items for, the
    other fields' values in the records, the sharedItems' attributes and the numbers spelled as Excel's (see
    src/pyopenvba/apps/excel/_pivot_cache_file.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    mine, theirs = caches(made[2] / "pivot_options.xlsx"), caches(FIXTURES / "pivot_options.xlsx")
    assert {name: mine.get(name) for name in names} == {name: theirs.get(name) for name in names}
