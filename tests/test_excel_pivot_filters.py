"""Hiding a pivot table's items and choosing its page fields' items, replayed against the in-memory model.

tests/fixtures/pivot_filters/ holds what scripts/measure_pivot_filters.py
saw in live Excel: a pivot table on a sheet of its own for each layout, over
one list of regions, products and amounts, with items hidden by
PivotItem.Visible, a page field's item chosen by CurrentPage or several by
EnableMultiplePageItems, filters cleared, fields moved and caches
refreshed; each table dumped as its cells, each field's items and whether
each shows, its page field's choice and its grand total; the workbook
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
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: The layouts whose sheet the macro deletes before the save: Excel cannot open the file it writes of them.
UNSAVED = {name for name in re.findall(r'Worksheets\("([^"]+)"\)\.Delete', RECORD["unsaved"])}
SAVED = [layout for layout in LAYOUTS if layout["name"] not in UNSAVED]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])

def tables(path: Path) -> dict[str, tuple[str, str]]:
    """Each pivot table's pivotFields and pageFields, and its cache's cacheFields, as a file keeps them, by the
    table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, tuple[str, str]] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            fields = re.search(r"<pivotFields\b.*?</pivotFields>", text, re.DOTALL)
            pages = re.search(r"<pageFields\b.*?</pageFields>", text, re.DOTALL)
            rels = package.read(part.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode("utf-8")
            target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
            cache = package.read(f"xl/pivotCache/{target.group(1)}").decode("utf-8") if target else ""
            cached = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = ((fields.group() if fields else "") + (pages.group() if pages else ""),
                                                    cached.group() if cached else "")
    return out


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def shared_values(cache: str, fields: str) -> list[list[str]]:
    """The shared items of each cache field a table lists items for, each as its value and whether it is unused."""
    listed = [bool(re.search(r"<items\b", field))
              for field in re.findall(r"<pivotField\b[^>]*?(?:/>|>.*?</pivotField>)", fields, re.DOTALL)]
    found: list[list[str]] = []
    for index, field in enumerate(re.findall(r"<cacheField\b.*?</cacheField>", cache, re.DOTALL)):
        if index < len(listed) and listed[index]:
            found.append([tag + ":" + value + (" unused" if 'u="1"' in rest else "")
                          for tag, value, rest in re.findall(r'<(\w)\b v="([^"]*)"([^>]*)/>', field)])
    return found


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model by the macro Excel ran, its dumps in the session and once the model opened
    its own file again, and the folder it saved to."""
    folder = tmp_path_factory.mktemp("pivot_filters")
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{page}", layout["page"])
                   .replace("{source}", layout["own_source"]).replace("{step}", str(index))
                   for index, layout in enumerate(LAYOUTS, 1))
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{source}", RECORD["source"])
            .replace("{layouts}", body).replace("{unsaved}", RECORD["unsaved"])
            .replace("{folder}\\", str(folder) + os.sep))
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    session, reopened = str(app.run("Build")).split("~", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    held = tables(FIXTURES / "pivot_filters.xlsx")
    assert {name: fields for name, (fields, _cache) in held.items()} == {
        layout["name"]: layout["file"] for layout in SAVED}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_filtered_pivot_table_shows_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                       layout: dict[str, Any]) -> None:
    """The cells, each field's items and whether each shows, the page field's choice and the grand total: hidden
    items and items a page field leaves out count nowhere, the last item showing cannot be hidden, and the model
    errs where Excel does."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_filtered_pivot_table_reads_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1].get(layout["name"], "") == layout["reopened"]


@pytest.mark.parametrize("layout", SAVED, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_filter_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                   layout: dict[str, Any]) -> None:
    """Each pivotField and pageField: the items in their order, hidden and gone ones flagged, the multiple choice
    allowed and the chosen item, written as Excel writes them."""
    fields, _cache = tables(made[2] / "pivot_filters.xlsx")[layout["name"]]
    assert fields.replace(" />", "/>") == layout["file"]


@pytest.mark.parametrize("layout", [one for one in SAVED if one["own_source"]], ids=lambda layout: layout["name"])
def test_a_refreshed_cache_keeps_the_values_a_report_lists_as_excels_does(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    """Values in the order the refreshed source gives them, then each value a report lists that the source lost,
    kept unused."""
    fields, cache = tables(made[2] / "pivot_filters.xlsx")[layout["name"]]
    assert shared_values(cache.replace(" />", "/>"), fields) == shared_values(layout["cache"], layout["file"])


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_filters.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", SAVED, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_a_page_field_choosing_an_item_and_allowing_several_saves_as_choosing_it(tmp_path: Path) -> None:
    """Excel saves such a field and cannot open the file it wrote (page_then_multiple); the model saves it as
    choosing the one item, which opens showing the same item."""
    app = ExcelApplication()
    app.add_workbook()
    layout = next(one for one in LAYOUTS if one["name"] == "page_then_multiple")
    index = LAYOUTS.index(layout) + 1
    app.add_module(HELPER + "Public Function One() As String\n"
                   "Dim wb As Object, data As Object, cache As Object, out As String, r As Long, regions As Variant\n"
                   "Dim products As Variant\nSet wb = ActiveWorkbook\n" + RECORD["source"] +
                   f'out = Layout(wb, cache, "page_then_multiple", "Region", "", "Product", "", {index})\n'
                   "One = out\nEnd Function\n", name="Builder")
    assert dumps(str(app.run("One")))["page_then_multiple"] == layout["dump"]
    app.save(tmp_path / "chosen.xlsx")
    fields, _cache = tables(tmp_path / "chosen.xlsx")["page_then_multiple"]
    assert 'multipleItemSelectionAllowed' not in fields
    assert '<pageField fld="1" item="0" hier="-1"' in fields.replace(" />", "/>")
    again = ExcelApplication.open(tmp_path / "chosen.xlsx", with_vba=False)
    again.add_module(HELPER + 'Public Function Read() As String\nRead = Dump(Worksheets("page_then_multiple")'
                     '.PivotTables(1))\nEnd Function\n', name="Reader")
    assert str(again.run("Read")) == layout["dump"].replace(
        "A=True,B=True,C=True,", "A=True,B=False,C=False,").replace("multiple=True", "multiple=False")


def test_a_data_field_cannot_be_filtered_by_its_items() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(HELPER + "Public Sub Probe()\n"
                   "Dim wb As Object, data As Object, cache As Object, r As Long, regions As Variant, "
                   "products As Variant, pt As Object\nSet wb = ActiveWorkbook\n" + RECORD["source"] +
                   'Set pt = Made(wb, wb.Worksheets.Add, cache, "Region", "", "")\n'
                   "pt.DataFields(1).EnableMultiplePageItems = True\nEnd Sub\n", name="Probe")
    with pytest.raises(VBAUnsupportedError, match="by its values"):
        app.run("Probe")


def test_a_sheet_renames_beside_a_cache_whose_source_text_is_r1c1() -> None:
    """A recorded macro gives a cache its source as R1C1 text naming the sheet; another sheet still renames."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Probe() As String\nDim ws As Object, cache As Object\n'
                   'Set ws = ActiveWorkbook.Worksheets(1)\nws.Name = "data"\n'
                   'ws.Range("A1:B2").Value = Array("Region", "Amount")\nws.Range("A2:B2").Value = Array("East", 10)\n'
                   'Set cache = ActiveWorkbook.PivotCaches.Create(xlDatabase, "data!R1C1:R2C2")\n'
                   'Set ws = ActiveWorkbook.Worksheets.Add\n'
                   'cache.CreatePivotTable ws.Range("A3"), "Report"\nws.Name = "report"\n'
                   'Probe = ws.Name & "|" & ws.PivotTables(1).SourceData\nEnd Function\n', name="Probe")
    assert app.run("Probe") == "report|data!R1C1:R2C2"


def test_the_file_keeps_hidden_flags_ahead_of_the_place_as_excel_does(
        made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    with zipfile.ZipFile(made[2] / "pivot_filters.xlsx") as package:
        text = "".join(package.read(name).decode("utf-8") for name in package.namelist()
                       if name.startswith("xl/pivotTables/pivotTable"))
    assert '<item h="1" x=' in text and '<item m="1" x=' in text and '<item x="0" h=' not in text
