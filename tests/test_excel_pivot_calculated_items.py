"""A pivot field's calculated items, PivotField.CalculatedItems, replayed against the in-memory model.

tests/fixtures/pivot_calculated_items/ holds what
scripts/measure_pivot_calculated_items.py saw in live Excel: a pivot table on
a sheet of its own for each layout, over one list of regions, products,
dates, amounts and units, with an item added to a field whose formula names
the field's other items, on the rows or the columns, inside or outside
another field, beside another or one it names, by every data field
function, hidden, renamed, sorted, changed, deleted and refreshed, and the
errors that raises; each table dumped as its cells and its fields' items,
each calculated one with its formula; the workbook Excel saved of them all,
each table's definition and its cache's fields and calculated items; and
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
from test_excel_pivot_layouts import sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_calculated_items"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_calculated_items.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: What Excel writes afresh each time: revision ids, cache numbering and refresh stamps.
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')
#: A cache field's shared items but for its calculated items, which the model lists for every field where Excel
#: lists only those a table shows (docs/host_completeness.md).
SHARED = re.compile(r"<sharedItems\b[^>]*/>|<sharedItems\b[^>]*>.*?</sharedItems>", re.DOTALL)
CALCULATED = re.compile(r'<s v="[^"]*" f="1"/>')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition, its cache's fields and its calculated items, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            rels = package.read(f"xl/pivotTables/_rels/{part.rsplit('/', 1)[1]}.rels")
            target = re.search(rb'Target="([^"]*pivotCacheDefinition\d+\.xml)"', rels)
            cache = ""
            if target:
                cache = package.read("xl/" + target.group(1).decode().replace("../", "")).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            body = re.sub(r"^<\?xml[^>]*\?>\s*", "", text)
            fields = re.search(r"<cacheFields\b.*?</cacheFields>(?:<calculatedItems\b.*?</calculatedItems>)?",
                               cache.replace(" />", "/>"), re.DOTALL)
            out[name.group(1) if name else part] = REVISIONS.sub("", body + (fields.group() if fields else ""))
    return out


def recorded(layout: dict[str, Any]) -> str:
    return REVISIONS.sub("", layout["file"].replace(" />", "/>"))


def kept(text: str) -> str:
    """A table's parts with its cache fields' shared items left out, their calculated items kept."""
    return SHARED.sub(lambda found: "".join(CALCULATED.findall(found.group())), text)


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
    folder = tmp_path_factory.mktemp("pivot_calculated_items")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    names = {layout["name"] for layout in LAYOUTS}
    assert {name: text for name, text in tables(FIXTURES / "pivot_calculated_items.xlsx").items()
            if name in names} == {layout["name"]: recorded(layout) for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_items_show_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                       layout: dict[str, Any]) -> None:
    """Every cell of the table and each field's items, a calculated one with its formula, with the errors adding,
    changing and deleting them raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_items_read_back_from_the_file_the_model_saved(made: tuple[dict[str, str], dict[str, str], Path],
                                                         layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_item_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                               layout: dict[str, Any]) -> None:
    """The table's definition whole, and its cache's fields and calculated items: each one's shared item, its
    formula as the cache keeps it and the pivotArea naming it."""
    assert kept(tables(made[2] / "pivot_calculated_items.xlsx")[layout["name"]]) == kept(recorded(layout))


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return (sheet_cells(made[2] / "pivot_calculated_items.xlsx", names),
            sheet_cells(FIXTURES / "pivot_calculated_items.xlsx", names))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_calculated_items.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(ws.Name)) & "
                   "\"|\"\nNext\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]
