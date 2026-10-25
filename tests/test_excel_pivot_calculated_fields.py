"""A pivot table's calculated fields, replayed against the in-memory model.

tests/fixtures/pivot_calculated_fields/ holds what
scripts/measure_pivot_calculated_fields.py saw in live Excel: a pivot
table on a sheet of its own for each layout, each over a cache of its own
of one list of regions, products, channels, amounts, units and unit costs,
with calculated fields added, put in the data area, their formulas read,
changed and spelled back, renamed and deleted; each table dumped as its
cells with their number formats, every data field's name, function and
source, and every calculated field's formula; the workbook Excel saved of
them all, its tables and cache fields; and each table as Excel read it
again from that file.
"""
from __future__ import annotations

import json
import os
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_pivot_layouts import column_widths, sheet_cells

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_calculated_fields"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_calculated_fields.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
#: Every table of the file by name, as the record keeps it: each layout's, and those a layout makes beside its own.
TABLES: dict[str, str] = {**{layout["name"]: layout["file"] for layout in LAYOUTS}, **RECORD["others"]}
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: What Excel writes afresh each time: revision ids, cache numbering, who refreshed the cache and when.
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')


def kept(text: str) -> str:
    """A table's definition and its cache's fields as compared: the definition whole, and of the cache's fields the
    calculated ones; how the model keeps the source fields' values is not what this measures."""
    head, _, cache = text.partition("<cacheFields")
    calculated = re.findall(r'<cacheField\b[^>]*\bdatabaseField="0"[^>]*/>', cache)
    return REVISIONS.sub("", head) + "".join(calculated)


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition and its cache's fields, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            rels = package.read(posixpath.join("xl/pivotTables/_rels", posixpath.basename(part) + ".rels"))
            target = re.search(rb'Target="([^"]*pivotCacheDefinition\d+\.xml)"', rels)
            cache = ""
            if target:
                cache = package.read(posixpath.normpath(posixpath.join(
                    "xl/pivotTables", target.group(1).decode()))).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            body = re.sub(r"^<\?xml[^>]*\?>\s*", "", text)
            fields = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = kept(body + (fields.group() if fields else ""))
    return out


def recorded(name: str) -> str:
    return kept(TABLES[name].replace(" />", "/>"))


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
    folder = tmp_path_factory.mktemp("pivot_calculated_fields")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_calculated_fields.xlsx") == {name: recorded(name) for name in TABLES}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_calculated_fields_show_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                                     layout: dict[str, Any]) -> None:
    """Every cell of the table with its number format, each data field's name, function and source, and each
    calculated field's formula as Excel spells it back, with the errors adding and changing them raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_calculated_fields_read_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("name", list(TABLES))
def test_the_file_keeps_each_calculated_field_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                             name: str) -> None:
    """The table's definition whole, and its cache's calculated fields, each with its formula as Excel spells it;
    a second table of a cache takes the cache's calculated fields too."""
    assert tables(made[2] / "pivot_calculated_fields.xlsx")[name] == recorded(name)


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return (sheet_cells(made[2] / "pivot_calculated_fields.xlsx", names),
            sheet_cells(FIXTURES / "pivot_calculated_fields.xlsx", names))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_calculated_fields.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_a_refused_formula_with_alerts_on() -> None:
    """A formula Excel refuses adds nothing with alerts off, as measured; with them on Excel shows a message, which
    reports itself unsupported."""
    app = ExcelApplication()
    app.add_module(
        "Public Function Run(alerts As Boolean) As String\nDim wb As Object, ws As Object, pt As Object\n"
        "Set wb = Workbooks.Add\nSet ws = wb.Worksheets(1)\n"
        'ws.Range("A1:B1").Value = Array("Region", "Amount")\nws.Range("A2:B2").Value = Array("East", 10)\n'
        'Set pt = wb.PivotCaches.Create(1, ws.Name & "!R1C1:R2C2").CreatePivotTable(ws.Range("D3"), "T")\n'
        "Application.DisplayAlerts = alerts\n"
        'pt.CalculatedFields.Add "Bad", "=Amount*"\nRun = CStr(pt.CalculatedFields.Count)\nEnd Function\n',
        name="Probe")
    assert app.run("Run", False) == "0"
    with pytest.raises(VBAUnsupportedError, match="shows a message"):
        app.run("Run", True)


def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's columns as wide as Excel fits them each time it lays the table out, bestFit beside them (see
    tests/test_excel_pivot_column_fit.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    assert column_widths(made[2] / "pivot_calculated_fields.xlsx", names) == column_widths(FIXTURES / "pivot_calculated_fields.xlsx", names)
