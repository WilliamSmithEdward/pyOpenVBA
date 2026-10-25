"""A pivot table's layout and its fields' subtotals, replayed against the in-memory model.

tests/fixtures/pivot_layouts/ holds what scripts/measure_pivot_layouts.py
saw in live Excel: a pivot table on a sheet of its own for each layout,
over one list of regions, products, channels, amounts and units, laid out
compact, outline or tabular by RowAxisLayout or by a field's LayoutForm
and LayoutCompactRow, with its fields' Subtotals, LayoutSubtotalLocation,
LayoutBlankLine, LayoutPageBreak, RepeatLabels and SubtotalName and its
data fields' Function set, and the table's RepeatAllLabels, captions,
grand totals and indent; each table dumped as its cells, each with its
indent and whether it is merged, the table's ranges and layout properties
and each field's; the workbook Excel saved of them all; and each table as
Excel read it again from that file.
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

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_layouts"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_layouts.json").read_text(encoding="utf-8"))
#: MergeLabels merges and centres the label cells, which the model reports unsupported.
UNSUPPORTED = {"merge_labels"}
LAYOUTS: list[dict[str, Any]] = [layout for layout in RECORD["layouts"] if layout["name"] not in UNSUPPORTED]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
KEPT = ("location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "extLst")
#: Excel's revision namespace and the random id it gives each table, and the number it gives each cache, which
#: no other writer reproduces.
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition as far as its layout shows in it, by the table's name: its attributes, the
    location, the fields, lines and the table's own extensions."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            start = re.search(r"<pivotTableDefinition\b[^>]*>", text)
            body = re.sub(r"<pivotFields\b.*?</pivotFields>", "", text, flags=re.DOTALL)
            kept = [REVISIONS.sub("", start.group()) if start else ""]
            for tag in KEPT:
                pattern = rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>"
                if tag == "extLst":
                    pattern = r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)"
                found = re.search(pattern, text if tag == "pivotFields" else body, re.DOTALL)
                if found:
                    kept.append(found.group())
            out[name.group(1) if name else part] = "".join(kept)
    return out


def sheet_cells(path: Path, names: set[str]) -> dict[str, dict[str, tuple[object, str, int]]]:
    """Each named sheet's cells: the value, the format the cell's xf writes out with its number format's code in
    place of its id, and the xf's place. A number is compared as the double it is: the model spells a cell's
    value in the fewest digits that read back the same, where Excel often writes seventeen."""
    with zipfile.ZipFile(path) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        relations = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
        shared = (package.read("xl/sharedStrings.xml").decode("utf-8")
                  if "xl/sharedStrings.xml" in package.namelist() else "")
        strings = [re.sub(r"<[^>]+>", "", one) for one in re.findall(r"<si>(.*?)</si>", shared, re.DOTALL)]
        codes = {int(found.group(1)): found.group(2)
                 for found in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles)}
        cell_xfs = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.DOTALL)
        assert cell_xfs is not None
        xfs = re.findall(r"<xf\b[^>]*/>|<xf\b.*?</xf>", cell_xfs.group(1), re.DOTALL)
        targets: dict[str, str] = {}
        for found in re.finditer(r"<Relationship\b[^>]*>", relations):
            ident = re.search(r'Id="([^"]+)"', found.group())
            target = re.search(r'Target="([^"]+)"', found.group())
            if ident and target:
                targets[ident.group(1)] = target.group(1)
        out: dict[str, dict[str, tuple[object, str, int]]] = {}
        for sheet in re.finditer(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', book):
            if sheet.group(1) not in names:
                continue
            xml = package.read("xl/" + targets[sheet.group(2)].lstrip("/").removeprefix("xl/")).decode("utf-8")
            cells: dict[str, tuple[object, str, int]] = {}
            for cell in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.DOTALL):
                style = re.search(r's="(\d+)"', cell.group(2))
                index = int(style.group(1)) if style else 0
                found_value = re.search(r"<v>(.*?)</v>", cell.group(3) or "")
                inline = re.search(r"<is>(.*?)</is>", cell.group(3) or "", re.DOTALL)
                kind = re.search(r't="(\w+)"', cell.group(2))
                value: object = None
                if inline:
                    value = re.sub(r"<[^>]+>", "", inline.group(1))
                elif found_value and kind and kind.group(1) == "s":
                    value = strings[int(found_value.group(1))]
                elif found_value and kind and kind.group(1) != "n":
                    value = found_value.group(1)
                elif found_value:
                    value = float(found_value.group(1))
                xf = xfs[index]
                number = re.search(r'numFmtId="(\d+)"', xf)
                if number and int(number.group(1)) in codes:
                    xf = xf.replace(number.group(), f'numFmt="{codes[int(number.group(1))]}"')
                cells[cell.group(1)] = (value, xf, index)
            out[sheet.group(1)] = cells
        return out


def column_widths(path: Path, names: set[str]) -> dict[str, dict[int, tuple[str, bool]]]:
    """Each named sheet's columns that have a width of their own: the width the file writes and whether it is the
    one AutoFit gave the column, bestFit (tests/test_excel_pivot_column_fit.py)."""
    with zipfile.ZipFile(path) as package:
        book = package.read("xl/workbook.xml").decode("utf-8")
        relations = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets: dict[str, str] = {}
        for found in re.finditer(r"<Relationship\b[^>]*>", relations):
            ident = re.search(r'Id="([^"]+)"', found.group())
            target = re.search(r'Target="([^"]+)"', found.group())
            if ident and target:
                targets[ident.group(1)] = target.group(1)
        out: dict[str, dict[int, tuple[str, bool]]] = {}
        for sheet in re.finditer(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', book):
            if sheet.group(1) not in names:
                continue
            xml = package.read("xl/" + targets[sheet.group(2)].lstrip("/").removeprefix("xl/")).decode("utf-8")
            columns: dict[int, tuple[str, bool]] = {}
            for col in re.findall(r"<col\b[^>]*>", xml):
                first = re.search(r'\bmin="(\d+)"', col)
                last = re.search(r'\bmax="(\d+)"', col)
                width = re.search(r'\bwidth="([^"]*)"', col)
                if first and last and width and "customWidth=\"1\"" in col:
                    for number in range(int(first.group(1)), min(int(last.group(1)), 16384) + 1):
                        columns[number] = (width.group(1), 'bestFit="1"' in col)
            out[sheet.group(1)] = columns
        return out


#: What a save writes afresh in a cache: its refresh stamp and author, and the namespaces and ids Excel adds.
CACHE_STAMPS = re.compile(r' refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"| xr:uid="[^"]*"'
                          r'| xmlns:\w+="[^"]*"| mc:Ignorable="[^"]*"')


def caches(path: Path) -> dict[str, tuple[str, str]]:
    """Each pivot table's cache as the file keeps it, by the table's name: its fields, their shared items among
    them, and its records (tests/fixtures/pivot_grouping/ and the rest)."""
    out: dict[str, tuple[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())

        def related(part: str, kind: str) -> str | None:
            folder, name = part.rsplit("/", 1)
            rels = f"{folder}/_rels/{name}.rels"
            if rels not in names:
                return None
            target = re.search(rf'Target="([^"]*{kind}\d+\.xml)"', package.read(rels).decode("utf-8"))
            return posixpath.normpath(posixpath.join(folder, target.group(1))) if target else None

        for part in sorted(name for name in names if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            table = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            cache = related(part, "pivotCacheDefinition")
            if table is None or cache is None:
                continue
            definition = package.read(cache).decode("utf-8").replace(" />", "/>")
            fields = re.search(r"<cacheFields\b.*?</cacheFields>", definition, re.DOTALL)
            records = related(cache, "pivotCacheRecords")
            body = (re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(records).decode("utf-8")).replace(" />", "/>")
                    if records else "")
            out[table.group(1)] = (CACHE_STAMPS.sub("", fields.group() if fields else ""), CACHE_STAMPS.sub("", body))
    return out


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it, the unsupported layouts left out."""
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{page}", layout["page"])
                   .replace("{source}", layout["own_source"]).replace("{step}", str(index))
                   for index, layout in enumerate(RECORD["layouts"], 1) if layout["name"] not in UNSUPPORTED)
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
    folder = tmp_path_factory.mktemp("pivot_layouts")
    session, reopened = str(built(folder).run("Build")).split("@@", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_layouts.xlsx") == {
        layout["name"]: REVISIONS.sub("", layout["file"].replace(" />", "/>")) for layout in RECORD["layouts"]}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_layout_shows_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                         layout: dict[str, Any]) -> None:
    """The cells, each with its indent, the table's ranges and layout, and each field's: the label columns and
    lines of each form, subtotals at the top, below or by their functions, blank lines, repeated labels, the
    captions, and the errors the members raise."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_layout_reads_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_layout_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                   layout: dict[str, Any]) -> None:
    """The table's flags, the fields' forms, subtotals and extensions, the items and lines, the location and the
    table's extensions, as Excel writes them."""
    assert tables(made[2] / "pivot_layouts.xlsx")[layout["name"]] == REVISIONS.sub(
        "", layout["file"].replace(" />", "/>"))


@pytest.fixture(scope="module")
def formats(made: tuple[dict[str, str], dict[str, str], Path]) -> tuple[
        dict[str, dict[str, tuple[object, str, int]]], dict[str, dict[str, tuple[object, str, int]]]]:
    names = {layout["name"] for layout in LAYOUTS}
    return sheet_cells(made[2] / "pivot_layouts.xlsx", names), sheet_cells(FIXTURES / "pivot_layouts.xlsx", names)


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_cells_keep_excels_formats(formats: tuple[dict[str, dict[str, tuple[object, str, int]]],
                                                      dict[str, dict[str, tuple[object, str, int]]]],
                                       layout: dict[str, Any]) -> None:
    """Each cell's value and format, and the format's place in the stylesheet: a field's button, a label aligned
    left and indented by its depth where a field is compact, a field's number format, and every value cell in its
    data field's format, an empty one too."""
    assert formats[0][layout["name"]] == formats[1][layout["name"]]


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it."""
    app = ExcelApplication.open(FIXTURES / "pivot_layouts.xlsx", with_vba=False)
    app.add_module(HELPER + "Public Function Read() As String\nDim ws As Object, out As String\n"
                   "For Each ws In ActiveWorkbook.Worksheets\n"
                   "If ws.PivotTables.Count > 0 Then out = out & ws.Name & \"^\" & Dump(ws.PivotTables(1)) & \"|\"\n"
                   "Next\nRead = out\nEnd Function\n", name="Reader")
    return dumps(str(app.run("Read")))


@pytest.mark.parametrize("layout", RECORD["layouts"], ids=lambda layout: layout["name"])
def test_the_model_reads_excels_file_as_excel_does(excel_file_read: dict[str, str], layout: dict[str, Any]) -> None:
    assert excel_file_read[layout["name"]] == layout["reopened"]


def test_merging_the_labels_reports_itself_unsupported() -> None:
    app = ExcelApplication()
    app.add_module(
        "Public Sub Run()\nDim wb As Object, ws As Object, pt As Object\nSet wb = Workbooks.Add\n"
        'Set ws = wb.Worksheets(1)\nws.Range("A1:B1").Value = Array("Region", "Amount")\n'
        'ws.Range("A2:B2").Value = Array("East", 1)\nws.Range("A3:B3").Value = Array("West", 2)\n'
        'Set pt = wb.PivotCaches.Create(1, ws.Name & "!R1C1:R3C2").CreatePivotTable(ws.Range("D3"), "T")\n'
        'pt.PivotFields("Region").Orientation = 1\npt.MergeLabels = True\nEnd Sub\n', name="Merge")
    with pytest.raises(VBAUnsupportedError, match="MergeLabels"):
        app.run("Run")


def test_the_columns_fit_as_excels_do(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's columns as wide as Excel fits them each time it lays the table out, bestFit beside them (see
    tests/test_excel_pivot_column_fit.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    assert column_widths(made[2] / "pivot_layouts.xlsx", names) == column_widths(FIXTURES / "pivot_layouts.xlsx", names)


def test_the_caches_are_excels(made: tuple[dict[str, str], dict[str, str], Path]) -> None:
    """Every table's cache as Excel saves it: shared items listed for the fields a table lists items for, the
    other fields' values in the records, the sharedItems' attributes and the numbers spelled as Excel's (see
    src/pyopenvba/apps/excel/_pivot_cache_file.py)."""
    names = {layout["name"] for layout in LAYOUTS}
    mine, theirs = caches(made[2] / "pivot_layouts.xlsx"), caches(FIXTURES / "pivot_layouts.xlsx")
    assert {name: mine.get(name) for name in names} == {name: theirs.get(name) for name in names}
