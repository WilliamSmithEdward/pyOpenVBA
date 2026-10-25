"""Date fields in a pivot table and its date filters, replayed against the in-memory model on the measured day.

tests/fixtures/pivot_date_filters/ holds what
scripts/measure_pivot_date_filters.py saw in live Excel on the day the
record names: a pivot table on a sheet of its own for each layout, over
one list of dates placed about that day, regions and amounts, with a date
field on its rows, columns or page and a date filter added to its
PivotFilters, by a date, two dates or a period; source columns in one
format or several, and caches refreshed; each table dumped as its cells,
the first date cell's type and format, the value and label cells'
formats, each field's items, format and filters, and each data field's
format; the workbook Excel saved of them all; and each table as Excel read
it again from that file. The model's clock is pinned to that day.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_pivot_layouts import column_widths

FIXTURES = Path(__file__).parent / "fixtures" / "pivot_date_filters"
RECORD: dict[str, Any] = json.loads((FIXTURES / "pivot_date_filters.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
HELPER = RECORD["build"].split("Public Function Build")[0].replace("{steps}", RECORD["steps"])
#: The day the record was measured, which the model's clock is pinned to: VBA's Date and the periods counted from
#: today read it.
TODAY = dt.datetime.fromisoformat(RECORD["today"] + "T09:00:00")


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's pivotFields, rowItems, colItems and filters as a file keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8").replace(" />", "/>")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            kept = [found.group() for tag in ("pivotFields", "rowItems", "colItems")
                    if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL))]
            filters = re.search(r'<filters count="\d+">.*?</filter></filters>', text, re.DOTALL)
            out[name.group(1) if name else part] = "".join(kept) + (filters.group() if filters else "")
    return out


def cached_dates(cache: str) -> tuple[str, list[str], str]:
    """A cache's Date field as it bears on the pivot table: its format, its dates, and its first and the midnight
    after its last."""
    field = re.search(r'<cacheField name="Date" numFmtId="(\d+)"[^>]*>(.*?)</cacheField>', cache, re.DOTALL)
    if field is None:
        return "", [], ""
    bounds = " ".join(re.findall(r'(?:minDate|maxDate)="[^"]*"', field.group(2)))
    return field.group(1), re.findall(r'<d v="([^"]*)"', field.group(2)), bounds


def dumps(text: str) -> dict[str, str]:
    return dict(one.split("^", 1) for one in text.split("|") if "^" in one)


def built(folder: Path) -> ExcelApplication:
    """The model with the macro Excel ran in it and its clock on the measured day."""
    body = "".join(RECORD["layout"].replace("{name}", layout["name"]).replace("{row}", layout["row"])
                   .replace("{column}", layout["column"]).replace("{page}", layout["page"])
                   .replace("{source}", layout["own_source"]).replace("{step}", str(index))
                   for index, layout in enumerate(LAYOUTS, 1))
    # The macro Excel ran names its file under a Windows folder; here it goes under this system's.
    code = (RECORD["build"].replace("{steps}", RECORD["steps"]).replace("{source}", RECORD["source"])
            .replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep))
    app = ExcelApplication()
    app.freeze_clock(TODAY)
    app.add_module(code, name="Builder")
    return app


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[str, str], dict[str, str], Path]:
    """Every layout made in the model, its dumps in the session and once the model opened its own file again, and
    the folder it saved to."""
    folder = tmp_path_factory.mktemp("pivot_date_filters")
    today, rest = str(built(folder).run("Build")).split("@", 1)
    assert today == RECORD["today"]
    session, reopened = rest.split("~", 1)
    return dumps(session), dumps(reopened), folder


def test_the_record_holds_the_file() -> None:
    assert tables(FIXTURES / "pivot_date_filters.xlsx") == {layout["name"]: layout["file"] for layout in LAYOUTS}


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_date_field_shows_what_excel_shows(made: tuple[dict[str, str], dict[str, str], Path],
                                             layout: dict[str, Any]) -> None:
    """The cells, the date cells' type and format, the value and label cells' formats, each field's items, format
    and filters, and each data field's format: dates shown in their field's format, or m/d/yyyy where the source's
    cells differ, a date filter keeping the dates it names, a period counted from the measured day, and the errors
    Add raises."""
    assert made[0][layout["name"]] == layout["dump"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_a_date_field_reads_back_from_the_file_the_model_saved(
        made: tuple[dict[str, str], dict[str, str], Path], layout: dict[str, Any]) -> None:
    assert made[1][layout["name"]] == layout["reopened"]


@pytest.mark.parametrize("layout", LAYOUTS, ids=lambda layout: layout["name"])
def test_the_file_keeps_each_date_filter_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                        layout: dict[str, Any]) -> None:
    """Each field's format, the items and lines, and each date filter as the autoFilter that spells it, a period
    with the days it kept and a whole-day filter in Excel 2013's extension."""
    assert tables(made[2] / "pivot_date_filters.xlsx")[layout["name"]] == layout["file"]


@pytest.mark.parametrize("layout", [one for one in LAYOUTS if one["row"] == "Date" or one["column"] == "Date"],
                         ids=lambda layout: layout["name"])
def test_the_cache_keeps_the_dates_as_excels_does(made: tuple[dict[str, str], dict[str, str], Path],
                                                  layout: dict[str, Any]) -> None:
    """A date field's format, its dates spelled to the second, a refresh dropping a date the source lost, and the
    midnight after the last date as maxDate."""
    with zipfile.ZipFile(made[2] / "pivot_date_filters.xlsx") as package:
        parts = {name: package.read(name).decode("utf-8") for name in package.namelist()}
    table = next(text for name, text in parts.items() if name.startswith("xl/pivotTables/pivotTable")
                 and f'name="{layout["name"]}"' in text)
    part = next(name for name in parts if name.startswith("xl/pivotTables/pivotTable") and parts[name] == table)
    rels = parts[part.replace("pivotTables/", "pivotTables/_rels/") + ".rels"]
    target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
    assert target is not None
    assert cached_dates(parts[f"xl/pivotCache/{target.group(1)}"]) == cached_dates(layout["cache"])


@pytest.fixture(scope="module")
def excel_file_read() -> dict[str, str]:
    """Every table of the file Excel saved, as the model reads it on the measured day."""
    app = ExcelApplication.open(FIXTURES / "pivot_date_filters.xlsx", with_vba=False)
    app.freeze_clock(TODAY)
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
    assert column_widths(made[2] / "pivot_date_filters.xlsx", names) == column_widths(FIXTURES / "pivot_date_filters.xlsx", names)
