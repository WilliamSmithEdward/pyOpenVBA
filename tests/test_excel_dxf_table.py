"""A workbook's dxfs as a save writes them: in Excel's order, without those nothing names, every name following.

tests/fixtures/dxf_table/ holds the workbooks scripts/measure_dxf_table.py
had Excel make, a dxf for each kind of name in them; copies with a dxf
nothing names and with two names of one dxf; and Excel's save of each,
opened as it was or after an edit.
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

FIXTURES = Path(__file__).parent / "fixtures" / "dxf_table"
RECORD: dict[str, Any] = json.loads((FIXTURES / "dxf_table.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
#: The cases Excel read, each in one of the two workbooks the measurement made.
READ = [case for case in CASES if case["answer"]]
#: Each edit Excel made to a file it opened, as a macro, and the file it was made to.
EDITS = {
    "made_added": ("made", 'Set f = ActiveWorkbook.Worksheets("three").Range("D1:D3").FormatConditions.Add('
                           'xlCellValue, xlGreater, "5")\nf.Interior.Color = RGB(128, 128, 0)'),
    "order_changed": ("order", 'ActiveWorkbook.Worksheets("alpha").Range("A1").FormatConditions(1).Interior.Color = '
                               'RGB(255, 192, 203)'),
    "order_deleted": ("order", 'ActiveWorkbook.Worksheets("gamma").Range("A1").FormatConditions(1).Delete'),
}
#: Each workbook's rules, its tables' first columns and its sort's colour, as the measurement's readers show them.
READER = '''Public Function Read(tables As Boolean, sorted As String) As String
Dim ws As Object, t As Object, f As Object, out As String
For Each ws In ActiveWorkbook.Worksheets
    out = out & ws.Name & ":"
    For Each f In ws.Cells.FormatConditions
        out = out & f.AppliesTo.Address(False, False) & "@" & f.Priority & "=" & f.Interior.Color & ","
    Next
    If tables Then
        For Each t In ws.ListObjects
            out = out & "table=" & t.ListColumns(1).DataBodyRange.Cells(1, 1).Interior.Color & ","
        Next
    End If
    out = out & ";"
Next
Read = out & "sort=" & ActiveWorkbook.Worksheets(sorted).Sort.SortFields(1).SortOnValue.Color & ";"
End Function

Public Function Pivot() As Long
Pivot = ActiveWorkbook.Worksheets("pivot").PivotTables(1).DataBodyRange.Cells(1, 1).Interior.Color
End Function
'''
#: An element's start tag, and the attributes that name a dxf by its place.
_ELEMENT = re.compile(r"<([\w:]+)\b([^>]*)>")
_NAMES = re.compile(r'\b(dxfId|dataDxfId|headerRowDxfId|totalsRowDxfId|headerRowBorderDxfId|tableBorderDxfId|'
                    r'totalsRowBorderDxfId)="(\d+)"')


def holdings(path: Path) -> dict[str, object]:
    """A file's dxfs, and each part's names of one in its order, as dxf_table.json keeps them: the tag, the
    attribute and the place, and for a rule its range and priority."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        names: dict[str, list[str]] = {}
        for part in sorted(package.namelist()):
            if not part.endswith(".xml") or part.startswith("xl/theme"):
                continue
            text = package.read(part).decode("utf-8")
            found: list[str] = []
            for element in _ELEMENT.finditer(text):
                tag, attributes = element.groups()
                for name in _NAMES.finditer(attributes):
                    extra = ""
                    if tag == "cfRule":
                        priority = re.search(r'\bpriority="(\d+)"', attributes)
                        area = text.rfind("<conditionalFormatting", 0, element.start())
                        sqref = re.search(r'sqref="([^"]*)"', text[area:element.start()]) if area >= 0 else None
                        extra = f":{sqref.group(1) if sqref else ''}@{priority.group(1) if priority else ''}"
                    found.append(f"{tag}.{name.group(1)}={name.group(2)}{extra}")
            if found:
                names[part] = found
    dxfs = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
    listed = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dxfs.group(1) or "", re.DOTALL) if dxfs else []
    return {"dxfs": listed, "names": names}


def excel_saved(name: str) -> dict[str, object]:
    """What Excel's save after the edit ``name`` held."""
    return RECORD["added"] if name == "made_added" else next(case["file"] for case in CASES if case["name"] == name)


def opened(name: str) -> ExcelApplication:
    """The file the edit ``name`` was made to, opened in the model."""
    return ExcelApplication.open(FIXTURES / f"{EDITS[name][0]}.xlsx", with_vba=False)


def edit(app: ExcelApplication, name: str) -> ExcelApplication:
    """``app`` edited as Excel edited a file in the edit ``name``."""
    app.add_module(f"Public Sub Edit()\nDim f As Object\n{EDITS[name][1]}\nEnd Sub\n", name="Editor")
    app.run("Edit")
    return app


def test_the_record_holds_each_file() -> None:
    assert all(holdings(FIXTURES / f"{case['name']}.xlsx") == case["file"]
               and holdings(FIXTURES / f"{case['name']}_resaved.xlsx") == case["resaved"] for case in CASES)
    assert holdings(FIXTURES / "made_added.xlsx") == RECORD["added"]
    assert all(holdings(FIXTURES / f"{name}.xlsx") == held for name, held in RECORD["copies"].items())


def test_a_sheet_copy_and_a_paste_make_dxfs_as_excel_does(tmp_path: Path) -> None:
    """A rule a sheet copy or a paste makes takes a new dxf ahead of every one made before it, as a rule made by
    FormatConditions.Add does."""
    app = ExcelApplication()
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    app.add_module(RECORD["copy_build"].replace("{folder}\\", str(tmp_path) + os.sep), name="Builder")
    assert app.run("Build") == "ok"
    assert {name: holdings(tmp_path / f"{name}.xlsx") for name in RECORD["copies"]} == RECORD["copies"]


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_a_file_saved_unchanged_writes_its_dxfs_as_excel_does(case: dict[str, Any], tmp_path: Path) -> None:
    """Each name of a dxf takes one of its own, in the order Excel reads the parts, a dxf nothing names is left out,
    and every part names each where the save wrote it; a second save keeps what the first wrote."""
    app = ExcelApplication.open(FIXTURES / f"{case['name']}.xlsx", with_vba=False)
    first = holdings(app.save(tmp_path / "first.xlsx"))
    assert first == case["resaved"]
    assert holdings(app.save(tmp_path / "second.xlsx")) == first


@pytest.mark.parametrize("name", EDITS)
def test_an_edit_writes_its_dxfs_as_excel_does(name: str, tmp_path: Path) -> None:
    """A rule made, or one given a new format, takes a new dxf ahead of the file's, and the dxf it had goes; a rule
    a Delete cuts keeps its dxf where it was."""
    assert holdings(edit(opened(name), name).save(tmp_path / f"{name}.xlsx")) == excel_saved(name)


def test_an_edit_after_a_save_writes_its_dxfs_as_excel_does(tmp_path: Path) -> None:
    """Saved and then edited, a file's dxfs stand as the save wrote them, the new rule's first, as Excel's save of
    the edit made straight after opening it holds them."""
    app = opened("made_added")
    assert holdings(app.save(tmp_path / "first.xlsx")) == next(
        case["resaved"] for case in CASES if case["name"] == EDITS["made_added"][0])
    assert holdings(edit(app, "made_added").save(tmp_path / "second.xlsx")) == RECORD["added"]


@pytest.mark.parametrize("case", READ, ids=lambda case: case["name"])
def test_a_save_leaves_each_format_reading_as_excel_read_it(case: dict[str, Any], tmp_path: Path) -> None:
    """The rules, the tables' columns, the sort and the pivot read the same after a save renumbers their dxfs, the
    rules, the tables and the sort as Excel read them in the file."""
    ordered = case["name"].startswith("order")
    app = ExcelApplication.open(FIXTURES / f"{case['name']}.xlsx", with_vba=False)
    app.add_module(READER, name="Reader")
    before = app.run("Read", ordered, "beta" if ordered else "three"), app.run("Pivot")
    app.save(tmp_path / "saved.xlsx")
    assert (app.run("Read", ordered, "beta" if ordered else "three"), app.run("Pivot")) == before
    assert before[0] == case["answer"].split("filter=")[0]
