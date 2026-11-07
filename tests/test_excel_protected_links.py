"""Links on protected sheets, and ranges a file protects with the older 16-bit password, as live Excel answered
(tests/fixtures/protected_links.json, tests/fixtures/protected_links/).

Each round's calls were new, with the model's link answers written down first; the password books are Excel's own with their
hash swapped for the older attribute, so that side ran after Excel. Two things are not here because Excel stops for a person:
Unprotect of such a range with no password, and ClearContents over one, each put up a prompt for the password.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

HERE = Path(__file__).parent / "fixtures"
RECORD: dict[str, Any] = json.loads((HERE / "protected_links.json").read_text(encoding="utf-8"))
READ = 's = s & "|" & ws.Hyperlinks.Count & "|" & ws.Range("A1").Font.Underline & "|" & ws.Range("B2").Font.Underline & "|" & ws.Range("A1").Value'


def _run(source: str) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name="M")
    return str(app.run("M.Probe"))


def _ranges(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        part = min(name for name in archive.namelist() if re.fullmatch(r"xl/worksheets/[^/]+\.xml", name))
        text = archive.read(part).decode("utf-8")
    found = re.search(r"<protectedRanges>.*?</protectedRanges>", text, re.S)
    return re.sub(r'(hashValue|saltValue)="[^"]*"', r'\1=""', found.group()) if found else ""


@pytest.mark.parametrize("name", sorted(RECORD["rounds"]))
def test_deleting_links_as_excel_did(name: str) -> None:
    wrong: list[tuple[Any, ...]] = []
    for case in RECORD["rounds"][name]:
        if case["kind"] != "link":
            continue
        source = ["Public Function Probe() As String", 'Dim wb, ws, s, h, l: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Sheet1": s = ""', *case["setup"],
                  "On Error Resume Next", "Err.Clear", case["act"], "s = s & Err.Number", READ, "wb.Close False", "Probe = s", "End Function", ""]
        got = _run("\n".join(source))
        if got != case["excel"]:
            wrong.append((case["setup"], case["act"], case["excel"], got))
    assert wrong == []


@pytest.mark.parametrize("name", sorted(RECORD["rounds"]))
def test_older_passwords_as_excel_read_them(name: str, tmp_path: Path) -> None:
    wrong: list[tuple[Any, ...]] = []
    save = tmp_path / "saved.xlsx"
    for case in RECORD["rounds"][name]:
        if case["kind"] != "password":
            continue
        save.unlink(missing_ok=True)
        path = HERE / "protected_links" / case["book"]
        source = ["Public Function Probe() As String", f'Dim wb, ws, s: Application.DisplayAlerts = False: Set wb = Workbooks.Open("{path}"): Set ws = wb.Worksheets(1): s = ""', "On Error Resume Next", "Err.Clear", case["act"],
                  's = s & "|" & Err.Number & "|" & ws.Range("A1").Value & "|" & ws.Range("C1").Value', f'wb.SaveAs "{save}", 51', "wb.Close False", "Probe = s", "End Function", ""]
        got = _run("\n".join(source)) + "~" + (_ranges(save) if save.exists() else "")
        if got != case["excel"]:
            wrong.append((case["act"], case["excel"], got))
    assert wrong == []


def test_adding_a_link_over_each_kind_of_cell() -> None:
    lines = ["Public Function Probe() As String", 'Dim wb, ws, c, s: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "Sheet1": s = ""', "On Error Resume Next"]
    for setup, call in RECORD["add"]["calls"]:
        lines += ["Err.Clear", 'Set c = ws.Range("A1"): c.Clear', setup, call, 's = s & Err.Number & "|" & TypeName(c.Value) & "|" & CStr(c.Formula) & "|" & c.Text & "|" & c.Font.Underline & "|" & ws.Hyperlinks.Count & ";"', "ws.Hyperlinks.Delete"]
    lines += ["wb.Close False", "Probe = s", "End Function", ""]
    assert _run("\n".join(lines)).split(";")[:-1] == RECORD["add"]["answers"]


def test_the_scores_of_the_last_rounds() -> None:
    assert [RECORD["scores"][name]["differ"] for name in sorted(RECORD["rounds"])] == [0] * len(RECORD["rounds"])
