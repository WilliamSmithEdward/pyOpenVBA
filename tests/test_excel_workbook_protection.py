"""Workbook protection, replayed against the in-memory model.

tests/fixtures/workbook_protection.json is what
scripts/measure_workbook_protection.py saw in live Excel: each case
protects a two-sheet workbook some way and tries a change, recording the
error and the state left behind; each file configuration was saved, its
workbookProtection element recorded, and the file read back.
"""

from __future__ import annotations

import base64
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._protection import spun
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "workbook_protection.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
FILES: list[dict[str, Any]] = RECORD["files"]

FRESH = 'Set wb = ActiveWorkbook\nwb.Worksheets(1).Name = "Alpha"\nwb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Beta"'


def _run(case: dict[str, Any]) -> str:
    lines = ["Public Function Probe() As String", "Dim out As String, wb As Object", FRESH,
             f'out = out & "{case["name"]}="', "On Error Resume Next", "Err.Clear"]
    if case["setup"]:
        lines += [case["setup"], 'If Err.Number <> 0 Then out = out & "setup " & Err.Number & ";"', "Err.Clear"]
    lines += [case["action"], 'out = out & "E" & Err.Number & ":" & Err.Description & ";"', "Err.Clear",
              f"out = out & {RECORD['state']}", "Probe = out", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(lines) + "\n", name="Probe")
    return str(app.run("Probe")).split("=", 1)[1]


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_workbook_protection_answers_as_excels_does(case: dict[str, Any]) -> None:
    assert _run(case) == case["answer"]


@pytest.mark.parametrize("action", ["wb.Unprotect", "wb.Protect"])
def test_a_password_excel_would_ask_for_says_so(action: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Probe()\nDim wb As Object\nSet wb = ActiveWorkbook\nwb.Protect "pw"\n{action}\nEnd Sub\n',
                   name="Probe")
    with pytest.raises(VBAUnsupportedError, match="dialog"):
        app.run("Probe")


def _saved(path: Path, action: str) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Public Sub Make()\nDim wb As Object\n{FRESH}\n{action}\nEnd Sub\n", name="Maker")
    app.run("Make")
    app.save(path)
    return path


def _workbook_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("xl/workbook.xml").decode("utf-8")


def _attributes(element: str) -> dict[str, str]:
    return dict(re.findall(r'(\w+)="([^"]*)"', element))


@pytest.mark.parametrize("case", FILES, ids=[case["name"] for case in FILES])
def test_the_model_writes_the_element_excel_writes(tmp_path: Path, case: dict[str, Any]) -> None:
    xml = _workbook_xml(_saved(tmp_path / "book.xlsx", case["action"]))
    written = re.search(r"<workbookProtection\b[^>]*/>", xml)
    if not case["element"]:
        assert written is None
        return
    assert written is not None
    # In its place: before bookViews, as Excel puts it.
    assert re.match(r"\s*<bookViews\b", xml[written.end():])
    if case["name"] != "password":
        assert written.group(0) == case["element"]
        return
    ours, excels = _attributes(written.group(0)), _attributes(case["element"])
    assert list(ours) == list(excels)
    assert ours["lockStructure"] == excels["lockStructure"] and ours["workbookSpinCount"] == excels["workbookSpinCount"]
    assert spun("pw", base64.b64decode(ours["workbookSaltValue"]), int(ours["workbookSpinCount"])) == \
        ours["workbookHashValue"]


def test_the_hash_is_the_one_excel_wrote() -> None:
    attributes = _attributes(next(case for case in FILES if case["name"] == "password")["element"])
    assert spun("pw", base64.b64decode(attributes["workbookSaltValue"]), int(attributes["workbookSpinCount"]),
                attributes["workbookAlgorithmName"]) == attributes["workbookHashValue"]


def _with_element(source: Path, target: Path, element: str) -> Path:
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as copy:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text = re.sub(r"<workbookProtection\b[^>]*/>", "", data.decode("utf-8"))
                data = text.replace("<bookViews", element + "<bookViews", 1).encode("utf-8")
            copy.writestr(item, data)
    return target


def _read(path: Path, code: str) -> str:
    app = ExcelApplication.open(path, with_vba=False)
    app.add_module(RECORD["helper"] + "Public Function Read() As String\nDim wb As Object\nSet wb = ActiveWorkbook\n"
                   f"On Error Resume Next\n{code}\nEnd Function\n", name="Reader")
    return str(app.run("Read"))


@pytest.mark.parametrize("case", FILES, ids=[case["name"] for case in FILES])
def test_excels_element_reads_as_excel_read_it(tmp_path: Path, case: dict[str, Any]) -> None:
    template = _saved(tmp_path / "template.xlsx", "")
    path = _with_element(template, tmp_path / "excels.xlsx", case["element"])
    assert _read(path, f"Read = {RECORD['state']}") == case["read"]


def test_a_protected_workbook_the_model_saved_is_protected_when_read_again(tmp_path: Path) -> None:
    path = _saved(tmp_path / "round.xlsx", 'wb.Protect "pw"')
    code = ('Err.Clear\nwb.Worksheets.Add\nRead = Err.Number\nErr.Clear\nwb.Unprotect "bad"\nRead = Read & "," & '
            'Err.Number\nErr.Clear\nwb.Unprotect "pw"\nRead = Read & "," & Err.Number & "," & wb.ProtectStructure')
    assert _read(path, code) == "1004,1004,0,False"
