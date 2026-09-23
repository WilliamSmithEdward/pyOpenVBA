"""Sheet protection in the file, against what Excel saved and read back.

tests/fixtures/protection_file.json is what scripts/measure_protection_file.py
saw: the sheetProtection element Excel wrote for each way of protecting a
sheet, what Excel read back from each file, and which passwords opened
sheets carrying the legacy 16-bit hash. The model writes the same
elements, reads Excel's the way Excel read them, and checks both hashes.
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
from pyopenvba.apps.excel._protection import legacy_hash, spun

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "protection_file.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]
#: The passwords the measured cases protected with.
PASSWORDS = {"password": "pw", "password_long": "Correct Horse 9!", "allow_everything_select_none": "pw"}


def _sheet_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _attributes(element: str) -> dict[str, str]:
    return dict(re.findall(r'(\w+)="([^"]*)"', element))


def _saved(tmp_path: Path, action: str) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Public Sub Make()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n{action}\nEnd Sub\n",
                   name="Probe")
    app.run("Make")
    path = tmp_path / "saved.xlsx"
    app.save(path)
    return path


def _with_element(source: Path, target: Path, element: str) -> Path:
    """``source`` with its first sheet's protection made ``element``, which may be none."""
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as copy:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = re.sub(r"<sheetProtection\b[^>]*/>", "", data.decode("utf-8"))
                text = text.replace("</sheetData>", "</sheetData>" + element, 1)
                data = text.encode("utf-8")
            copy.writestr(item, data)
    return target


def _read(path: Path, code: str) -> str:
    app = ExcelApplication.open(path, with_vba=False)
    app.add_module(f"Public Function Read() As String\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n"
                   f"On Error Resume Next\n{code}\nEnd Function\n", name="Reader")
    return str(app.run("Read"))


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_the_model_writes_the_element_excel_writes(tmp_path: Path, case: dict[str, Any]) -> None:
    written = re.search(r"<sheetProtection\b[^>]*/>", _sheet_xml(_saved(tmp_path, case["action"])))
    if not case["element"]:
        assert written is None
        return
    assert written is not None
    if case["name"] not in PASSWORDS:
        assert written.group(0) == case["element"]
        return
    ours, excels = _attributes(written.group(0)), _attributes(case["element"])
    # The salt is random, Excel's and ours: every other attribute, in its order, matches.
    assert [key for key in ours] == [key for key in excels]
    assert {key: value for key, value in ours.items() if key not in ("hashValue", "saltValue")} == \
        {key: value for key, value in excels.items() if key not in ("hashValue", "saltValue")}
    assert spun(PASSWORDS[case["name"]], base64.b64decode(ours["saltValue"]), int(ours["spinCount"])) == \
        ours["hashValue"]


@pytest.mark.parametrize("name", sorted(PASSWORDS))
def test_the_hash_is_the_one_excel_wrote(name: str) -> None:
    case = next(one for one in CASES if one["name"] == name)
    attributes = _attributes(case["element"])
    assert spun(PASSWORDS[name], base64.b64decode(attributes["saltValue"]), int(attributes["spinCount"]),
                attributes["algorithmName"]) == attributes["hashValue"]


@pytest.fixture(scope="module")
def template(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return _saved(tmp_path_factory.mktemp("template"), "ws.Range(\"B2\").Value = 1")


@pytest.mark.parametrize("case", CASES, ids=[case["name"] for case in CASES])
def test_excels_element_reads_as_excel_read_it(template: Path, tmp_path: Path, case: dict[str, Any]) -> None:
    path = _with_element(template, tmp_path / "excels.xlsx", case["element"])
    assert _read(path, f"Read = {RECORD['read']}") == case["read"]


@pytest.mark.parametrize("case", RECORD["legacy"], ids=[case["name"] for case in RECORD["legacy"]])
def test_a_legacy_password_opens_as_it_did_in_excel(template: Path, tmp_path: Path, case: dict[str, Any]) -> None:
    assert legacy_hash(case["password"]) == case["hash"]
    element = f'<sheetProtection password="{case["hash"]}" sheet="1" objects="1" scenarios="1"/>'
    path = _with_element(template, tmp_path / "legacy.xlsx", element)
    code = f'Err.Clear\nws.Unprotect "{case["attempt"]}"\nRead = Err.Number & "," & ws.ProtectContents'
    assert _read(path, code) == case["answer"]


def test_a_protected_sheet_the_model_saved_is_protected_when_read_again(tmp_path: Path) -> None:
    """The round trip: Protect with a password, save, open, and the password still opens it."""
    path = _saved(tmp_path, 'ws.Protect "pw", AllowSorting:=True')
    code = ('Err.Clear\nws.Range("A1").Value = 1\nRead = Err.Number & "," & ws.Protection.AllowSorting\n'
            'Err.Clear\nws.Unprotect "bad"\nRead = Read & "," & Err.Number\nErr.Clear\nws.Unprotect "pw"\n'
            'Read = Read & "," & Err.Number & "," & ws.ProtectContents')
    assert _read(path, code) == "1004,True,1004,0,False"
