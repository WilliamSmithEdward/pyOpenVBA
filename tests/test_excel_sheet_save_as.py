"""Worksheet.SaveAs, replayed from live Excel (scripts/measure_sheet_save_as.py, tests/fixtures/sheet_save_as.json):
a text format writes the sheet whose SaveAs runs, which takes the file's name while the active sheet stays, and a
workbook format saves the workbook."""

from __future__ import annotations

import base64
import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "sheet_save_as.json").read_text(encoding="utf-8"))
CASES = {name: file for name, file, _ in RECORD["cases"]}


@pytest.mark.parametrize("name", list(CASES))
def test_each_save_is_excels(name: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")
    assert str(app.run("Probe.Probe", name, str(tmp_path / CASES[name]))) == RECORD["answers"][name]
    found = sorted(tmp_path.glob(CASES[name] + "*"))
    expected = RECORD["files"][name]
    assert len(found) == 1 and found[0].name == expected["file"]
    if "sheets" in expected:
        with zipfile.ZipFile(found[0]) as package:
            workbook = package.read("xl/workbook.xml").decode("utf-8")
        assert re.findall(r'<sheet\b[^>]*name="([^"]+)"', workbook) == expected["sheets"]
    else:
        assert found[0].read_bytes() == base64.b64decode(expected["data"])
