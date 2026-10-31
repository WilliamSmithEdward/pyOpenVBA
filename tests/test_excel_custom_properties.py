"""Worksheet.CustomProperties, replayed against live Excel (tests/fixtures/custom_properties.json,
scripts/measure_custom_properties.py)."""
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

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/custom_properties.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: A name of 300 characters, which Excel refuses with error 7 at a length the model does not know.
UNSUPPORTED = {23}


def _parts(path: Path) -> dict[str, str]:
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if re.fullmatch(r"xl/worksheets/sheet[0-9]+\.xml", name):
                xml = package.read(name).decode("utf-8")
                found[name] = (re.search(r"<customProperties>.*?</customProperties>", xml, re.DOTALL) or [""])[0]
            elif re.fullmatch(r"xl/worksheets/_rels/sheet[0-9]+\.xml\.rels", name):
                found[name] = package.read(name).decode("utf-8")
            elif name.startswith("xl/customProperty"):
                found[name] = package.read(name).hex()
        found["defaults"] = " ".join(re.findall(r"<Default [^>]*/>", package.read("[Content_Types].xml").decode()))
    return found


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_custom_properties_answer_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    if index in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", str(tmp_path), index, index)
        return
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, statements
    if saved in FIXTURE["parts"]:
        expected = {key: value for key, value in FIXTURE["parts"][saved].items() if not key.endswith(":root")}
        expected["defaults"] = " ".join(re.findall(r"<Default [^>]*/>", expected.pop("[Content_Types].xml")))
        assert _parts(tmp_path / saved) == expected, statements
