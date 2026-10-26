"""Workbook.SaveAs to CSV and text formats, replayed from live Excel (scripts/measure_text_saves.py,
tests/fixtures/text_saves.json): the bytes each format writes of the active sheet, and what the workbook is after
it -- its name, FileFormat, Saved and sheets -- and after a later Save."""

from __future__ import annotations

import base64
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / "fixtures" / "text_saves.json").read_text(encoding="utf-8"))
CASES = {name: (file, file_format) for name, file, file_format in RECORD["cases"]}
#: The formats the model does not write, which report themselves unsupported: the Mac ones and the printer's.
UNSUPPORTED = {"csv_mac": "xlCSVMac", "text_mac": "xlTextMac", "printer": "xlTextPrinter"}


def _saved(name: str, folder: Path) -> tuple[str, bytes]:
    file, file_format = CASES[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")
    answer = str(app.run("Probe.Probe", name, str(folder / file), file_format))
    written = sorted(folder.glob(file + "*"))
    assert len(written) == 1, written
    return answer, written[0].read_bytes()


@pytest.mark.parametrize("name", [name for name in CASES if name not in UNSUPPORTED])
def test_each_save_is_excels(name: str, tmp_path: Path) -> None:
    answer, data = _saved(name, tmp_path)
    assert answer == RECORD["answers"][name]
    assert data == base64.b64decode(RECORD["files"][name])


@pytest.mark.parametrize("name", list(UNSUPPORTED))
def test_a_format_the_model_does_not_write_reports_itself(name: str, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match=UNSUPPORTED[name]):
        _saved(name, tmp_path)
