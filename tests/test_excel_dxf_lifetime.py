"""Which dxfs a save keeps once nothing names them, as Excel's does.

tests/fixtures/dxf_lifetime/ holds the workbooks
scripts/measure_dxf_lifetime.py had Excel make: rules made, one taken
away by FormatCondition.Delete, the collection's Delete, ClearFormats,
Clear, a paste, its rows or its sheet deleted, after a new colour, before
a new rule, or in a workbook saved and opened again; and Excel's save.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

sys.path.insert(0, str(Path(__file__).parent))
from test_excel_dxf_table import holdings  # noqa: E402

FIXTURES = Path(__file__).parent / "fixtures" / "dxf_lifetime"
RECORD: dict[str, Any] = json.loads((FIXTURES / "dxf_lifetime.json").read_text(encoding="utf-8"))
FILES: dict[str, Any] = RECORD["files"]


@pytest.fixture(scope="module")
def made(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """The folder the model saved every layout to, made by the macro Excel ran."""
    folder = tmp_path_factory.mktemp("dxf_lifetime")
    saves = {layout["name"]: "".join(f'wb.SaveAs "{{folder}}\\{save}.xlsx", 51\n' for save in layout["saves"])
             for layout in RECORD["layouts"]}
    body = "".join(RECORD["layout"].replace("{steps}", layout["steps"]).replace("{name}", layout["name"])
                   .replace("{saves}", saves[layout["name"]]) for layout in RECORD["layouts"])
    # The macro Excel ran names its files under a Windows folder; here they go under this system's.
    code = RECORD["build"].replace("{layouts}", body).replace("{folder}\\", str(folder) + os.sep)
    app = ExcelApplication()
    app.add_module(code, name="Builder")
    assert app.run("Build") == "ok"
    return folder


def test_the_record_holds_each_file() -> None:
    assert all(holdings(FIXTURES / f"{name}.xlsx") == held for name, held in FILES.items())


@pytest.mark.parametrize("name", FILES)
def test_a_save_keeps_the_dxfs_excel_keeps(made: Path, name: str) -> None:
    """A rule a Delete, a Clear or a paste took away leaves its dxf where it was, through every save until the
    workbook is opened again; a rule that went with its rows or its sheet, or a format a new one replaced, takes its
    dxf with it."""
    assert holdings(made / f"{name}.xlsx") == FILES[name]
