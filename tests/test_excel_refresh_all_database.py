"""Workbook.RefreshAll over query tables on a database, a sheet's and a ListObject's, against what Excel did with an Access file through
its own ACE OLEDB provider (scripts/measure_refresh_all_database.py, tests/fixtures/refresh_all_database.json).

The function here stands for the Access file: table T of six numbered rows, chosen by the command's WHERE.
"""

from __future__ import annotations

import importlib.util
import json
import re
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest

HERE = Path(__file__).parent / "fixtures"
MEASURED: dict[str, Any] = json.loads((HERE / "refresh_all_database.json").read_text(encoding="utf-8"))
FILE = "C:\\data\\rows.accdb"


def _script(name: str) -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location(name, folder / f"{name}.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


MEASURE = _script("measure_refresh_all_database")


def numbered(request: HookRequest) -> list[list[object]]:
    assert request.source is not None
    found = re.fullmatch(r"SELECT \* FROM T WHERE Id <= (\d+) ORDER BY Id", request.source.command_text)
    assert found is not None, request.source.command_text
    return [["Id", "Name"], *[[float(at), f"n{at}"] for at in range(1, int(found.group(1)) + 1)]]


@pytest.mark.parametrize("case", sorted(MEASURED["rows"]))
def test_refresh_all_over_a_database_does_what_excel_did(case: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.set_hook("Refresh", numbered)
    code = MEASURE.CASE.replace("{provider}", MEASURED["provider"]).replace(
        "{steps}", "\n".join(f"    {step}" for step in MEASURED["cases"][case]))
    app.add_module(code, name="M")
    assert str(app.run("Go", FILE)).split("\t") == MEASURED["rows"][case]
