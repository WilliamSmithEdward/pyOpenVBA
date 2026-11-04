"""Six small members, replayed from live Excel (scripts/measure_small_rows.py, tests/fixtures/small_rows.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/small_rows.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]


def _app() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='SmallProbe')
    return app


@pytest.mark.parametrize('index', range(len(NAMES)), ids=NAMES)
def test_each_case_as_excel_answered(index: int) -> None:
    assert str(_app().run('SmallProbe.Probe', index)) == RECORD['cases'][index][1]
