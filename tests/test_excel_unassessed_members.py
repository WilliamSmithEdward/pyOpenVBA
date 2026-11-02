"""Members the model registered with no measurement behind them, replayed from live Excel
(scripts/measure_unassessed_members.py, tests/fixtures/unassessed_members.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/unassessed_members.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='UnassessedProbe')
    return str(app.run('UnassessedProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[read for read, _ in RECORD['cases']])
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
