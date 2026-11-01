"""How comments and shapes share names and ids, replayed from live Excel (scripts/measure_comment_shape_names.py,
tests/fixtures/comment_shape_names.json): every shape's name, id and type after each case."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/comment_shape_names.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='NamesProbe')
    return str(app.run('NamesProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
