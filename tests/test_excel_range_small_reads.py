"""Small Range reads replayed from live Excel (scripts/measure_range_small_reads.py,
tests/fixtures/range_small_reads.json): HasRichDataType, LinkedDataTypeState, SavedAsArray, Summary, MDX,
CommentThreaded, QueryTable and ServerActions on plain cells, a spill and an array formula."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_small_reads.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='SmallReads')
    return str(app.run('SmallReads.Probe')).split('~')


@pytest.mark.parametrize('index', range(len(RECORD['reads'])),
                         ids=[f'{target}.{name}' for target, name, _ in RECORD['reads']])
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['reads'][index][2]
