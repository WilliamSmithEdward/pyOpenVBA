"""The 3-D view, style and label level reads for each chart type the model makes, and the line objects of a chart's
first group, replayed from live Excel (scripts/measure_chart_type_reads.py, tests/fixtures/chart_type_reads.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_type_reads.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='TypeReadsProbe')
    return str(app.run('TypeReadsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case for case, _ in RECORD['cases']])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
