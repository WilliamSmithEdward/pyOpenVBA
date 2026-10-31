"""Application.FindFormat and ReplaceFormat replayed from live Excel (scripts/measure_find_format.py,
tests/fixtures/find_format.json): CellFormat's reads, Range.Find with SearchFormat and the finds after it, and
Range.Replace with SearchFormat and ReplaceFormat."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/find_format.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='FormatProbe')
    return str(app.run('FormatProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][3]
