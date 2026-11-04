"""The worksheet functions macros call most, each with each kind of value for each argument, on a WorksheetFunction typed and
on one held As Object, replayed from live Excel (scripts/measure_worksheet_function_args.py,
tests/fixtures/worksheet_function_args.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/worksheet_function_args.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}" for number, case in CASES]


def _answers(source: str) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Functions')
    return str(app.run('Functions.Probe')).split('@@')


@pytest.fixture(scope='module')
def late() -> list[str]:
    return _answers(RECORD['late_source'])


@pytest.fixture(scope='module')
def typed() -> list[str]:
    return _answers(RECORD['typed_source'])


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_call_late_bound(late: list[str], number: int, case: list[str]) -> None:
    assert late[number] == case[2], case[1]


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_call_typed(typed: list[str], number: int, case: list[str]) -> None:
    assert typed[number] == case[3], case[1]
