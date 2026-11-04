"""Range's methods on ranges of several areas, a search of notes and MatchByte, replayed from live Excel
(scripts/measure_range_areas.py, tests/fixtures/range_areas.json): the error, what it answers, what the sheet then reads,
and the same after a save and reopen."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_areas.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[1]}" for number, case in CASES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    app.add_module(RECORD['source'].replace('Application.PathSeparator', f'"{os.sep}"'), name='Areas')
    app.run('Areas.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}

    def call(name: str) -> str:
        try:
            return str(app.run(f'Areas.{name}'))
        except VBAUnsupportedError:
            return 'gap'

    for number in range(len(CASES)):
        found['set'][number] = call(f'S{number}')
    app.run('Areas.Reopen', str(tmp_path_factory.mktemp('areas')))
    for number in range(len(CASES)):
        found['after'][number] = call(f'A{number}')
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_call_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    assert run['set'][number] == case[3]
    assert run['after'][number] == case[4]
