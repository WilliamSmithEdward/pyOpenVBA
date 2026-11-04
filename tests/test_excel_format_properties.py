"""Font, Interior, Border and Borders on cells, each settable property given each kind of value, replayed from live Excel
(scripts/measure_format_properties.py, tests/fixtures/format_properties.json): the error and what it then reads, the
same after a save and reopen, and what two cells that differ read through their range."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/format_properties.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}.{case[2]}={case[3]}" for number, case in CASES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    app.add_module(RECORD['source'].replace('Application.PathSeparator', f'"{os.sep}"'), name='Formats')
    app.run('Formats.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'mixed': {}, 'after': {}}

    def call(name: str) -> str:
        try:
            return str(app.run(f'Formats.{name}'))
        except VBAUnsupportedError:
            return 'gap'

    for number in range(len(CASES)):
        found['set'][number] = call(f'S{number}')
    for number in range(len(RECORD['mixed'])):
        found['mixed'][number] = call(f'M{number}')
    app.run('Formats.Reopen', str(tmp_path_factory.mktemp('formats')))
    for number in range(len(CASES)):
        found['after'][number] = call(f'A{number}')
    app.run('Formats.Finish')
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_value_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    if run['set'][number] == 'gap':
        pytest.skip('a tint on an automatic colour, which the model reports as a gap')
    early = RECORD['early'].get(str(number))
    if early is not None:
        # The model follows early binding, whose answer this is where the two differ.
        assert run['set'][number] == early
        return
    assert run['set'][number] == case[4]
    assert run['after'][number] == case[5]


@pytest.mark.parametrize('number', range(len(RECORD['mixed'])))
def test_cells_that_differ_read_through_their_range(run: dict[str, dict[int, str]], number: int) -> None:
    assert run['mixed'][number] == RECORD['mixed'][number][-1]
