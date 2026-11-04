"""The local formula and format properties set on a block, and a Range's properties read over cells that differ, replayed
from live Excel (scripts/measure_range_more_blocks.py, tests/fixtures/range_more_blocks.json)."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_more_blocks.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}={case[3][:16]}" for number, case in CASES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'].replace('folder & Application.PathSeparator', f'folder & "{os.sep}"'), name='More')
    app.run('More.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}
    for number in range(len(CASES)):
        found['set'][number] = str(app.run(f'More.S{number}'))
    app.run('More.Reopen', str(tmp_path_factory.mktemp('blocks')))
    for number in range(len(CASES)):
        found['after'][number] = str(app.run(f'More.A{number}'))
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_value_on_a_block_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    assert run['set'][number] == case[4]
    assert run['after'][number] == case[5]


def test_each_read_over_cells_that_differ_as_excel_answered() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['differing_source'], name='Differing')
    found = str(app.run('Differing.Probe')).split('@@')
    assert [[case[0], case[2], answer] for case, answer in zip(RECORD['differing'], found, strict=True)] == [
        [case[0], case[2], case[4]] for case in RECORD['differing']]
    # Binding changes none of these.
    assert all(case[3] == case[4] for case in RECORD['differing'])
