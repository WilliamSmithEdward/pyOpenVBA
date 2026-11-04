"""The rest of a Range's settable properties, each given each kind of value on the ranges it applies to, replayed from live
Excel (scripts/measure_range_more.py, tests/fixtures/range_more.json): the error and what the range then reads, and the
same after a save and reopen. Where binding changes the answer, "early" holds it, and the model follows that."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_more.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}={case[3][:16]}" for number, case in CASES]


def _plain(answer: str) -> str:
    """A name's formula without its sheet, whose name is the session's."""
    return re.sub(r'=[^|!]*!', '=<sheet>!', answer)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    app.add_module(RECORD['source'].replace('folder & Application.PathSeparator', f'folder & "{os.sep}"'), name='More')
    app.run('More.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}

    def call(name: str) -> str:
        try:
            return str(app.run(f'More.{name}'))
        except VBAUnsupportedError:
            return 'gap'

    for number in range(len(CASES)):
        found['set'][number] = call(f'S{number}')
    app.run('More.Reopen', str(tmp_path_factory.mktemp('more')))
    for number in range(len(CASES)):
        found['after'][number] = call(f'A{number}')
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_value_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    early = RECORD['early'].get(str(number))
    if early is not None:
        assert _plain(run['set'][number]) == _plain(early)
        return
    assert _plain(run['set'][number]) == _plain(case[4])
    assert _plain(run['after'][number]) == _plain(case[5])
