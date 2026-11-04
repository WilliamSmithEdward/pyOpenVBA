"""Range.Value, Value2, Formula, NumberFormat and Style, each given each kind of value on one cell and on two, fresh and
already formatted, replayed from live Excel (scripts/measure_range_values.py, tests/fixtures/range_values.json): the
error and what the first cell then reads, and the same after a save and reopen."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_values.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}={case[3][:20]}" for number, case in CASES]
#: Text of a format of bare digits, where Excel shows 7 under "1.5" as 17.5. Range.Text's row notes it.
BARE_DIGITS = {('NumberFormat', '1.5')}


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    app.add_module(RECORD['source'].replace('folder & Application.PathSeparator', f'folder & "{os.sep}"'), name='Values')
    app.run('Values.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}

    def call(name: str) -> str:
        try:
            return str(app.run(f'Values.{name}'))
        except VBAUnsupportedError:
            return 'gap'

    for number in range(len(CASES)):
        found['set'][number] = call(f'S{number}')
    app.run('Values.Reopen', str(tmp_path_factory.mktemp('values')))
    for number in range(len(CASES)):
        found['after'][number] = call(f'A{number}')
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_value_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    if (case[0], case[3]) in BARE_DIGITS and case[2]:
        pytest.skip('a format of bare digits, which Range.Text does not render yet')
    assert run['set'][number] == RECORD['early'].get(str(number), case[4])
    assert run['after'][number] == case[5]
