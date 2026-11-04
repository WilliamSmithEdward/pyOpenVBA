"""Methods of notes, hyperlinks, tables, ranges, sorts and validation, each with no argument and each kind of value, replayed
from live Excel (scripts/measure_method_audit.py, tests/fixtures/method_audit.json): the error, what it answers and what
the object then reads, and the same after a save and reopen. Where binding changes the answer, "early" holds it, and the
model follows that."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/method_audit.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}.{case[1]}" for number, case in CASES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    app.add_module(RECORD['source'].replace('Application.PathSeparator', f'"{os.sep}"'), name='Audit')
    app.run('Audit.Setup')
    found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}

    def call(name: str) -> str:
        try:
            return str(app.run(f'Audit.{name}'))
        except VBAUnsupportedError:
            return 'gap'

    for number in range(len(CASES)):
        found['set'][number] = call(f'S{number}')
    app.run('Audit.Reopen', str(tmp_path_factory.mktemp('audit')))
    for number in range(len(CASES)):
        found['after'][number] = call(f'A{number}')
    return found


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_call_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    early = RECORD['early'].get(str(number))
    if early is not None:
        assert run['set'][number] == early
        return
    assert run['set'][number] == case[3]
    assert run['after'][number] == case[4]
