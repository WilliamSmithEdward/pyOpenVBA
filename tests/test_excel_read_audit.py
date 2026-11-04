"""Read-only members read in each state that can change them, replayed from live Excel (scripts/measure_read_audit.py,
tests/fixtures/read_audit.json), and again after a save and reopen where the file keeps what they read."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/read_audit.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}.{case[1]}" for number, case in CASES]
#: The application's and a window's answers are the session's, not the file's.
NOT_REOPENED = {'Application', 'Window'}
#: Measured differences the checklist keeps PARTIAL. Each read of ActiveWindow hands back a new object in Excel, so `Is`
#: between two is False. A loop's CircularReference is found by calculating, which the model does not do on opening.
KNOWN = {('Application', 'ActiveWindow'), ('Worksheet', 'CircularReference')}


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[int, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # The probe as recorded, its file joined to the folder with this platform's separator.
    source = RECORD['source'].replace('folder & Application.PathSeparator', f'folder & "{os.sep}"')
    app.add_module(source, name='Audit')
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
def test_each_read_as_excel_answered(run: dict[str, dict[int, str]], number: int, case: list[str]) -> None:
    if (case[0], case[1]) in KNOWN:
        pytest.skip('a measured difference its row notes')
    assert run['set'][number] == case[3]
    if case[0] not in NOT_REOPENED:
        assert run['after'][number] == case[4]
