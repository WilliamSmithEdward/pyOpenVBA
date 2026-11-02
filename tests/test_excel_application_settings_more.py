"""More Application settings read as Excel starts, replayed from live Excel
(scripts/measure_application_settings_more.py, tests/fixtures/application_settings_more.json). The read-only flags'
setters are left out of the replay: the model answers their reads only."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/application_settings_more.json').read_text(encoding='utf-8'))
READ = ('Public Function Read(ByVal name As String) As String\n'
        '    Read = CStr(CallByName(Application, name, VbGet))\nEnd Function\n')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(READ, name='SettingsRead')
    return made


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[name for name, _ in RECORD['cases']])
def test_each_setting_starts_as_excels_did(app: ExcelApplication, index: int) -> None:
    name, answer = RECORD['cases'][index]
    start = answer.split('|')[0].removeprefix('start=')
    if start.startswith('E'):
        pytest.skip(f'Excel raised {start} reading it')
    assert str(app.run('SettingsRead.Read', name)) == start
