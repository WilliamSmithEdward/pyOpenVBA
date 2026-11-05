"""Range's most used methods, each with no argument and each kind of value for each argument, replayed from live Excel
(scripts/measure_range_methods.py, tests/fixtures/range_methods_*.json): the error, what it answers, what the sheet then
reads, and the same after a save and reopen. Where binding changes the answer, "early" holds it, and the model follows
that."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

GROUPS = ('moves', 'clipboard', 'search', 'order')
RECORDS: dict[str, dict[str, Any]] = {
    group: json.loads((Path(__file__).parent / f'fixtures/range_methods_{group}.json').read_text(encoding='utf-8')) for group in GROUPS}
CASES = [(group, number, case) for group in GROUPS for number, case in enumerate(cast('list[list[str]]', RECORDS[group]['cases']))]
IDS = [f"{group}-{number}-{case[1]}" for group, number, case in CASES]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> dict[str, dict[str, dict[int, str]]]:
    answers: dict[str, dict[str, dict[int, str]]] = {}
    for group in GROUPS:
        record = RECORDS[group]
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        # The probe as recorded, its file joined to the folder with this platform's separator.
        app.add_module(record['source'].replace('Application.PathSeparator', f'"{os.sep}"'), name='Methods')
        app.run('Methods.Setup')
        found: dict[str, dict[int, str]] = {'set': {}, 'after': {}}

        def call(name: str, app: ExcelApplication = app) -> str:
            try:
                return str(app.run(f'Methods.{name}'))
            except VBAUnsupportedError:
                return 'gap'

        for number in range(len(record['cases'])):
            found['set'][number] = call(f'S{number}')
        app.run('Methods.Reopen', str(tmp_path_factory.mktemp('methods')))
        for number in range(len(record['cases'])):
            found['after'][number] = call(f'A{number}')
        answers[group] = found
    return answers


#: Late macros no measured rule covers yet.
STILL_TYPED = {('moves', 100), ('clipboard', 79), ('clipboard', 60), ('order', 18), ('order', 96), ('order', 98)}


@pytest.mark.parametrize(('group', 'number', 'case'), CASES, ids=IDS)
def test_each_call_as_excel_answered(run: dict[str, dict[str, dict[int, str]]], group: str, number: int, case: list[str]) -> None:
    # Late macros answer as recorded; what they answer typed is replayed in tests/test_vba_typed_replay.py.
    if (group, number) in STILL_TYPED:
        pytest.xfail('the model answers this late macro otherwise, as no measured rule covers it yet')
    assert run[group]['set'][number] == case[3]
    assert run[group]['after'][number] == case[4]
