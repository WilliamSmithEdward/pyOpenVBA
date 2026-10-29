"""Application's custom lists replayed from live Excel (scripts/measure_custom_lists.py,
tests/fixtures/custom_lists.json): CustomListCount, GetCustomListContents, GetCustomListNum, AddCustomList and
DeleteCustomList, and the lists at work in AutoFill and in Range.Sort's OrderCustom. The probe runs its steps in
order, each on the lists the ones before it left."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/custom_lists.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> dict[str, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ListsProbe')
    return dict(one.split('=', 1) for one in str(app.run('ListsProbe.Probe')).split('@@@'))


@pytest.mark.parametrize('label', list(RECORD['answers']))
def test_each_step_as_excel_answered(answers: dict[str, str], label: str) -> None:
    assert answers[label] == RECORD['answers'][label]
