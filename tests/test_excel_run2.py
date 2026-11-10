"""Application._Run2 beside Application.Run, typed and late bound, replayed from live Excel (scripts/measure_run2.py,
tests/fixtures/run2.json)."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/run2.json').read_text(encoding='utf-8'))
CASES: list[list[str]] = RECORD['cases']
#: What the model reports rather than answers: a Currency given to _Run2, which arrives as the machine's currency text.
REFUSED = {index for index, (call, _) in enumerate(CASES) if 'CCur(' in call}


def _source() -> str:
    """The measurement's project, each refused call answering a word instead."""
    source = RECORD['source']
    for index in REFUSED:
        source, count = re.subn(rf'(    Case {index}\n)(        If how = \d Then r = .*\n){{3}}', r'\1        r = "refused"\n', source)
        assert count == 1
    return source


def _id(index: int) -> str:
    return CASES[index][0]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='P')
    return str(app.run('P.Probe')).split('~')[:-1]


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if index not in REFUSED], ids=_id)
def test_run_and_run2_answer_as_excel_did(index: int, answers: list[str]) -> None:
    assert answers[index] == CASES[index][1]


@pytest.mark.parametrize('index', sorted(REFUSED), ids=_id)
def test_a_currency_given_to_run2_reports_itself(index: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    with pytest.raises(VBAUnsupportedError, match='Currency'):
        app.run('P.Tried', 1, index)
