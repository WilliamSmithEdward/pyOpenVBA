"""An empty argument at the end of a call does not compile, in any form, where one in the middle does: each line compiled alone in live Excel
(scripts/measure_trailing_arguments.py)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

RECORD = json.loads((Path(__file__).parent / 'fixtures/trailing_arguments.json').read_text(encoding='utf-8'))
ANSWERS: dict[str, str] = RECORD['answers']


def _app(line: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'].replace('LINE', line), name='H')
    return app


@pytest.mark.parametrize('line', [line for line, answer in ANSWERS.items() if answer.startswith('does not compile')])
def test_it_does_not_compile(line: str) -> None:
    with pytest.raises(VBACompileError) as raised:
        _app(line).run('H.Probe')
    assert ANSWERS[line].endswith(raised.value.message)


@pytest.mark.parametrize('line', [line for line, answer in ANSWERS.items() if not answer.startswith('does not compile')])
def test_it_answers_as_excel_did(line: str) -> None:
    assert _app(line).run('H.Probe') == ANSWERS[line]
