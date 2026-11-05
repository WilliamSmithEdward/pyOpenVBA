"""What Len and LenB compile with and what they answer, each line compiled alone in live Excel (scripts/measure_vba_len_arguments.py).
A number, a date or a Boolean with no variable to read as storage does not compile; a variable reads as its storage however it is reached."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

RECORD = json.loads((Path(__file__).parent / 'fixtures/vba_len_arguments.json').read_text(encoding='utf-8'))
ANSWERS: dict[str, str] = RECORD['answers']
#: These answer the clock.
BY_THE_CLOCK = ('Len(Now)',)


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


@pytest.mark.parametrize('line', [line for line, answer in ANSWERS.items() if not answer.startswith('does not compile') and line not in BY_THE_CLOCK])
def test_it_answers_as_excel_did(line: str) -> None:
    assert _app(line).run('H.Probe') == ANSWERS[line]


def test_a_procedure_called_len_is_no_len() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Private Function Len(x As Long) As Long\n    Len = x + 1\nEnd Function\nFunction Probe()\n    Probe = Len(5)\nEnd Function\n', name='H')
    assert app.run('H.Probe') == 6
