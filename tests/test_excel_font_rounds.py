"""Every font round's questions replayed against what Excel answered (scripts/font_round.py, at 96 DPI): each answer as Excel gave it, or a refusal the model gives on purpose."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = ('arial', 'times_new_roman', 'arial2', 'arial3', 'times_new_roman3', 'aptos_narrow3', 'calibri3')
#: Why the model may refuse a question here, each for a reason of its own.
REFUSALS = ('AutoFit of a column whose text wraps, is indented or is turned', 'Justify on a line about as wide as its column', 'Justify of more lines than its range has rows',
            'changing the Normal font to', 'wrapped italic text')


def _cases() -> list[tuple[str, int, dict[str, Any], str, str]]:
    found = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'font_round_{stem}.json').read_text('utf-8'))
        predictions = json.loads((FIXTURES / f'font_round_{stem}_predictions.json').read_text('utf-8'))['cases']
        for number, (case, answer) in enumerate(zip(predictions, record['answers'], strict=True)):
            found.append((stem, number, case['call'], answer, record['source']))
    return found


CASES = _cases()


@pytest.mark.parametrize('stem, number, call, answer, source', CASES, ids=[f'{one[0]}-{one[1]}-{one[2][0]}' for one in CASES])
def test_each_question_as_excel_answered_or_refused_on_purpose(stem: str, number: int, call: list[Any], answer: str, source: str) -> None:
    del stem, number
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Round')
    try:
        found = str(app.run(f'Round.{call[0]}', *call[1:]))
    except VBAUnsupportedError as refusal:
        assert any(reason in str(refusal) for reason in REFUSALS), str(refusal)
        return
    assert found == answer


def test_every_round_was_measured_at_96_dpi() -> None:
    for stem in ROUNDS:
        assert json.loads((FIXTURES / f'font_round_{stem}.json').read_text('utf-8'))['standard_height'] == '15', stem
