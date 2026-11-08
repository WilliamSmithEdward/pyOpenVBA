"""Cell shifts through merged cells, with Excel's prompt to unmerge them answered OK or Cancel or with DisplayAlerts off, replayed
against Excel's answers (scripts/measure_merge_shifts.py): the first look the rules were read from, and each held-out round,
its predictions written before Excel was asked. The model takes each case's answer from answer_dialogs, 1 for OK and 2 for
Cancel, as the driver pressed the button in Excel."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._merges import UNMERGE_PROMPT

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('merge_shifts', 'merge_shifts2') if (FIXTURES / f'{stem}.json').exists())


def _cases() -> list[tuple[str, list[Any], str, str, str]]:
    found: list[tuple[str, list[Any], str, str, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, call, answer, alert, record['source'])
                  for call, answer, alert in zip(record['cases'], record['answers'], record['alerts'], strict=True)]
    return found


CASES = _cases()


def _app(source: str, call: list[Any]) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    app.answer_dialogs(2 if call[4] == 'cancel' else 1)
    return app


@pytest.mark.parametrize('stem, call, answer, alert, source', CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_shift_leaves_the_merges_as_excel_did(stem: str, call: list[Any], answer: str, alert: str, source: str, tmp_path: Path,
                                                   monkeypatch: pytest.MonkeyPatch) -> None:
    del stem
    # The probe saves and opens the book in TEMP: a folder of its own for each case, as workers run side by side.
    monkeypatch.setenv('TEMP', str(tmp_path))
    app = _app(source, call)
    assert str(app.run('Probe.Go', *call[:4])) == answer
    asked = [one for one in app.interpreter.dialogs if one[1] == UNMERGE_PROMPT]
    assert len(asked) == (1 if 'unmerge' in alert else 0)
    if asked:
        assert asked[0] == ('MsgBox', UNMERGE_PROMPT, 'Microsoft Excel', 1)
        assert "('OK', True)" in alert


def test_with_nothing_queued_the_prompt_takes_ok(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv('TEMP', str(tmp_path))
    record = json.loads((FIXTURES / 'merge_shifts.json').read_text('utf-8'))
    call, answer = next((call, answer) for call, answer, alert in zip(record['cases'], record['answers'], record['alerts'])
                        if call[4] == 'ok' and 'unmerge' in alert)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(record['source'], name='Probe')
    assert str(app.run('Probe.Go', *call[:4])) == answer


def test_every_case_was_answered() -> None:
    for stem in ROUNDS:
        assert not any(answer.startswith('ERROR') for answer in json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))['answers']), stem
