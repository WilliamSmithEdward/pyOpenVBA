"""A delete left over an AutoFilter's whole height, replayed against Excel's answers (scripts/measure_filter_left_deletes.py): the
first look the rule was read from, and each held-out round, its predictions written before Excel was asked."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('filter_left_deletes', 'filter_left_deletes2') if (FIXTURES / f'{stem}.json').exists())


def _cases() -> list[tuple[str, list[Any], str, str]]:
    found: list[tuple[str, list[Any], str, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, call, answer, record['source']) for call, answer in zip(record['cases'], record['answers'], strict=True)]
    return found


CASES = _cases()


@pytest.mark.parametrize('stem, call, answer, source', CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_delete_leaves_the_filter_as_excel_did(stem: str, call: list[Any], answer: str, source: str, tmp_path: Path,
                                                   monkeypatch: pytest.MonkeyPatch) -> None:
    del stem
    # The probe saves and opens the book in TEMP: a folder of its own for each case, as workers run side by side.
    monkeypatch.setenv('TEMP', str(tmp_path))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    assert str(app.run('Probe.Go', *call)) == answer


def test_every_round_was_asked_and_answered() -> None:
    assert ROUNDS
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        assert not any(answer.startswith('ERROR') for answer in record['answers']), stem
