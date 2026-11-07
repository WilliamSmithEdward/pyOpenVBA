"""Inserts and deletes of cells, rows and columns on a sheet with an advanced filter, replayed against Excel's answers
(scripts/measure_advanced_filter_shifts.py): the first looks the rules were read from, and each held-out round, its
predictions written before Excel was asked.

Filtered in place, the sheet takes edits as one an AutoFilter's criteria filter does: inserting cells is 1004, a
delete up takes whole rows, a delete left the visible cells, whole rows deleted are the visible ones, and the list
follows its _FilterDatabase name. Criteria or Unique that hide no row leave the sheet out of filter mode.

The later looks, on sheets with no filter too, are which way a Delete or Insert with no Shift goes (_d to _h): by the
range's shape, unless nothing lies right of a tall range deleted, or below a wide one; an Insert counts the range's own
cells as below it.

Where the model refuses, it is for one of two reasons: a Delete with no Shift on a filtered sheet with DisplayAlerts on,
where Excel asked first in some runs and not others, and an AdvancedFilter whose criteria header names no column of the
list, which is the filter itself and not the edit.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('advanced_filter_shifts', 'advanced_filter_shifts_b', 'advanced_filter_shifts_c', 'advanced_filter_shifts_d',
                                  'advanced_filter_shifts_e', 'advanced_filter_shifts_f', 'advanced_filter_shifts_g', 'advanced_filter_shifts_h',
                                  'advanced_filter_shifts2', 'advanced_filter_shifts3', 'advanced_filter_shifts4')
               if (FIXTURES / f'{stem}.json').exists())
REFUSALS = ('where Excel asks first', 'unmatched criteria headers')


def _cases() -> list[tuple[str, list[Any], str, str]]:
    found: list[tuple[str, list[Any], str, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, call, answer, record['source']) for call, answer in zip(record['cases'], record['answers'], strict=True)]
    return found


CASES = _cases()


@pytest.mark.parametrize('stem, call, answer, source', CASES, ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_edit_leaves_the_sheet_as_excel_did(stem: str, call: list[Any], answer: str, source: str, tmp_path: Path,
                                                 monkeypatch: pytest.MonkeyPatch) -> None:
    del stem
    # The probe saves and opens the book in TEMP: a folder of its own for each case, as workers run side by side.
    monkeypatch.setenv('TEMP', str(tmp_path))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    try:
        found = str(app.run('Probe.Go', *call))
    except VBAUnsupportedError as refusal:
        assert any(reason in str(refusal) for reason in REFUSALS), str(refusal)
        return
    assert found == answer


def test_the_only_stops_were_excel_prompts() -> None:
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        for call, answer in zip(record['cases'], record['answers'], strict=True):
            assert not answer.startswith(('ERROR', 'CRASH')) or (
                'modal' in answer and call[1].startswith('delete') and call[1].split()[2] == 'auto'), (stem, call)
