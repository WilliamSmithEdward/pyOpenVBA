"""Inserts and deletes of cells, rows and columns beside sparklines, replayed against Excel's answers
(scripts/measure_sparkline_shifts.py): the looks the rules were read from (_a to _f) and each held-out round, its
predictions written before Excel was asked.

A sparkline moves with its cell and goes with it, a group left with none going too. Sources and date ranges follow as
formulas' references do, $ or not; a source whose cells all go reads #REF! in its group and "" on its own, and a
file keeps no formula for it; a date range that goes reads "". A cell put in takes the sparkline of the cell before
it, its source moved as far. A group whose sparklines no longer run along a line, reading a block's rows or columns
in turn, reads its Location and SourceData cell by cell. A reopened sheet reads its groups oldest first, and a save
writes them the other way round.

Where a fixture holds the saved book's sparkline markup, the model's save is held to it too, revision ids aside.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('sparkline_shifts_a', 'sparkline_shifts_b', 'sparkline_shifts_c', 'sparkline_shifts_d',
                                  'sparkline_shifts_e', 'sparkline_shifts_f', 'sparkline_shifts1', 'sparkline_shifts2',
                                  'sparkline_shifts3')
               if (FIXTURES / f'{stem}.json').exists())


def _cases() -> list[tuple[str, list[Any], str, list[str] | None, str]]:
    found: list[tuple[str, list[Any], str, list[str] | None, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        files = record.get('files') or [None] * len(record['cases'])
        found += [(stem, call, answer, saved, record['source'])
                  for call, answer, saved in zip(record['cases'], record['answers'], files, strict=True)]
    return found


CASES = _cases()


def _groups(part: str) -> str:
    return re.sub(r' xr2:uid="[^"]*"', '', ''.join(re.findall(r'<x14:sparklineGroups.*?</x14:sparklineGroups>', part,
                                                              re.S)))


@pytest.mark.parametrize('stem, call, answer, saved, source', CASES,
                         ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_edit_leaves_the_sparklines_as_excel_did(stem: str, call: list[Any], answer: str, saved: list[str] | None,
                                                      source: str, tmp_path: Path,
                                                      monkeypatch: pytest.MonkeyPatch) -> None:
    del stem
    # The probe saves and opens the book in TEMP: a folder of its own for each case, as workers run side by side.
    monkeypatch.setenv('TEMP', str(tmp_path))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    assert str(app.run('Probe.Go', *call)) == answer
    if saved is not None:
        with zipfile.ZipFile(tmp_path / 'pyovba_sparkline_shift.xlsx') as book:
            parts = [book.read(f'xl/worksheets/sheet{number}.xml').decode('utf-8') for number in (1, 2)]
        assert [_groups(part) for part in parts] == [_groups(part) for part in saved]


@pytest.mark.parametrize('stem', [stem for stem in ROUNDS if (FIXTURES / f'{stem}_predictions.json').exists()])
def test_each_held_out_round_was_predicted_before_excel_answered(stem: str) -> None:
    """The cases Excel was asked are the ones predicted, in their order."""
    record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
    predicted = json.loads((FIXTURES / f'{stem}_predictions.json').read_text('utf-8'))
    assert [case['call'] for case in predicted['cases']] == record['cases']
