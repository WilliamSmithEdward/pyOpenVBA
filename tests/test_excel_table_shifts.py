"""Inserts and deletes of cells, rows and columns beside and through tables, replayed against Excel's answers
(scripts/measure_table_shifts.py): the looks the rules were read from (_a to _c) and each held-out round, its
predictions written before Excel was asked.

Cells put in or taken away inside a table go as table rows or columns, and cells moving along carry a table whole
where it lies wholly in their band; a shift that would cut a table is error 1004, as in src's _table_shifts. In a
filtered table, ListRows.Add inside it, ListRow.Delete and a Delete up are 1004.

Where the model refuses, it is for a reason named in REFUSALS, none of them measured: a Delete of every column of a
table by its cells; a Delete with no Shift in a filtered table, where Excel asks first with DisplayAlerts on; cells
moved across a table with no Shift, put in across its columns and past its rows, or deleted left past its edge; one
edit across several tables; and, as before, whole columns through a table and a table's header row taken by whole
rows.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('table_shifts_a', 'table_shifts_b', 'table_shifts_c', 'table_shifts1',
                                  'table_shifts2', 'table_shifts3', 'table_shifts4', 'table_shifts5')
               if (FIXTURES / f'{stem}.json').exists())
REFUSALS = ("deleting all of a table's columns by deleting cells", 'where Excel asks first',
            'across a table with no Shift', 'across several tables', 'across a table and beyond its rows',
            'inserting or deleting columns through a table', "deleting a table's header row",
            'deleting cells left past the edge of a table')


def _cases() -> list[tuple[str, list[Any], str, str]]:
    found: list[tuple[str, list[Any], str, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, call, answer, record['source'])
                  for call, answer in zip(record['cases'], record['answers'], strict=True)]
    return found


CASES = _cases()


@pytest.mark.parametrize('stem, call, answer, source', CASES,
                         ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_edit_leaves_the_tables_as_excel_did(stem: str, call: list[Any], answer: str, source: str,
                                                  tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
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


@pytest.mark.parametrize('stem', [stem for stem in ROUNDS if (FIXTURES / f'{stem}_predictions.json').exists()])
def test_each_held_out_round_was_predicted_before_excel_answered(stem: str) -> None:
    record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
    predicted = json.loads((FIXTURES / f'{stem}_predictions.json').read_text('utf-8'))
    assert [case['call'] for case in predicted['cases']] == record['cases']
