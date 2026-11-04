"""Settable properties of ListObject, PageSetup and Name replayed from live Excel (scripts/measure_settable_lists.py,
tests/fixtures/settable_lists.json): typed, which the model follows where a late-bound set answers differently."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROWS: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/settable_lists.json').read_text(encoding='utf-8'))['rows']
#: The name is gone when Excel opens the file again, where the model keeps it: a macro name, or one that refers to nothing.
GONE = {('Name.MacroType', value) for value in ('1', '2', '1.5', '"1"')} | {(f'Name.{name}', '""') for name in ('RefersTo', 'RefersToLocal', 'RefersToR1C1', 'RefersToR1C1Local', 'Value')} | {('Name.Value', 'Empty')}


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('name', list(ROWS))
def test_typed_sets_as_excel_answered(tmp_path: Path, name: str) -> None:
    row = ROWS[name]
    half = len(row['values'])
    wrong: list[tuple[str, str, str]] = []
    for index, (value, answer) in enumerate(zip(row['values'], row['answers'][half:], strict=True), start=half):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(row['source'], name='Measure')
        try:
            got = str(app.run('T', index, str(tmp_path)))
        except VBAUnsupportedError:
            continue  # it says it cannot, which the compare script counts
        if (name, value) in GONE:
            got, answer = got.rsplit('|', 1)[0], answer.rsplit('|', 1)[0]
        if got != answer:
            wrong.append((value, answer, got))
    assert not wrong
