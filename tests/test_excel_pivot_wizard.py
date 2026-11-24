"""Worksheet.PivotTableWizard, and PivotTableAfterValueChange on a table over a list, replayed from live Excel
(scripts/measure_pivot_wizard.py, tests/fixtures/pivot_wizard.json): the events a sink saw, what the call gave back,
every sheet's tables and the active cell after it, and the table part Excel saved.

The default names count through the Excel session, which ran the cases one after another, where each run here starts
a session of its own: the numbers are taken out before comparing. So are a part's random xr:uid and the cache's id.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_pivot_layouts import REVISIONS

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/pivot_wizard.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = RECORD['cases']
#: The wizard while the active cell is in a table, given arguments: Excel moves that table, which is not built.
MOVES = {9}
#: The consolidation source given as an array of ranges, which Excel refuses with a type mismatch; the model reports
#: a source other than a list unsupported first.
CONSOLIDATION = {11}
_NUMBERED = re.compile(r'PivotTable\d+')
_UID = re.compile(r' xr:uid="\{[0-9A-F-]+\}"')
_CACHE = re.compile(r' cacheId="\d+"')


def _plain(text: str) -> str:
    return _CACHE.sub('', _UID.sub('', _NUMBERED.sub('PivotTable', text.replace('\r\n', '\n'))))


def _run(index: int, folder: Path) -> tuple[str, dict[str, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['log'], name='EventLog')
    app.add_module(RECORD['sink'], name='ExcelSink', kind='class')
    app.add_module(RECORD['source'], name='ProbeModule')
    # TableStyle2 of a table in no style is a NUL, which the harness gave back as a space.
    answer = str(app.run('ProbeModule.Probe', str(folder), index)).replace('\x00', ' ')
    saved = folder / f'c{index}.xlsx'
    parts: dict[str, str] = {}
    if saved.exists():
        with zipfile.ZipFile(saved) as package:
            parts = {name: package.read(name).decode('utf-8') for name in package.namelist()
                     if name.startswith('xl/pivotTables/pivotTable')}
    return answer, parts


def _part(text: str) -> str:
    """A table part from its root element on, without the revision markup Excel adds and the model leaves out."""
    return _plain(REVISIONS.sub('', text[text.find('<pivotTableDefinition'):].replace(' />', '/>')))


def _excel_tables(index: int) -> dict[str, str]:
    return {name: text for name, text in CASES[index][2].items() if name.startswith('xl/pivotTables/pivotTable')}


@pytest.mark.parametrize('index', [one for one in range(len(CASES)) if one not in MOVES | CONSOLIDATION])
def test_the_wizard_answers_as_excel_does(index: int, tmp_path: Path) -> None:
    answer, _ = _run(index, tmp_path)
    assert _plain(answer) == _plain(CASES[index][1])


@pytest.mark.parametrize('index', [one for one in range(len(CASES)) if one not in MOVES | CONSOLIDATION])
def test_the_table_part_saves_as_excels(index: int, tmp_path: Path) -> None:
    _, parts = _run(index, tmp_path)
    excel = _excel_tables(index)
    assert {name: _part(text) for name, text in parts.items()} == {name: _part(text) for name, text in excel.items()}


@pytest.mark.parametrize('index', sorted(MOVES | CONSOLIDATION))
def test_what_was_not_built_reports_itself(index: int, tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index, tmp_path)
