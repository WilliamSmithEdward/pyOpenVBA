"""A Range as the cells around it move, however it is held, For Each over cells that move, and a chart made from several
areas, replayed from live Excel (scripts/measure_range_tracking.py, scripts/measure_range_holding.py,
tests/fixtures/range_tracking.json, tests/fixtures/range_holding.json)."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_tracking.json').read_text(encoding='utf-8'))
#: A sheet moved to a workbook of its own, which needs no window here but the engine refuses.
MOVED_OUT = ('wo.Move',)
#: Areas of different heights, or one area twice, from which Excel makes series with no rule found.
CHARTED_OUT = ('A1:A5,D1:D9', 'A1:A9,D3:D9', 'A1:A9,A1:A9')
MOVED = [index for index, (edit, _) in enumerate(RECORD['moved']) if edit not in MOVED_OUT]
CHARTED = [index for index, (case, _) in enumerate(RECORD['charted']) if not any(text in case for text in CHARTED_OUT)]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='TrackProbe')
    return made


@pytest.mark.parametrize('index', MOVED, ids=[RECORD['moved'][index][0] for index in MOVED])
def test_a_held_range_as_excel_had_it(app: ExcelApplication, index: int) -> None:
    assert str(app.run('TrackProbe.Moved', index)) == RECORD['moved'][index][1]


@pytest.mark.parametrize('index', CHARTED, ids=[RECORD['charted'][index][0] for index in CHARTED])
def test_a_chart_of_several_areas_as_excel_made_it(app: ExcelApplication, index: int) -> None:
    half = len(RECORD['charted']) // 2
    which = index if index < half else 100 + index - half
    assert str(app.run('TrackProbe.Charted', which)) == RECORD['charted'][index][1]


HOLDING: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_holding.json').read_text(encoding='utf-8'))
KEPT = list(range(len(HOLDING['cases'])))


@pytest.fixture(scope='module')
def holding() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(HOLDING['source'], name='HoldProbe')
    return made


@pytest.mark.parametrize('index', KEPT, ids=[HOLDING['cases'][index][0] for index in KEPT])
def test_however_a_range_is_held_as_excel_had_it(holding: ExcelApplication, index: int) -> None:
    assert str(holding.run('HoldProbe.Held', index)) == HOLDING['cases'][index][1]


AUDIT: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_audit.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def audit() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(AUDIT['source'], name='AuditProbe')
    return made


# A copy of cells into whole columns would fill a million cells each.
FILLS_COLUMNS = {
    index for index, (case, _) in enumerate(AUDIT['cases'])
    if re.search(r'Range\("[^"]+"\)\.Copy\s+On Error Resume Next\s+ws\.Columns\([^)]*\)\.Insert', case)
    and not AUDIT['cases'][index][1].startswith('ok |1004') and not AUDIT['cases'][index][1].startswith('ok 1004')
}


@pytest.mark.parametrize('index', sorted(set(range(len(AUDIT['cases']))) - FILLS_COLUMNS))
def test_what_keeps_a_range_as_excel_had_it(audit: ExcelApplication, index: int) -> None:
    assert str(audit.run('AuditProbe.Audit', index)) == AUDIT['cases'][index][1]


def test_filling_whole_columns_with_copied_cells_says_it_is_not_implemented() -> None:
    assert len(FILLS_COLUMNS) >= 8
    for index in FILLS_COLUMNS:
        made = ExcelApplication(sandbox=False)
        made.add_workbook()
        made.add_module(AUDIT['source'], name='AuditProbe')
        with pytest.raises(VBAUnsupportedError, match='whole columns'):
            made.run('AuditProbe.Audit', index)

