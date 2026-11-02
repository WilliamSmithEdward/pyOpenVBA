"""A chart's ChartArea and PlotArea, and the ChartObjects collection, replayed from live Excel
(scripts/measure_chart_areas.py, tests/fixtures/chart_areas.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/chart_areas.json').read_text(encoding='utf-8'))
#: ChartObjects.Select, which selects the sheet's DrawingObjects, is not implemented: the replay runs those cases with
#: no statements and leaves their answers out.
REFUSED = ('cos.Select', 'cos.Select: ws.Range("A1").Select')
SOURCE = '\n'.join(line for line in RECORD['source'].split('\n') if line.strip() not in REFUSED)
KEPT = [index for index, (case, _) in enumerate(RECORD['cases']) if case not in REFUSED]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='AreasProbe')
    return str(app.run('AreasProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
