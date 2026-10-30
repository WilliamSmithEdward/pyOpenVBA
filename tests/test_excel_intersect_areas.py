"""Application.Intersect of ranges with several areas, as live Excel answered (issue #257;
scripts/measure_intersect_areas.py, tests/fixtures/intersect_areas.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/intersect_areas.json').read_text(encoding='utf-8'))


def test_each_intersection_as_excel_answered() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Intersections')
    got = str(app.run('Intersections.Probe')).split('|')
    assert got == [answer for _, answer in RECORD['cases']]
