"""Sheets.FillAcrossSheets, replayed from live Excel (scripts/measure_fill_across_sheets.py,
tests/fixtures/fill_across_sheets.json): the probe's own macro runs in the model, every case's sheets read back as
Excel read them, and the Change events a fill raises."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/fill_across_sheets.json').read_text(encoding='utf-8'))


def test_each_case_as_excel_fills_it() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='FillProbe')
    got = str(app.run('FillProbe.Probe')).split('|')
    for case, mine, excels in zip(RECORD['cases'], got, RECORD['cases'].values(), strict=True):
        assert mine.split('~') == excels.split('~'), case


def test_one_change_a_sheet_the_last_first() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['sink'], name='FillSink', kind='class')
    app.add_module(RECORD['events_source'], name='FillEvents')
    assert app.run('FillEvents.Events') == RECORD['events']
