"""The VBA from issue reports #249, #253, #254, #258, #259, #260 and #261 answers as live Excel does
(scripts/measure_reported_behaviours.py, tests/fixtures/reported_behaviours.json). Each was wrong in 6.5.0 and is
right on main; this keeps it so."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/reported_behaviours.json').read_text(encoding='utf-8'))


def test_each_report_as_excel_answered() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Reported')
    got = str(app.run('Reported.Probe')).split('|')
    assert dict(zip(RECORD['cases'], got, strict=True)) == RECORD['cases']
