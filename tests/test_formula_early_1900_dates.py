"""Bond and date functions at the start of the 1900 calendar, measured by scripts/measure_early_1900_dates.py.

Excel's bond functions give #NUM! for a settlement in the first months
of 1900, its coupon dates there stop at serial 0, and YEARFRAC on basis 1
averages the real years' lengths, 1900 having 365 days. The engine, as
copied from pyOfficeEditor, does none of that yet: those items are in
GAPS and must still differ. The day-count functions that take the dates
as they are already agree.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

PROBE: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/formula/early_1900_dates.json').read_text(encoding='utf-8'))['probes'][0]
#: ODDLPRICE, ODDLYIELD, PRICE, YIELD, COUPDAYS, COUPPCD, ACCRINT, YEARFRAC and DISC near serial 60.
GAPS = frozenset({0, 1, 2, 3, 4, 5, 6, 11, 12, 13, 14, 15, 16, 18, 19, 20})


def test_dates_at_the_start_of_the_1900_calendar() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(PROBE['macro'], name='Harness')
    got = str(app.run('Harness.Probe')).split('|')
    want = PROBE['result'].split('|')
    assert len(got) == len(want) == 27
    for index, (excel, model) in enumerate(zip(want, got, strict=True)):
        if index in GAPS:
            assert model != excel, f'item {index} now matches Excel: take it out of GAPS'
        else:
            assert model == excel, f'item {index}'
