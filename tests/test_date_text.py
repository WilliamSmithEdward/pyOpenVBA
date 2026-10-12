"""Text read as a date as native VBA reads it: IsDate and CDate (scripts/measure_date_text.py)."""
from __future__ import annotations
import datetime as dt
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

FIXTURE = json.loads((Path(__file__).parent / 'fixtures/date_text.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, str]] = FIXTURE['probes']
MEASURED_YEAR: int = FIXTURE['measured_year']


def expected(result: str) -> str:
    """Text with no year takes the year it is read in; no probe spells out the
    year it was measured in, so that year in a result stands for this one."""
    return result.replace(f"|{MEASURED_YEAR}-", f"|{dt.date.today().year}-", 1)


@pytest.mark.parametrize('probe', PROBES, ids=[repr(probe['text']) for probe in PROBES])
def test_native_date_text(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == expected(probe['result'])


def test_no_probe_spells_out_the_measured_year() -> None:
    assert not any(str(MEASURED_YEAR) in probe['text'] for probe in PROBES)
