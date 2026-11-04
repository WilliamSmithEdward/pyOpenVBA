"""How a value fits where it goes, against live Excel (scripts/measure_excel_fits.py)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
RECORD = json.loads((ROOT / 'tests/fixtures/excel_fits.json').read_text(encoding='utf-8'))
ANSWERS: dict[str, str] = RECORD['answers']


@pytest.mark.parametrize('key', sorted(ANSWERS))
def test_a_value_fits_as_excel_fitted_it(key: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['helpers'] + RECORD['sources'][key], name='Harness')
    try:
        assert app.run('Harness.Probe') == ANSWERS[key]
    except VBAUnsupportedError:
        pytest.fail('reports itself as not implemented')
