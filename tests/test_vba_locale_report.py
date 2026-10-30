"""The rows of issue report #252 run in en-US Excel and in the model (scripts/measure_locale_report.py,
tests/fixtures/locale_report.json). Those that do not depend on the locale, AscW's sign, IsNumeric(Empty), Val of
full-width digits, IsNumeric and CDec of formatted numbers, StrComp across kana and widths, and a table name beyond
ASCII, answer as Excel does; those that do, dates and full-width typing, answer as the en-US Excel the model is."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/locale_report.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> dict[str, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Locale')
    return dict(zip(RECORD['cases'], str(app.run('Locale.Probe')).split('|'), strict=True))


@pytest.mark.parametrize('case', list(RECORD['cases']))
def test_each_row_as_en_us_excel_answered(answers: dict[str, str], case: str) -> None:
    assert answers[case] == RECORD['cases'][case]
