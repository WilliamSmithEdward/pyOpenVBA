"""When a number's column widens on its own, measured by scripts/measure_general_widening.py.

Writing a value never widens its column. Setting its font's size widens
the column for a number not in General, a date among them, as a number
format does, wherever AutoFit would make it wider; bold does not. A number
General shows widens too, when General is set on it or its font's size
changes, but to a width that is not AutoFit's for its whole text, which
the model does not work out yet: those cases are known gaps.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/general_widening.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: The ways that can widen a column for a number General shows.
GENERAL_WAYS = ('value_general', 'size_after')
GAP = 'a number General shows widens its column to a width of its own, which the model does not work out'


def _cases(result: object) -> dict[str, str]:
    return dict(part.split('=', 1) for part in str(result).split('|') if part)


def _gap(case: str, measured: str) -> bool:
    return case.rsplit('/', 1)[1] in GENERAL_WAYS and measured.split('~')[0] != '64'


CASES = [pytest.param(name, case, id=case,
                      marks=[pytest.mark.xfail(reason=GAP, strict=True)] if _gap(case, measured) else [])
         for name, probe in PROBES.items() for case, measured in _cases(probe['result']).items()]


@pytest.fixture(scope='module')
def replayed() -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    for name, probe in PROBES.items():
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(probe['macro'], name='Harness')
        out[name] = _cases(app.run('Harness.Probe'))
    return out


@pytest.mark.parametrize(('name', 'case'), CASES)
def test_a_column_widens_as_excel_widens_it(replayed: dict[str, dict[str, str]], name: str, case: str) -> None:
    assert replayed[name][case] == _cases(PROBES[name]['result'])[case]
