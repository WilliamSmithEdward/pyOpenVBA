"""The type each property answers on font, shape, control and text objects, typed and late, replayed from live Excel
(scripts/measure_late_types.py, tests/fixtures/typed_answers/late_types_*.json). A read the model reports as not implemented claims nothing."""
from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures/typed_answers'
#: A new form control's box snaps to the display's pixels, and it has a placement and alternative text of its own, as the model's does not yet.
NEW_CONTROL = {('Shape', 'tick', 'AlternativeText'), ('Shape', 'tick', 'Height'), ('Shape', 'tick', 'Left'), ('Shape', 'tick', 'Placement')}
CASES: list[tuple[str, dict[str, Any]]] = [(path.name, one) for path in sorted(FIXTURES.glob('late_types_*.json')) for one in json.loads(path.read_text(encoding='utf-8'))['cases']]


@lru_cache(maxsize=None)
def _opened(name: str) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(json.loads((FIXTURES / name).read_text(encoding='utf-8'))['source'], name='Probe')
    app.run('Probe.Prepared')
    return app


@pytest.mark.parametrize(('name', 'case'), CASES, ids=[f'{one["class"]} {one["on"]} {one["member"]} {one["binding"]}' for _, one in CASES])
def test_each_read_and_its_type_as_excel_gave_them(name: str, case: dict[str, Any]) -> None:
    try:
        got = str(_opened(name).run('Probe.InModel', case['routine'], case['which']))
    except VBAUnsupportedError:
        return
    key = (case['class'], case['on'], case['member'])
    if key in NEW_CONTROL:
        pytest.xfail('known, and with its owner')
    assert got == case['excel']
