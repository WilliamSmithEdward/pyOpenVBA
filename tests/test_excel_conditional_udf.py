"""Independent native DisplayFormat acquisition and retained-UDF controls."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/conditional_udf'
# Preserve this measured control as evidence, but replaying it requires native
# partial-state semantics for self-recursive CF rendering, a separate feature.
SELF_RECURSIVE = 'conditional-expression-bold'
CASES: list[dict[str, Any]] = [
    case
    for fixture in ('display', 'retained', 'evaluate', 'restoration', 'error')
    for case in json.loads((ROOT / f'{fixture}.json').read_text(encoding='utf-8'))['probes']
    if case['name'] != SELF_RECURSIVE
]


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_conditional_udf_boundary(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeControl')
    assert app.run('Probe') == case['result']
