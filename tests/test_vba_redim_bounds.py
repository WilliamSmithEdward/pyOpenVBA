"""ReDim with an upper bound below its lower bound is error 9, as live VBA refuses it (issue #255;
scripts/measure_redim_bounds.py, tests/fixtures/redim_bounds.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/redim_bounds.json').read_text(encoding='utf-8'))


def test_each_redim_as_vba_answered() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Bounds')
    assert str(app.run('Bounds.Probe')).split('|') == list(RECORD['cases'].values())
