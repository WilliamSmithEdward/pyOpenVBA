"""Native aggregate metadata, compared by name because native enumeration varies."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'functions.json').read_text(encoding='utf-8'))


def test_native_aggregate_function_getters(tmp_path: Path) -> None:
    path = ROOT / 'native-functions.xlsx'
    for _ in range(3):
        app = ExcelApplication.open(path, with_vba=False)
        app.add_module(RECORD['inspection_source'], name='Inspection')
        observed = app.run('InspectFunctions')
        assert isinstance(observed, str)
        assert sorted(observed.split(';')) == sorted(RECORD['inspection_result'].split(';'))
        path = app.save(tmp_path / 'functions.xlsx')
