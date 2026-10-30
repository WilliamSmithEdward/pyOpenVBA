"""A Worksheet variable whose sheet was deleted fails as Excel's does, and a new sheet of the same name never takes its
place; a sheet's own name stays with the sheet through a rename (issue #246; scripts/measure_deleted_sheet.py,
tests/fixtures/deleted_sheet.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/deleted_sheet.json').read_text(encoding='utf-8'))


def test_each_member_as_excel_answered() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Deleted')
    got = str(app.run('Deleted.Probe')).split('@@@')
    assert dict(zip(RECORD['cases'], got, strict=True)) == RECORD['cases']
