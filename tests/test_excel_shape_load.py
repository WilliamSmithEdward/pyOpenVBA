"""Where a shape in a file is, by its anchor whatever its transform says, replayed from live Excel (scripts/measure_drawing_anchors.py,
tests/fixtures/drawing_anchors.json)."""
from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/drawing_anchors.json').read_text(encoding='utf-8'))


def test_each_shape_is_where_excel_read_it(tmp_path: Path) -> None:
    path = tmp_path / 'anchors.xlsx'
    path.write_bytes(base64.b64decode(RECORD['file']))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['reader'], name='Reader')
    assert str(app.run('Reader.Tried', str(path))).split('|') == RECORD['answer'].split('|')
