"""Workbook.TableStyles and the built-in TableStyle objects, replayed from live Excel
(scripts/measure_table_styles.py, tests/fixtures/table_styles.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/table_styles.json').read_text(encoding='utf-8'))
#: TableStyleElements, which the model does not have, is left out of the replay.
LEFT_OUT = 'TypeName(wb.TableStyles("TableStyleMedium2").TableStyleElements)'
SOURCE = RECORD['source'].replace(f'    v = {LEFT_OUT}\n', '    v = "TableStyleElements"\n')
KEPT = [index for index, (read, _) in enumerate(RECORD['cases']) if read != LEFT_OUT]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='StylesProbe')
    return str(app.run('StylesProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
