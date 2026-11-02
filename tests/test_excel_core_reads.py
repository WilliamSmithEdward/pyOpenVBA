"""Core members everyday macros read, replayed from live Excel (scripts/measure_core_reads.py,
tests/fixtures/core_reads.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/core_reads.json').read_text(encoding='utf-8'))
#: Reads that answer for the harness's session rather than for the workbook: it runs with alerts off.
SESSION = {'Application.DisplayAlerts'}
#: Reads the model answers differently, and why.
KNOWN = {'ws.DrawingObjects.Count': 'DrawingObjects is the sheet\'s Shapes in the model; Excel\'s leaves notes out'}


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='CoreReadsProbe')
    return str(app.run('CoreReadsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', [
    pytest.param(index, marks=pytest.mark.xfail(reason=KNOWN[read], strict=True)) if read in KNOWN else index
    for index, (read, _) in enumerate(RECORD['cases'])], ids=[read for read, _ in RECORD['cases']])
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    read, answer = RECORD['cases'][index]
    if read in SESSION:
        pytest.skip("the harness's own session setting")
    assert answers[index] == answer
