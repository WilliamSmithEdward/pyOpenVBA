"""A Set into a variable of a class takes only an object of that class: any other is run-time error 13, and the variable keeps what it held.
Each line compiled alone in live Excel (scripts/measure_typed_set.py)."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_typed_set', ROOT / 'scripts/measure_typed_set.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD = json.loads((ROOT / 'tests/fixtures/typed_set.json').read_text(encoding='utf-8'))
ANSWERS: dict[str, str] = RECORD['answers']
READS = dict(MEASURE.LINES)


def _app(line: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURE.macro(line, READS[line]), name='H')
    return app


@pytest.mark.parametrize('line', [line for line, answer in ANSWERS.items() if answer.startswith('does not compile')])
def test_it_does_not_compile(line: str) -> None:
    with pytest.raises(VBACompileError) as raised:
        _app(line).run('H.Probe')
    assert ANSWERS[line].endswith(raised.value.message)


@pytest.mark.parametrize('line', [line for line, answer in ANSWERS.items() if not answer.startswith('does not compile')])
def test_it_answers_as_excel_did(line: str) -> None:
    assert _app(line).run('H.Probe') == ANSWERS[line]
