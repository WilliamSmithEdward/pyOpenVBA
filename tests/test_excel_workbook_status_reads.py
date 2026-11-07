"""Workbook.UserStatus's items and Workbook.ReadOnly after each way a macro reaches it, typed and late, as live Excel answered
(scripts/measure_workbook_status_reads.py). The model's answers were written down first; the one it missed, a file whose read-only attribute
is set opening read-only, is fixed. The macro runs from Python with the sandbox off, in an empty temporary folder."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_workbook_status_reads', ROOT / 'scripts/measure_workbook_status_reads.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/workbook_status_reads.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
def test_each_read_answers_as_excel_did(declared: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(MEASURE.macro(declared), name='H')
    assert str(app.run('H.Probe', str(tmp_path))).split(MEASURE.SEPARATOR) == RECORD['excel'][declared]


def test_the_predictions_missed_only_the_read_only_attribute() -> None:
    missed = {line for declared in ('Workbook', 'Object')
              for mine, theirs in zip(RECORD['predicted'][declared], RECORD['excel'][declared]) if mine != theirs for line in (theirs,)}
    assert missed == {'attribute Boolean:True'}
