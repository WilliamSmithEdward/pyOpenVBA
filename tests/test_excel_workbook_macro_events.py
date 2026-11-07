"""The pivot table events a macro fires, as live Excel raised them (scripts/measure_workbook_macro_events.py): PivotTableChangeSync on the sheet
and the workbook, never the application, before PivotTableUpdate on a refresh and after it on any other edit, and alone when a value written
into the table is refused. The data-model actions raise events the model does not have a data model for, so they are not replayed here."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_workbook_macro_events', ROOT / 'scripts/measure_workbook_macro_events.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/workbook_macro_events.json').read_text(encoding='utf-8'))
#: The data model's: ModelChange, PivotTableOpenConnection and SheetTableUpdate fire there, and the model has no data model.
MODEL_ACTIONS = frozenset({'model_add', 'model_pivot', 'model_table'})


@pytest.mark.parametrize('action', sorted(set(MEASURE.ACTIONS) - MODEL_ACTIONS))
def test_the_pivot_events_fire_as_in_excel(action: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Log As String\n', name='EventLog')
    app.add_module(MEASURE.SINK, name='BookSink', kind='class')
    app.add_module(MEASURE.probe(MEASURE.ACTIONS[action]), name='H')
    assert app.run('H.Probe') == RECORD['answers'][action]


def test_the_data_model_actions_fire_what_the_model_has_no_data_model_for() -> None:
    fired = {event for action in MODEL_ACTIONS for event in RECORD['answers'][action].split('~')[1].split('|') if event}
    assert {'ModelChange', 'PivotTableOpenConnection', 'SheetTableUpdate'} <= fired
