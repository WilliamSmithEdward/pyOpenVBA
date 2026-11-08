"""A method called on a whole old collection, Worksheet.DropDowns.AddItem "z", Worksheet.Buttons.BringToFront and the like, typed and late,
replayed from live Excel (scripts/measure_old_collection_methods.py, tests/fixtures/old_collection_methods.json): what the call answers,
the sheet's drawings and lists after it, and the same once the workbook is saved and opened again."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_methods.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_old_collection_methods', Path(__file__).parents[1] / 'scripts/measure_old_collection_methods.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
row_source: Any = getattr(_script, 'row_source')
#: Calls the model does not yet answer as Excel did, each with why.
NOT_YET: dict[str, str] = {
    f'{collection}.Duplicate': 'a shape read from a file, a form control or a line does not duplicate yet, and reports itself'
    for collection in ('Arcs', 'Buttons', 'CheckBoxes', 'DrawingObjects', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures',
                       'Rectangles', 'TextBoxes')
}
REPLAYED = [case for case in RECORD['cases'] if f"{case['collection']}.{case['call']}" not in NOT_YET]


@pytest.mark.parametrize('case', REPLAYED, ids=lambda case: f"{case['collection']}.{case['call']} {case['binding']}")
def test_each_call_as_excel_answered(case: dict[str, str], tmp_path: Path) -> None:
    book = tmp_path / 'book.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(row_source(case), name='Probe')
    assert str(app.run('Probe.Row', str(book), str(tmp_path / 'saved.xlsx'))) == case['excel']


def test_not_yet_names_measured_calls_only() -> None:
    measured = {f"{case['collection']}.{case['call']}" for case in RECORD['cases']}
    assert set(NOT_YET) <= measured, set(NOT_YET) - measured
