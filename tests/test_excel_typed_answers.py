"""Fixtures that kept answers as text, replayed with their types (scripts/measure_typed_answers.py, tests/fixtures/typed_answers/)."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAError

FIXTURES = Path(__file__).parent / 'fixtures'
_spec = importlib.util.spec_from_file_location('measure_typed_answers', Path(__file__).parents[1] / 'scripts/measure_typed_answers.py')
assert _spec is not None and _spec.loader is not None
_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_module)
typed_source: Any = getattr(_module, 'typed_source')
probes_of: Any = getattr(_module, 'probes_of')

CASES: list[tuple[str, str, str]] = []
for _path in sorted((FIXTURES / 'typed_answers').glob('*.json')):
    if _path.stem == 'control_value_types' or _path.stem.startswith('late_types_'):
        continue
    _record: dict[str, Any] = json.loads(_path.read_text(encoding='utf-8'))
    CASES.extend((_path.stem, _one['name'], _one['typed']) for _one in _record['cases'] if not _one['drifted'])


@pytest.mark.parametrize(('fixture', 'name', 'typed'), CASES, ids=[f'{one[0]} {one[1]}' for one in CASES])
def test_the_answer_and_its_type_as_excel_gave_them(fixture: str, name: str, typed: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.interpreter.sandbox = False
    found = next(one for one in probes_of(fixture.replace('__', '/')) if one['name'] == name)
    app.add_module(typed_source(found['source']), name='Harness')
    try:
        got = str(app.run(found['routine']))
    except VBAError as failed:
        got = f'failed {failed}'
    assert got == typed


def test_a_control_value_is_typed_as_it_was_bound() -> None:
    record: dict[str, Any] = json.loads((FIXTURES / 'typed_answers/control_value_types.json').read_text(encoding='utf-8'))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record['source'], name='Harness')
    assert str(app.run('Tried')) == record['answer']

