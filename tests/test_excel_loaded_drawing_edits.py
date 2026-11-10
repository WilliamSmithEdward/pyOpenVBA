"""A setting changed on a drawing loaded from a file, through its old object, replayed from live Excel (scripts/measure_loaded_drawing_edits.py,
tests/fixtures/loaded_drawing_edits.json): a rectangle, an oval, an arc and a text box, what the set answers, what the item reads in the
session and after a reopen, and what the save changed in the drawing part, tag by tag. Excel gives every shape a new creationId on each
save, so those are left out of both sides."""
from __future__ import annotations

import importlib.util
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_drawing_edits.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_loaded_drawing_edits', Path(__file__).parents[1] / 'scripts/measure_loaded_drawing_edits.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
changes: Any = getattr(_script, 'changes')
value_of: Any = getattr(_script, 'value_of')
DRAWING = re.compile(r'xl/drawings/drawing\d+\.xml$')
KINDS = ('Rectangles', 'Ovals', 'Arcs', 'TextBoxes')
#: The settings the model does not yet make on a loaded drawing as Excel does, each with why.
NOT_YET: dict[tuple[str, str, str], str] = {
    **{(kind, 'Orientation', 'xlHorizontal'): 'Excel writes vert="horz" though the text was horizontal; the model writes nothing' for kind in KINDS[:3]},
    **{(kind, 'RoundedCorners', 'True'): 'the old object\'s setter reports itself: Excel renames the shape "Rectangle: Rounded Corners"'
       for kind in ('Rectangles', 'TextBoxes')},
    **{(kind, 'AutoSize', 'True'): 'fitting a shape to its text needs the text measured in its font' for kind in ('Rectangles', 'Ovals', 'TextBoxes')},
}
CASES = [one for one in RECORD['cases'] if (one['kind'], one['member'], one['value']) not in NOT_YET]


def _quiet(found: list[list[Any]]) -> list[list[Any]]:
    """Without the creationIds, and with each TxLink field's id, new on every save, compared as an id."""
    return [[tag, [re.sub(r'\{[0-9A-F-]{36}\}', '{id}', piece) for piece in old], [re.sub(r'\{[0-9A-F-]{36}\}', '{id}', piece) for piece in new]]
            for tag, old, new in found if not any('creationId' in piece for piece in old + new)]


def _saved(case: dict[str, Any], folder: Path, change: bool) -> tuple[str, dict[str, str]]:
    book, out = folder / 'in.xlsx', folder / f'out_{change}.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    member, value = (case['member'], value_of(case['value'])) if change else ('Visible', True)
    answer = str(app.run('Probe.Row', str(book), str(out), case['kind'], member, value, change))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if DRAWING.search(name)}


@pytest.mark.parametrize('case', CASES, ids=[f'{one["kind"]} {one["member"]} {one["value"]}' for one in CASES])
def test_each_setting_answers_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case, tmp_path, True)
    assert answer == case['excel']
    _, untouched = _saved(case, tmp_path, False)
    model = {name: _quiet(changes(untouched[name], part)) for name, part in parts.items()}
    excel = {name: _quiet(case['parts'].get(name, [])) for name in parts}
    assert model == excel


def test_not_yet_names_measured_settings_only() -> None:
    assert set(NOT_YET) <= {(one['kind'], one['member'], one['value']) for one in RECORD['cases']}
