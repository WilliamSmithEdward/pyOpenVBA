"""A property read on a whole old collection, typed and late, replayed from live Excel (scripts/measure_old_collection_reads.py,
tests/fixtures/old_collection_reads.json): on the workbook Excel saved and on one the macro makes, as each is, with every item set
alike and with the second alone set. Each walk is one collection, book, state and binding, a fresh book for each member."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_reads.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_old_collection_reads', Path(__file__).parents[1] / 'scripts/measure_old_collection_reads.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
walk_source: Any = getattr(_script, 'walk_source')
members: Any = getattr(_script, 'members')
tiny_png: Any = getattr(_script, 'tiny_png')
SCALAR: dict[tuple[str, str], bool] = {(collection, member): scalar for collection, found in members().items() for member, scalar in found}
#: Members the model does not yet answer as Excel did in every state and binding, each with why (step (b) of the old collections lane).
NOT_YET: dict[str, str] = {
    'Buttons.AutoScaleFont': 'an item setter it needs reports itself',
    'Buttons.AutoSize': 'an item setter it needs reports itself',
    'Buttons.Formula': 'an item setter it needs reports itself',
    'TextBoxes.Formula': "linking again the text box whose file links its text reports itself: what Excel writes is not measured",
    'DrawingObjects.AutoSize': 'an item setter it needs reports itself',
    'DrawingObjects.Shadow': 'an item setter it needs reports itself',
    'DrawingObjects.Visible': 'reports itself: Excel reads True as the book stands though an item is hidden, False with the second hidden; a split probe comes first',
    'Pictures.Formula': 'an item setter it needs reports itself',
    'Pictures.Shadow': 'an item setter it needs reports itself',
    'TextBoxes.AutoSize': 'an item setter it needs reports itself',
}
WALKS: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
for _case in RECORD['cases']:
    if f"{_case['collection']}.{_case['member']}" not in NOT_YET:
        WALKS.setdefault((_case['collection'], _case['book'], _case['state'], _case['binding']), []).append(_case)


def ran(key: tuple[str, str, str, str], folder: Path) -> list[str]:
    collection, book, state, binding = key
    cases = WALKS[key]
    picture = folder / 'dot.png'
    picture.write_bytes(tiny_png())
    saved = folder / 'old_collections.xlsx'
    saved.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(walk_source(collection, binding, book, state, [(one['member'], SCALAR[(collection, one['member'])]) for one in cases]), name='Probe')
    return str(app.run('Probe.Walk', str(saved), str(picture), len(cases))).split('~~')[:len(cases)]


@pytest.mark.parametrize('key', list(WALKS), ids=' '.join)
def test_each_read_as_excel_answered(key: tuple[str, str, str, str], tmp_path: Path) -> None:
    wrong = [(one['member'], one['excel'], got) for one, got in zip(WALKS[key], ran(key, tmp_path), strict=True) if got != one['excel']]
    assert not wrong
