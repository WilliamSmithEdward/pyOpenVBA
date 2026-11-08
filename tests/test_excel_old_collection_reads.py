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
    'Arcs.AddIndent': 'an item setter it needs reports itself',
    'Arcs.AutoScaleFont': 'the item read reports itself',
    'Arcs.Formula': 'an item setter it needs reports itself',
    'Arcs.LockedText': 'an item setter it needs reports itself',
    'Arcs.ReadingOrder': 'an item setter it needs reports itself',
    'Buttons.AddIndent': 'an item setter it needs reports itself',
    'Buttons.AutoScaleFont': 'an item setter it needs reports itself',
    'Buttons.AutoSize': 'an item setter it needs reports itself',
    'Buttons.CancelButton': 'an item setter it needs reports itself',
    'Buttons.DefaultButton': 'an item setter it needs reports itself',
    'Buttons.DismissButton': 'an item setter it needs reports itself',
    'Buttons.Formula': 'an item setter it needs reports itself',
    'Buttons.HelpButton': 'an item setter it needs reports itself',
    'Buttons.HorizontalAlignment': 'an item setter it needs reports itself',
    'Buttons.Orientation': 'an item setter it needs reports itself',
    'Buttons.PhoneticAccelerator': 'an item setter it needs reports itself',
    'Buttons.ReadingOrder': 'an item setter it needs reports itself',
    'Buttons.VerticalAlignment': 'an item setter it needs reports itself',
    'CheckBoxes.Display3DShading': 'an item setter it needs reports itself',
    'CheckBoxes.PhoneticAccelerator': 'an item setter it needs reports itself',
    'DrawingObjects.AddIndent': 'an item setter it needs reports itself',
    'DrawingObjects.ArrowHeadLength': 'an item setter it needs reports itself',
    'DrawingObjects.ArrowHeadStyle': 'an item setter it needs reports itself',
    'DrawingObjects.ArrowHeadWidth': 'an item setter it needs reports itself',
    'DrawingObjects.AutoSize': 'an item setter it needs reports itself',
    'DrawingObjects.CancelButton': 'an item setter it needs reports itself',
    'DrawingObjects.DefaultButton': 'an item setter it needs reports itself',
    'DrawingObjects.DismissButton': 'an item setter it needs reports itself',
    'DrawingObjects.Display3DShading': 'an item setter it needs reports itself',
    'DrawingObjects.HelpButton': 'an item setter it needs reports itself',
    'DrawingObjects.HorizontalAlignment': 'an item setter it needs reports itself',
    'DrawingObjects.LockedText': 'an item setter it needs reports itself',
    'DrawingObjects.Orientation': 'an item setter it needs reports itself',
    'DrawingObjects.PhoneticAccelerator': 'an item setter it needs reports itself',
    'DrawingObjects.ReadingOrder': 'an item setter it needs reports itself',
    'DrawingObjects.RoundedCorners': 'an item setter it needs reports itself',
    'DrawingObjects.Shadow': 'an item setter it needs reports itself',
    'DrawingObjects.VerticalAlignment': 'an item setter it needs reports itself',
    'DrawingObjects.Visible': 'reports itself: Excel reads True as the book stands though an item is hidden, False with the second hidden; a split probe comes first',
    'DropDowns.Display3DShading': 'an item setter it needs reports itself',
    'Lines.ArrowHeadLength': 'an item setter it needs reports itself',
    'Lines.ArrowHeadStyle': 'an item setter it needs reports itself',
    'Lines.ArrowHeadWidth': 'an item setter it needs reports itself',
    'ListBoxes.Display3DShading': 'an item setter it needs reports itself',
    'OptionButtons.Display3DShading': 'an item setter it needs reports itself',
    'OptionButtons.PhoneticAccelerator': 'an item setter it needs reports itself',
    'Ovals.AddIndent': 'an item setter it needs reports itself',
    'Ovals.AutoScaleFont': 'the item read reports itself',
    'Ovals.Formula': 'an item setter it needs reports itself',
    'Ovals.LockedText': 'an item setter it needs reports itself',
    'Ovals.ReadingOrder': 'an item setter it needs reports itself',
    'Ovals.Shadow': 'an item setter it needs reports itself',
    'Pictures.Formula': 'an item setter it needs reports itself',
    'Pictures.Shadow': 'an item setter it needs reports itself',
    'Rectangles.AddIndent': 'an item setter it needs reports itself',
    'Rectangles.AutoScaleFont': 'the item read reports itself',
    'Rectangles.Formula': 'an item setter it needs reports itself',
    'Rectangles.LockedText': 'an item setter it needs reports itself',
    'Rectangles.ReadingOrder': 'an item setter it needs reports itself',
    'Rectangles.RoundedCorners': 'an item setter it needs reports itself',
    'Rectangles.Shadow': 'an item setter it needs reports itself',
    'TextBoxes.AddIndent': 'an item setter it needs reports itself',
    'TextBoxes.AutoScaleFont': 'the item read reports itself',
    'TextBoxes.AutoSize': 'an item setter it needs reports itself',
    'TextBoxes.Formula': 'an item setter it needs reports itself',
    'TextBoxes.HorizontalAlignment': 'an item setter it needs reports itself',
    'TextBoxes.LockedText': 'an item setter it needs reports itself',
    'TextBoxes.Orientation': 'an item setter it needs reports itself',
    'TextBoxes.ReadingOrder': 'an item setter it needs reports itself',
    'TextBoxes.RoundedCorners': 'an item setter it needs reports itself',
    'TextBoxes.Shadow': 'an item setter it needs reports itself',
    'TextBoxes.VerticalAlignment': 'an item setter it needs reports itself',
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
