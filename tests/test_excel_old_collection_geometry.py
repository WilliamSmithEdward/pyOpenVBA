"""Left, Top, Width and Height read and set on a whole old collection, typed and late, replayed from live Excel
(scripts/measure_old_collection_geometry.py, tests/fixtures/old_collection_geometry.json, and scripts/measure_old_collection_top_split.py,
tests/fixtures/old_collection_top_split.json).

Both runs were at StandardHeight 14.5 (144 DPI), and the model emulates 96 DPI, where a form control added at Top 70 sits at 69.75. So the
model is held to what does not depend on the display: on the book a macro makes, set to values on both displays' pixels (multiples of
1.5pt), the member set reads on every item as Excel's did, and the answers are Excel's. The rule itself is checked on Excel's own numbers:
a size goes to each item, an arc taking half; a place moves the items together so the last lands on the value, shifted back so none is
above the sheet. The saved book's shapes are anchored to cells whose rows are 14.5pt there and 15pt in the model, and 31.25 snaps to the
display's pixels; both are measured and left out."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_geometry.json').read_text(encoding='utf-8'))
SPLIT: dict[str, Any] = json.loads((FIXTURES / 'old_collection_top_split.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_old_collection_geometry', Path(__file__).parents[1] / 'scripts/measure_old_collection_geometry.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
row_source: Any = getattr(_script, 'row_source')
OFF_GRID: str = getattr(_script, 'OFF_GRID')
FIELDS = {'Left': 0, 'Top': 1, 'Width': 2, 'Height': 3}
#: What the model does not set yet: the places of arcs, which keep their Left in Excel, and of every drawing, where each kind moves its own
#: way; and a list box's height, which Excel fits to whole rows of the display (the item refuses a size it has not measured).
REFUSED = {('Arcs', 'Left'), ('Arcs', 'Top'), ('DrawingObjects', 'Left'), ('DrawingObjects', 'Top'), ('ListBoxes', 'Height'), ('DrawingObjects', 'Height')}
CASES = [one for one in RECORD['cases'] if one['book'] == 'made' and one['value'] != OFF_GRID]
#: The kind of each of a made book's drawings, in DrawingObjects' order.
ORDER = ('Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Rectangles', 'TextBoxes', 'Pictures')


def _boxes(text: str) -> list[list[float]]:
    return [[float(one) for one in box.split(',')] for box in text.split(';') if box]


def _parts(answer: str) -> tuple[str, list[list[float]], list[list[float]]]:
    before, told, after, items, reopened, _ = answer.split('|')
    return f'{before}|{told}|{after}', _boxes(items), _boxes(reopened.removeprefix('reopened '))


def _kinds(collection: str) -> list[str]:
    return [kind for _ in range(2) for kind in ORDER] if collection == 'DrawingObjects' else [collection] * 2


def _png() -> bytes:
    return bytes(getattr(_script, '_reads_module').tiny_png())


def _model(case: dict[str, str], folder: Path) -> str:
    picture = folder / 'dot.png'
    picture.write_bytes(_png())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(row_source(case), name='Probe')
    return str(app.run('Probe.Row', str(folder / 'unused.xlsx'), str(folder / 'saved.xlsx'), str(picture)))


def test_both_runs_record_their_display() -> None:
    assert RECORD['standard_height'] == SPLIT['standard_height'] == 14.5
    assert all(one['excel'].endswith('|14.5,14.5') for one in RECORD['cases'])


def test_excel_reads_none_and_takes_every_set() -> None:
    assert {_parts(one['excel'])[0] for one in RECORD['cases']} == {'E1004|E0|E1004'}


@pytest.mark.parametrize('case', [one for one in CASES if one['member'] in ('Width', 'Height')], ids=lambda one: f'{one["collection"]} {one["member"]}')
def test_excel_gives_a_size_to_each_item_an_arc_half(case: dict[str, str]) -> None:
    _, items, _ = _parts(case['excel'])
    wanted = float(case['value'])
    for kind, box in zip(_kinds(case['collection']), items, strict=True):
        assert abs(box[FIELDS[case['member']]] - (wanted / 2 if kind == 'Arcs' else wanted)) < 0.01


def _moved(places: list[float], value: float) -> list[float]:
    shift = value - places[-1]
    shift += max(0.0, -min(place + shift for place in places))
    return [place + shift for place in places]


@pytest.mark.parametrize('which', ['top 300', 'top 0', 'top 37.5'])
def test_excel_moves_the_items_so_the_last_lands_on_the_value(which: str) -> None:
    def tops(answer: str) -> list[float]:
        return [box[1] for box in _boxes(answer.split('|')[2])]

    assert tops(SPLIT['cases'][which]) == _moved(tops(SPLIT['cases']['none']), float(which.split()[1]))


def test_excel_halves_an_arc_only_through_its_collection() -> None:
    arcs = {which: [box[3] for box in _boxes(answer.split('|')[3])] for which, answer in SPLIT['cases'].items()}
    assert arcs['one arc height'][0] == 37.5 and arcs['arcs height'] == [37.5, 37.5]


@pytest.mark.parametrize('case', CASES, ids=[f'{one["collection"]} {one["member"]} {one["binding"]}' for one in CASES])
def test_the_model_sets_each_item_as_excel_did(case: dict[str, str], tmp_path: Path) -> None:
    if (case['collection'], case['member']) in REFUSED:
        with pytest.raises(VBAUnsupportedError):
            _model(case, tmp_path)
        return
    answers, items, reopened = _parts(_model(case, tmp_path))
    excel_answers, excel_items, excel_reopened = _parts(case['excel'])
    assert answers == excel_answers
    field = FIELDS[case['member']]
    # An arc's width and height carry the display's float noise in Excel (60.0001564025879 for 60).
    assert [round(box[field], 2) for box in items] == [round(box[field], 2) for box in excel_items]
    assert [round(box[field], 2) for box in reopened] == [round(box[field], 2) for box in excel_reopened]


@pytest.mark.parametrize('which', ['top 300', 'top 0', 'top 37.5', 'one arc height', 'arcs height'])
def test_the_model_follows_the_split(which: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SPLIT['source'], name='Split')
    before, after = (str(app.run('Split.Row', one)).split('|') for one in ('none', which))
    if which.startswith('top'):
        assert [box[1] for box in _boxes(after[2])] == _moved([box[1] for box in _boxes(before[2])], float(which.split()[1]))
    else:
        excel = SPLIT['cases'][which].split('|')
        assert [box[3] for box in _boxes(after[3])] == [box[3] for box in _boxes(excel[3])]
