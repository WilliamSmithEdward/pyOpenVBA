"""A drawing's text linked to a cell through its old object's Formula, and AutoScaleFont, on a drawing loaded from a file and on a new
one, replayed from live Excel (scripts/measure_drawing_text_links.py, tests/fixtures/drawing_text_links.json): what the item reads in the
session and after a reopen, and what the save changed in the drawing part, tag by tag. Excel gives each field and each new shape an id of
its own on every save, so ids are compared as ids; the display's row height, which the run records last, is left out."""
from __future__ import annotations

import importlib.util
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'drawing_text_links.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_loaded_drawing_edits', Path(__file__).parents[1] / 'scripts/measure_loaded_drawing_edits.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
changes: Any = getattr(_script, 'changes')
DRAWING = re.compile(r'xl/drawings/drawing\d+\.xml$')
CASES: list[dict[str, Any]] = RECORD['cases']
IDS = [f'{one["kind"]} {"new" if one["fresh"] else "loaded"} {one["case"]}' for one in CASES]


def _ids(found: list[list[Any]]) -> list[list[Any]]:
    return [[tag, [re.sub(r'\{[0-9A-F-]{36}\}', '{id}', one) for one in old], [re.sub(r'\{[0-9A-F-]{36}\}', '{id}', one) for one in new]]
            for tag, old, new in found]


def _without_new_ids(found: list[list[Any]]) -> list[list[Any]]:
    """A new shape's creationId differs from save to save, in Excel's run as in the model's."""
    return [one for one in _ids(found) if not any('a16:creationId' in piece for piece in one[1] + one[2])]


def _saved(case: dict[str, Any], folder: Path, change: bool) -> tuple[str, dict[str, str]]:
    book, out = folder / 'in.xlsx', folder / f'out_{change}.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    answer = str(app.run('Probe.Row', str(book), str(out), case['kind'], case['fresh'], case['formula'], case['after'], case['scaled'], change))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if DRAWING.search(name)}


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_link_reads_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case, tmp_path, True)
    assert answer.rsplit('|', 1)[0] == case['excel'].rsplit('|', 1)[0]
    _, untouched = _saved(case, tmp_path, False)
    model = {name: _without_new_ids(changes(untouched[name], part)) for name, part in parts.items() if untouched.get(name) != part}
    excel = {name: _without_new_ids(found) for name, found in case['parts'].items() if DRAWING.search(name)}
    assert {name: found for name, found in model.items() if found} == {name: found for name, found in excel.items() if found}


def _linked(formula: str, cell: str = '') -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Function Go() As String
    Dim o As Object
    {cell}
    Set o = ActiveSheet.Rectangles.Add(10, 10, 60, 40)
    o.Formula = "{formula}"
    Go = o.Formula
End Function
''', name='Link')
    return app


@pytest.mark.parametrize('formula', ['=A1:B2', '=Other!A1', 'A1', '=SUM(A1)'])
def test_what_is_not_one_cell_on_the_sheet_says_so(formula: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        _linked(formula).run('Link.Go')


def test_a_cell_in_a_font_not_measured_says_so() -> None:
    with pytest.raises(VBAUnsupportedError):
        _linked('=$B$2', 'ActiveSheet.Range("B2").Value = "x": ActiveSheet.Range("B2").Font.Italic = True').run('Link.Go')
