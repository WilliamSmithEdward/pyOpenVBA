"""A line's old ArrowHeadStyle, ArrowHeadLength and ArrowHeadWidth, on a line loaded from a file and on a new one, replayed from live Excel
(scripts/measure_line_arrowheads.py, tests/fixtures/line_arrowheads.json): what the old object and the shape's LineFormat read in the
session and after a reopen, and what the save changed in the drawing part, tag by tag. A new line's creationId differs on every save, in
Excel's run as in the model's, so it is left out."""
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
RECORD: dict[str, Any] = json.loads((FIXTURES / 'line_arrowheads.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_loaded_drawing_edits', Path(__file__).parents[1] / 'scripts/measure_loaded_drawing_edits.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
changes: Any = getattr(_script, 'changes')
DRAWING = re.compile(r'xl/drawings/drawing\d+\.xml$')
CASES: list[dict[str, Any]] = RECORD['cases']
IDS = [f'{"new" if one["fresh"] else "loaded"} {one["case"]}' for one in CASES]


def _quiet(found: list[list[Any]]) -> list[list[Any]]:
    return [one for one in found if not any('a16:creationId' in piece for piece in one[1] + one[2])]


def _saved(fresh: bool, which: int, folder: Path) -> tuple[str, dict[str, str]]:
    book, out = folder / 'in.xlsx', folder / f'out_{which}.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    answer = str(app.run('Probe.Row', str(book), str(out), fresh, which))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if DRAWING.search(name)}


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_arrowhead_reads_and_saves_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    which = [one['case'] for one in CASES if one['fresh'] == case['fresh']].index(case['case'])
    answer, parts = _saved(case['fresh'], which, tmp_path)
    assert answer == case['excel']
    _, untouched = _saved(case['fresh'], -1, tmp_path)
    model = {name: _quiet(changes(untouched[name], part)) for name, part in parts.items()}
    excel = {name: _quiet(case['parts'].get(name, [])) for name in parts}
    assert model == excel
