"""The hidden _Default of CheckBoxes, OptionButtons, DropDowns, ListBoxes and DrawingObjects, which the type library gives As Long, replayed
from live Excel (scripts/measure_old_collection_defaults.py, tests/fixtures/old_collection_defaults.json): read, read as the collection's
value and set, with the items as they are, alike and different."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_defaults.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_old_collection_defaults', Path(__file__).parents[1] / 'scripts/measure_old_collection_defaults.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
source_of: Any = getattr(_script, 'source_of')
#: Late bound, Excel answers 438 for [_Default], as for every hidden member on an As Object variable; the model has no notion of a variable's
#: type there and follows the typed answers. Read as a value, the collection is 1004 either way, and is replayed.
LATE_HIDDEN = {'read', 'set'}
#: Set over every drawing, Excel sets each control's Value but leaves both option buttons off, by a rule not worked out: it reports itself.
NOT_YET = ('DrawingObjects', 'Typed', 'set')
CASES = [one for one in RECORD['cases'] if not (one['binding'] == 'Late' and one['operation'] in LATE_HIDDEN)
         and (one['collection'], one['binding'], one['operation']) != NOT_YET]


@pytest.mark.parametrize('case', CASES, ids=[f'{one["collection"]} {one["binding"]} {one["state"]} {one["operation"]}' for one in CASES])
def test_each_default_as_excel_answered(case: dict[str, str], tmp_path: Path) -> None:
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source_of(case), name='Probe')
    assert str(app.run('Probe.Row', str(book), case['state'], case['operation'])) == case['excel']


def test_the_set_over_every_drawing_says_it_is_not_built(tmp_path: Path) -> None:
    case = next(one for one in RECORD['cases'] if (one['collection'], one['binding'], one['operation']) == NOT_YET)
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source_of(case), name='Probe')
    with pytest.raises(VBAUnsupportedError, match='_Default'):
        app.run('Probe.Row', str(book), case['state'], case['operation'])
