"""ListObject members replayed from live Excel (scripts/measure_table_members.py, tests/fixtures/table_members.json):
SourceType, Comment, AlternativeText, Summary, ShowAutoFilterDropDown, InsertRowRange, Active, DisplayRightToLeft,
Slicers and the members a table of the sheet's own cells refuses, and the table part a save writes for each."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/table_members.json').read_text(encoding='utf-8'))


def _part(xml: str) -> str:
    """A table part from its table element on, each revision id, which every save makes anew, alike."""
    return re.sub(r'\{[0-9A-F-]{36}\}', '{UID}', xml[xml.index('<table '):])


#: Cases the model refuses: a table over its header row alone, which Excel gives an insert row.
REFUSED = {index for index, case in enumerate(RECORD['cases']) if 'ws.Range("A1:C1"), , xlYes)' in ' '.join(case[0])}


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'].replace('Private Function Tried', 'Public Function Tried'), name='TableProbe')
    return made


@pytest.mark.parametrize('index', range(len(RECORD['cases'])),
                         ids=[' '.join(case[0][-1:]) or case[1][:40] for case in RECORD['cases']])
def test_as_excel_answered(app: ExcelApplication, tmp_path: Path, index: int) -> None:
    statements, _, saved, want = RECORD['cases'][index]
    if index in REFUSED:
        with pytest.raises(VBAUnsupportedError, match='insert row'):
            app.run('TableProbe.Tried', index, str(tmp_path))
        return
    assert app.run('TableProbe.Tried', index, str(tmp_path)) == want
    if saved:
        with zipfile.ZipFile(tmp_path / saved) as package:
            mine = package.read('xl/tables/table1.xml').decode('utf-8')
        assert _part(mine) == _part(RECORD['parts'][saved]), statements
