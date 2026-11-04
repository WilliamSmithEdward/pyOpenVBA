"""What a save keeps of a macro name and of a name that refers to nothing, replayed from live Excel (scripts/measure_name_saves.py,
tests/fixtures/name_saves.json): the Names collection before and after, and workbook.xml's definedNames."""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/name_saves.json').read_text(encoding='utf-8'))
CASES = [f'{name} {extension}' for name in RECORD['cases'] for extension in ('xlsx', 'xlsm')]
#: A macro name's sheet reference comes to name its workbook in Excel, which the model reports.
REPORTED = ('macro then refers',)


def _names_of(path: Path) -> str:
    with zipfile.ZipFile(path) as saved:
        xml = saved.read('xl/workbook.xml').decode('utf-8')
    found = re.search(r'<definedNames>.*?</definedNames>|<definedNames/>', xml, re.DOTALL)
    return found.group(0) if found else ''


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('case', CASES)
def test_a_save_as_excel_wrote_it(tmp_path: Path, case: str) -> None:
    name, extension = case.rsplit(' ', 1)
    path = tmp_path / f'saved.{extension}'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Measure')
    index = list(RECORD['cases']).index(name)
    kind = 51 if extension == 'xlsx' else 52
    if name in REPORTED:
        with pytest.raises(VBAUnsupportedError, match='macro name'):
            app.run('Tried', index, str(path), kind)
        return
    assert str(app.run('Tried', index, str(path), kind)) == RECORD['answers'][case]['read']
    assert _names_of(path) == RECORD['answers'][case]['file']
