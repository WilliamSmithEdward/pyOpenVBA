"""Protection.AllowEditRanges, replayed from live Excel (scripts/measure_allow_edit_ranges.py,
tests/fixtures/allow_edit_ranges.json): what each member answers, how a protected sheet treats a range, what a save writes."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/allow_edit_ranges.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Accounts other than Everyone, and rights other than edit, depend on the machine: the model reports them.
REPORTED = ('"Nobody Here"', '"Everyone", False')
#: Left out: assigning a range to Range without Set, which writes to its cells, and a range of another sheet.
LEFT_OUT = ('.Range = ws.Range', '.Range = "D4"', 'other.Range', '.Parent', '.Application')
#: The same 438, in the model's words: an edit range has no default member.
WORDED = [index for index, name in enumerate(NAMES) if name.endswith('r = ws.Protection.AllowEditRanges(1)')]
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in REPORTED + LEFT_OUT) and index not in WORDED]
SHOWN = [index for index, name in enumerate(NAMES) if any(text in name for text in REPORTED)]


def _run(index: int, path: Path) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='RangesProbe')
    return str(app.run('RangesProbe.Tried', index, str(path)))


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    assert _run(index, tmp_path / 'Book.xlsx') == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', SHOWN, ids=[NAMES[index] for index in SHOWN])
def test_what_is_not_followed_reports_itself(tmp_path: Path, index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index, tmp_path / 'Book.xlsx')


def _salted(text: str) -> str:
    """A saved element, with the hash and salt, which are new each time, taken out."""
    return re.sub(r'(hashValue|saltValue)="[^"]*"', r'\1=""', text)


SAVED = [int(index) for index in RECORD['saved']]


@pytest.mark.parametrize('index', SAVED, ids=[NAMES[index] for index in SAVED])
def test_a_save_writes_them_as_excel_does(tmp_path: Path, index: int) -> None:
    path = tmp_path / 'Book.xlsx'
    _run(index, path)
    with zipfile.ZipFile(path) as saved:
        found = re.search(r'<protectedRanges>.*?</protectedRanges>', saved.read('xl/worksheets/sheet1.xml').decode('utf-8'), re.DOTALL)
    assert _salted(found.group(0) if found else '') == _salted(RECORD['saved'][str(index)])
