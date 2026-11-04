"""Workbook.UpdateLink, ChangeLink, OpenLinks and SetLinkOnData, replayed from live Excel over the files it saved
(scripts/measure_link_edits.py, tests/fixtures/link_edits.json): what each answers, and what a save writes of the links."""
from __future__ import annotations

import base64
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._files import spelled

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/link_edits.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: A file named without its folder, which Excel looks for in its current folder, and a save to another folder.
REFUSED = [index for index, name in enumerate(NAMES)
           if name.endswith(', "Gamma.xlsx"') or name.endswith(', 5') or 'C:\\Windows' in name]
#: On opening a file it saved after a ChangeLink, Excel lists the links in an order that is neither the file's nor the
#: names', which the model does not follow.
REOPENED = [index for index, name in enumerate(NAMES) if 'Set d = Workbooks.Open' in name]
#: The same 1004 as the statement's, in other words: Excel says it could not get a property.
READ = [index for index, name in enumerate(NAMES) if name == 'r = d.UpdateLink']
KEPT = [index for index in range(len(NAMES)) if index not in REFUSED and index not in REOPENED and index not in READ]


def _opened(folder: Path) -> ExcelApplication:
    for name, data in RECORD['files'].items():
        (folder / name).write_bytes(base64.b64decode(data))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='LinkProbe')
    return app


def _clean(text: str, folder: Path) -> str:
    # Off Windows a path in the folder has a slash after it, where Excel's has a backslash.
    text = text.replace(f"file:///{str(folder).lstrip('/')}", 'file:///<folder>')  # a path that starts with / after file:///
    return re.sub(re.escape(str(folder)), '<folder>', text, flags=re.IGNORECASE).replace('<folder>/', '<folder>\\')


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    app = _opened(tmp_path)
    assert _clean(str(app.run('LinkProbe.Tried', index, str(tmp_path))), tmp_path) == RECORD['cases'][index][1]


def test_reading_update_link_with_no_arguments_fails(tmp_path: Path) -> None:
    answer = str(_opened(tmp_path).run('LinkProbe.Tried', READ[0], str(tmp_path)))
    assert answer.startswith('Err 1004 ') and RECORD['cases'][READ[0]][1].startswith('Err 1004 ')


@pytest.mark.parametrize('index', REFUSED, ids=[NAMES[index] for index in REFUSED])
def test_what_is_not_followed_reports_itself(tmp_path: Path, index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _opened(tmp_path).run('LinkProbe.Tried', index, str(tmp_path))


SAVED = [index for index in map(int, RECORD['saved']) if index not in REFUSED]


@pytest.mark.parametrize('index', SAVED, ids=[NAMES[index] for index in SAVED])
def test_a_save_writes_the_links_as_excel_does(tmp_path: Path, index: int) -> None:
    _opened(tmp_path).run('LinkProbe.Tried', index, str(tmp_path))
    with zipfile.ZipFile(tmp_path / 'Dest.xlsx') as saved:
        parts = {name: _clean(saved.read(name).decode('utf-8'), tmp_path).replace('C:\\folder', '<folder>')
                 for name in saved.namelist() if 'externalLink' in name}
    wanted = {name: text for name, text in RECORD['saved'][str(index)].items() if 'externalLink' in name}
    assert parts == wanted


@pytest.mark.parametrize('separator', ['\\', '/', 'native'])
def test_a_formula_names_a_linked_workbook_by_its_folder_in_either_separator(tmp_path: Path, separator: str) -> None:
    """Off Windows the folder is written with /, and a macro may join its file on with either."""
    _opened(tmp_path).run('LinkProbe.Tried', 0, str(tmp_path))
    app = ExcelApplication.open(tmp_path / 'Dest.xlsx', with_vba=False)
    mark = os.sep if separator == 'native' else separator
    app.add_module('Public Sub Put(ByVal formula As String)\n    ActiveSheet.Range("H9").Formula = formula\nEnd Sub\n', name='Writer')
    app.run('Writer.Put', f"='{tmp_path}{mark}[Beta.xlsx]Sheet1'!A1")
    assert app.workbook.sheets_[0].cells_[(9, 8)].formula == '=[2]Sheet1!A1'


def test_a_disk_that_ignores_case_still_gives_the_files_own_spelling(monkeypatch: pytest.MonkeyPatch) -> None:
    """On a Mac ALPHA.XLSX exists, and a workbook opened by that name is still named Alpha.xlsx. Faked here, as CI's other disks do not ignore case."""
    folders = {'/': ['Reports'], '/Reports': ['Alpha.xlsx', 'alpha.bak']}

    def listed(folder: str) -> list[str]:
        return folders[folder.replace('\\', '/')]

    monkeypatch.setattr(os, 'listdir', listed)
    assert spelled('/REPORTS/ALPHA.XLSX', '/').replace('\\', '/') == '/Reports/Alpha.xlsx'
    assert spelled('/REPORTS/new.xlsx', '/').replace('\\', '/') == '/Reports/new.xlsx'


def test_saving_over_a_file_in_another_case_writes_that_file_and_takes_the_name_as_given(tmp_path: Path) -> None:
    """As measured (scripts/measure_case_saves.py). On a disk with case the model must find that file itself."""
    record: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/case_saves.json').read_text(encoding='utf-8'))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(record['source'], name='Saves')
    assert str(app.run('Saves.Tried', str(tmp_path))) == record['answer']
    assert sorted(path.name for path in tmp_path.iterdir()) == record['files']

