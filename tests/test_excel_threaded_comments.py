"""Threaded comments, replayed from live Excel (scripts/measure_threaded_comments.py,
tests/fixtures/threaded_comments.json): what each case reads, the parts its workbook saves, and what the saved
workbook reads once opened again.

The fixture's identities and GUIDs are masked, the GUIDs numbered in the order they first appear, so the model's
parts are masked the same way before they are compared. Excel measured signed in to an account, whose person
carries that account's name, id and provider; the model, with no account, writes Application.UserName with the
provider "None", so the persons part is compared with those three left out.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

#: The parts a threaded comment touches, as scripts/measure_threaded_comments.py records them.
PARTS = r'threadedComment|person|comments\d|vmlDrawing|\[Content_Types\]|sheet1\.xml\.rels|workbook\.xml\.rels'


def _masked(text: str, secrets: list[str], guids: dict[str, str]) -> str:
    """The text masked as the probe masks Excel's: each secret <user>, each time <time>, each GUID numbered."""
    for secret in secrets:
        text = text.replace(secret, '<user>')
    text = re.sub(r'\d{4}-\d\d-\d\dT[\d:.]+Z?', '<time>', text)

    def guid(found: re.Match[str]) -> str:
        return guids.setdefault(found.group(0).upper(), f'G{len(guids) + 1}')

    return re.sub(r'[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}', guid, text)


def parts(path: Path, secrets: list[str]) -> dict[str, str]:
    """The touched parts of a saved workbook, masked as the probe masks them."""
    guids: dict[str, str] = {}
    with zipfile.ZipFile(path) as book:
        names = sorted(name for name in book.namelist() if re.search(PARTS, name))
        texts = {name: book.read(name).decode('utf-8') for name in names}
    return {name: _masked(texts[name], secrets, guids) for name in names}

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/threaded_comments.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
#: The name the model writes threaded comments as.
USER = 'Model User'


def _without_account(text: str) -> str:
    """A persons part with each person's account fields left out."""
    return re.sub(r' (?:displayName|userId|providerId)="[^"]*"', '', text)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('threaded_comments')
    app = ExcelApplication()
    app.application.user_name = USER
    app.add_workbook()
    app.add_module(SOURCE, name='ThreadsProbe')
    return str(app.run('ThreadsProbe.Probe', str(folder), USER)).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    answer = RECORD['cases'][index][1]
    got = run[0][index].replace('/None,', '/Windows Live,')
    assert got == answer


#: Opens a saved workbook and lists its threads as the probe lists them.
REOPEN = '''Private who As String

Private Function Thread(ByVal c As Object) As String
    Dim s As String, i As Long
    s = c.Parent.Address(0, 0) & ":" & c.Text & ",resolved=" & c.Resolved & ",replies=" & c.Replies.Count
    s = s & ",author=" & (c.Author.Name = who) & "/" & c.Author.ProviderID & ",date=" & TypeName(c.Date)
    For i = 1 To c.Replies.Count
        s = s & ";reply " & c.Replies(i).Text & "/" & TypeName(c.Replies(i).Parent)
    Next
    Thread = s
End Function

Public Function Reopen(ByVal path As String, ByVal name As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String, i As Long
    who = name
    Set wb = Workbooks.Open(path): Set ws = wb.Worksheets(1)
    s = ",ws.CommentsThreaded.Count=" & ws.CommentsThreaded.Count & ",ws.Comments.Count=" & ws.Comments.Count
    For i = 1 To ws.CommentsThreaded.Count: s = s & "|" & Thread(ws.CommentsThreaded(i)): Next
    wb.Close False
    Reopen = "ok" & s
End Function
'''


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_saved_workbook_reads_back_its_threads(run: tuple[list[str], Path], index: int) -> None:
    answer = RECORD['cases'][index][1]
    if not answer.startswith('ok'):
        pytest.skip('the case raised before it saved a thread')
    app = ExcelApplication()
    app.application.user_name = USER
    app.add_workbook()
    app.add_module(REOPEN, name='ThreadsReopen')
    got = str(app.run('ThreadsReopen.Reopen', str(run[1] / f'c{index}.xlsx'), USER))
    assert got.replace('/None,', '/Windows Live,') == answer


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_workbook_saves_the_parts_excel_saved(run: tuple[list[str], Path], index: int) -> None:
    expected: dict[str, str] = RECORD['cases'][index][2]
    got = parts(run[1] / f'c{index}.xlsx', [USER])
    assert sorted(got) == sorted(expected)
    for name in expected:
        if 'person' in name:
            assert _without_account(got[name]) == _without_account(expected[name]), name
        else:
            assert got[name] == expected[name], name
