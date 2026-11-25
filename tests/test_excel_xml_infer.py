"""XML imported with no map, replayed from live Excel (scripts/measure_xml_infer.py, tests/fixtures/xml_infer.json): the schema
Excel infers, the map, the table or cells it fills, their types and number formats, and the parts a save writes; workbooks Excel
saved, read back by the model; and the shapes the model does not build, refused as not implemented."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

HERE = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((HERE / 'xml_infer.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = RECORD['cases']
BOOKS = HERE / 'xml_infer'
#: Shapes whose map Excel cannot export, mixed content, and a Destination over a value, which the model refuses.
UNBUILT = {index for index, case in enumerate(CASES)
           if any(shape in case[0] for shape in ('<list><title>T</title>', '<b><m>3</m></b>', 'kind=""k""', '<tag>x</tag><tag>y</tag>',
                                                 '<item>text<n>', 'ws.Range("B2").Value = "x"'))}
#: An import from a file, which Excel keeps as a workbook connection: the model answers, and refuses to save.
FROM_FILE = {index for index, case in enumerate(CASES) if 'XmlImport(f' in case[0]}
_UID = re.compile(r' xr\d?:uid="\{[0-9A-F-]{36}\}"')
_PARTS = re.compile(r'xl/xmlMaps\.xml|xl/tables/.*\.xml')


def _id(index: int) -> str:
    return CASES[index][0][:120]


@pytest.fixture(scope='module')
def replayed(tmp_path_factory: pytest.TempPathFactory) -> list[tuple[str, dict[str, str]]]:
    folder = tmp_path_factory.mktemp('xml_infer')
    for name, text in RECORD['data'].items():
        (folder / name).write_text(text, encoding='utf-8')
    out: list[tuple[str, dict[str, str]]] = []
    for index in range(len(CASES)):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(RECORD['source'], name='P')
        try:
            answer = str(app.run('P.Probe', str(folder), index)).replace(str(folder), '<folder>').replace('\r\n', '\n')
        except VBAUnsupportedError as refused:
            answer = f'refused: {refused}'
        saved = folder / f'c{index}.xlsx'
        parts: dict[str, str] = {}
        if saved.exists():
            with zipfile.ZipFile(saved) as package:
                parts = {name: package.read(name).decode('utf-8') for name in package.namelist() if _PARTS.fullmatch(name)}
            saved.unlink()
        out.append((answer, parts))
    return out


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if index not in UNBUILT | FROM_FILE], ids=_id)
def test_each_case_answers_and_saves_as_excel_did(index: int, replayed: list[tuple[str, dict[str, str]]]) -> None:
    answer, parts = replayed[index]
    assert answer == CASES[index][1]
    theirs = {name: _UID.sub('', text) for name, text in CASES[index][2].items() if _PARTS.fullmatch(name)}
    assert {name: _UID.sub('', text) for name, text in parts.items()} == theirs


@pytest.mark.parametrize('index', sorted(UNBUILT), ids=_id)
def test_shapes_the_model_does_not_build_are_refused(index: int, replayed: list[tuple[str, dict[str, str]]]) -> None:
    assert replayed[index][0].startswith('refused: ') and 'not implemented' in replayed[index][0]


@pytest.mark.parametrize('index', sorted(FROM_FILE), ids=_id)
def test_an_import_from_a_file_answers_and_reports_itself_on_saving(index: int, replayed: list[tuple[str, dict[str, str]]]) -> None:
    assert replayed[index][0].startswith('refused: saving an XML map imported from a file')


_DIGEST = RECORD['source'].split('Private Sub Act(', 1)[0] + '''
Public Function ReadBack(ByVal path As String) As String
    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    ReadBack = Digest(wb)
    wb.Close False
End Function
'''


@pytest.mark.parametrize('book', sorted(path.name for path in BOOKS.glob('*.xlsx')))
def test_a_workbook_excel_saved_reads_back_as_it_was(book: str, tmp_path: Path) -> None:
    index = int(book[1:-5])
    shutil.copy(BOOKS / book, tmp_path / book)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_DIGEST, name='P')
    if index in UNBUILT:
        with pytest.raises(VBAUnsupportedError, match='XML maps of this file'):
            app.run('P.ReadBack', str(tmp_path / book))
        return
    assert str(app.run('P.ReadBack', str(tmp_path / book))) == CASES[index][1].split('|', 2)[2]
