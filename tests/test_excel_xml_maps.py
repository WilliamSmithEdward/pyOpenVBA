"""XML maps from a schema, replayed from live Excel (scripts/measure_xml_maps.py, tests/fixtures/xml_maps.json): each case's
answer and the parts its save writes, in the order and the one folder the measurement ran them in; workbooks Excel saved,
read back by the model; a file whose maps the model does not build, kept as it came; and the sandbox's gate."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBASandboxError, VBAUnsupportedError

HERE = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((HERE / 'xml_maps.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = RECORD['cases']
BOOKS = HERE / 'xml_maps'
#: A map imported from a file, which Excel keeps as a workbook connection to it: the model answers, and refuses to save.
REFUSED = {index for index, case in enumerate(CASES) if re.search(r'm\.Import\(f|wb\.XmlImport\(f', case[0])}
#: Cases that read a column's width, which AdjustColumnWidth fits in the display's points: measured at 144 DPI, where the model
#: emulates 96, so the width they answer is left out and the rest compared.
DISPLAY = {index for index, case in enumerate(CASES) if 'ColumnWidth' in case[0]}
_UID = re.compile(r' xr\d?:uid="\{[0-9A-F-]{36}\}"')
_PARTS = re.compile(r'xl/xmlMaps\.xml|xl/tables/.*\.xml')


def _id(index: int) -> str:
    return CASES[index][0]


@pytest.fixture(scope='module')
def replayed(tmp_path_factory: pytest.TempPathFactory) -> list[tuple[str, dict[str, str]]]:
    """Every case run in order in one folder, as the measurement ran them: an Export finds the file an earlier case wrote."""
    folder = tmp_path_factory.mktemp('xml_maps')
    for name, text in {**RECORD['schemas'], **RECORD['data']}.items():
        (folder / name).write_text(text, encoding='utf-8')
    out: list[tuple[str, dict[str, str]]] = []
    for index in range(len(CASES)):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(RECORD['source'], name='P')
        try:
            # The harness hands text back with its CR LF as LF, as the measurement's answers have it.
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


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if index not in REFUSED], ids=_id)
def test_each_case_answers_and_saves_as_excel_did(index: int, replayed: list[tuple[str, dict[str, str]]]) -> None:
    answer, parts = replayed[index]
    if index in DISPLAY:
        ours, theirs_answer = answer.split('|', 2), CASES[index][1].split('|', 2)
        assert (ours[0], ours[2]) == (theirs_answer[0], theirs_answer[2])
    else:
        assert answer == CASES[index][1]
    theirs = {name: _UID.sub('', text) for name, text in CASES[index][2].items() if _PARTS.fullmatch(name)}
    assert {name: _UID.sub('', text) for name, text in parts.items()} == theirs


@pytest.mark.parametrize('index', sorted(REFUSED), ids=_id)
def test_a_map_imported_from_a_file_reports_itself_on_saving(index: int, replayed: list[tuple[str, dict[str, str]]]) -> None:
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
    """The maps, schemas, tables and mapped cells of a file Excel saved, read by the model as Excel had them before saving."""
    index = int(book[1:-5])
    shutil.copy(BOOKS / book, tmp_path / book)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_DIGEST, name='P')
    assert str(app.run('P.ReadBack', str(tmp_path / book))) == CASES[index][1].split('|', 2)[2]


def _validating(source: Path, target: Path) -> None:
    """A copy of a workbook Excel saved, its map set to show validation errors, which the model does not build."""
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if item.filename == 'xl/xmlMaps.xml':
                data = data.replace(b'ShowImportExportValidationErrors="false"', b'ShowImportExportValidationErrors="true"')
            new.writestr(item, data)


@pytest.mark.parametrize('member', ['ActiveWorkbook.XmlMaps.Count', 'ActiveWorkbook.XmlNamespaces.Count', 'ActiveSheet.Range("A1").XPath.Value',
                                    'ActiveSheet.ListObjects(1).XmlMap.Name', 'ActiveSheet.ListObjects(1).ListColumns(1).XPath.Value'])
def test_maps_the_model_does_not_build_open_and_report_themselves(member: str, tmp_path: Path) -> None:
    book = tmp_path / 'validating.xlsx'
    _validating(BOOKS / 'c22.xlsx', book)
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module(f'Public Function Go() As Variant\n    Go = {member}\nEnd Function\n', name='P')
    with pytest.raises(VBAUnsupportedError, match='XML maps of this file'):
        app.run('P.Go')


def test_maps_the_model_does_not_build_are_saved_as_they_came(tmp_path: Path) -> None:
    book = tmp_path / 'validating.xlsx'
    _validating(BOOKS / 'c22.xlsx', book)
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveSheet.Range("H9").Value = 1\nEnd Sub\n', name='P')
    app.run('P.Go')
    saved = app.save(tmp_path / 'saved.xlsx')
    with zipfile.ZipFile(book) as before, zipfile.ZipFile(saved) as after:
        for name in ('xl/xmlMaps.xml', 'xl/tables/table1.xml'):
            assert after.read(name) == before.read(name), name


@pytest.mark.parametrize('statement', ['ActiveWorkbook.XmlMaps(1).Import folder & "/there.xml"',
                                       'ActiveWorkbook.XmlMaps(1).Export folder & "/out.xml", True',
                                       'ActiveWorkbook.XmlMaps.Add folder & "/there.xsd", "list"'])
def test_the_sandbox_refuses_a_maps_files(statement: str, tmp_path: Path) -> None:
    shutil.copy(BOOKS / 'c22.xlsx', tmp_path / 'book.xlsx')
    app = ExcelApplication.open(tmp_path / 'book.xlsx', with_vba=False)
    app.add_module(f'Public Sub Go(ByVal folder As String)\n    {statement}\nEnd Sub\n', name='P')
    before = sorted(path.name for path in tmp_path.iterdir())
    with pytest.raises(VBASandboxError, match='not supported with sandbox on'):
        app.run('P.Go', str(tmp_path))
    assert sorted(path.name for path in tmp_path.iterdir()) == before
