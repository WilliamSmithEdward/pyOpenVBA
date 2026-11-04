"""The text, connections and picture of a shape in a file, replayed from live Excel: what each reads
(scripts/measure_loaded_shape_text.py, scripts/measure_loaded_shape_foreign.py), and what a set reads after and leaves in the file
(scripts/measure_loaded_shape_sets.py). A read or a set the model cannot answer must report itself; it must never answer wrongly."""
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

FIXTURES = Path(__file__).parent / 'fixtures'
TEXT: dict[str, Any] = json.loads((FIXTURES / 'loaded_shape_text.json').read_text(encoding='utf-8'))
FOREIGN: dict[str, Any] = json.loads((FIXTURES / 'loaded_shape_foreign.json').read_text(encoding='utf-8'))
SETS: dict[str, Any] = json.loads((FIXTURES / 'loaded_shape_sets.json').read_text(encoding='utf-8'))
#: What the model reports rather than answers: where lines wrap, a crop's size, whether a font can be embedded, the words of text
#: with a space at its start or an @, and a text warp's kind.
REFUSED_READS = re.compile(r'\.Lines\.Count$|PictureFormat\.Crop|Font\.Embedd|\("(Wordy|Spaces|Hyphens)"\).*\.Words|\("Warped"\).*(Path|Warp)Format$')
#: Sets the model reports: routing an elbow or a curve but for its plain case, connecting one that is turned, and moving a shape one is
#: held to, which any read of a connector on the sheet then reports.
REFUSED_SETS = re.compile(r'^(Elbow|Curve): s\.ConnectorFormat\.(Begin|End)Connect|^(Elbow|Joined): ws\.Shapes\("End1"\)\.Left')


def _one_read(number: int, read: str) -> str:
    return '\n'.join([f'Public Function R{number}() As String', '    Dim ws As Object, v As Variant', '    Set ws = ActiveWorkbook.Worksheets(1)', '    On Error Resume Next', f'    v = {read}',
                      '    If Err.Number Then', f'        R{number} = "E" & Err.Number', '    ElseIf IsObject(v) Then', f'        R{number} = "object"', '    ElseIf IsArray(v) Or IsNull(v) Then',
                      f'        R{number} = TypeName(v)', '    Else', f'        R{number} = CStr(v)', '    End If', 'End Function', ''])


def _as_harness_answers(text: str) -> str:
    """The measuring harness hands answers back with CR as LF and other control characters as spaces."""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ', text.replace('\r', '\n'))


def _reads(book: Path, reads: list[str]) -> tuple[ExcelApplication, list[str | None]]:
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module(''.join(_one_read(number, read) for number, read in enumerate(reads)), name='Reader')
    answers: list[str | None] = []
    for number in range(len(reads)):
        try:
            answers.append(_as_harness_answers(str(app.run(f'Reader.R{number}'))))
        except VBAUnsupportedError:
            answers.append(None)
    return app, answers


@pytest.mark.parametrize(('record', 'book'), [(TEXT, 'loaded_shape_text.xlsx'), (FOREIGN, 'loaded_shape_foreign.xlsx')], ids=['excel_written', 'other_writers'])
def test_every_read_is_excels_or_reports_itself(record: dict[str, Any], book: str, tmp_path: Path) -> None:
    app, answers = _reads(FIXTURES / book, record['reads'])
    wrong = [(read, truth, got) for read, truth, got in zip(record['reads'], record['excel'], answers, strict=True) if got is not None and got != truth]
    refused = [read for read, got in zip(record['reads'], answers, strict=True) if got is None and not REFUSED_READS.search(read)]
    assert not wrong
    assert not refused
    # Reading all of that changes nothing in the drawing a save writes.
    app.save(tmp_path / 'read.xlsx')
    with zipfile.ZipFile(tmp_path / 'read.xlsx') as saved, zipfile.ZipFile(FIXTURES / book) as original:
        assert saved.read('xl/drawings/drawing1.xml') == original.read('xl/drawings/drawing1.xml')


def _anchor(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    for found in re.finditer(r'<xdr:(?:twoCell|oneCell|absolute)Anchor\b.*?</xdr:(?:twoCell|oneCell|absolute)Anchor>', xml, re.S):
        if f'name="{name}"' in found.group():
            return found.group()
    return ''


@pytest.mark.parametrize('number', range(len(SETS['cases'])), ids=[f"{one['shape']}-{index}" for index, one in enumerate(SETS['cases'])])
def test_each_set_reads_and_saves_as_in_excel(number: int, tmp_path: Path) -> None:
    case = SETS['cases'][number]
    book = tmp_path / 'book.xlsx'
    shutil.copyfile(FIXTURES / 'loaded_shape_text.xlsx', book)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SETS['source'], name='Probe')
    if REFUSED_SETS.search(f"{case['shape']}: {case['statement']}"):
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Case1', number, str(book), str(tmp_path / 'saved.xlsx'))
        return
    answer = app.run('Probe.Case1', number, str(book), str(tmp_path / 'saved.xlsx'))
    assert _as_harness_answers(str(answer)) == case['answer']
    assert _anchor(tmp_path / 'saved.xlsx', case['shape']) == case['anchor']


def test_a_changed_setting_over_markup_the_model_cannot_write_reports_itself(tmp_path: Path) -> None:
    """French text, say: writing it afresh would lose its language, so the set is refused rather than lost."""
    book = tmp_path / 'book.xlsx'
    shutil.copyfile(FIXTURES / 'loaded_shape_foreign.xlsx', book)
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets(1).Shapes("French").TextFrame.Characters.Font.Bold = True\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go')


def test_new_words_keep_a_foreign_shapes_frame_and_first_run(tmp_path: Path) -> None:
    book = tmp_path / 'book.xlsx'
    shutil.copyfile(FIXTURES / 'loaded_shape_foreign.xlsx', book)
    app = ExcelApplication.open(book, with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets(1).Shapes("French").TextFrame.Characters.Text = "Salut"\nEnd Sub\n', name='Probe')
    app.run('Probe.Go')
    app.save(tmp_path / 'saved.xlsx')
    assert '<a:rPr lang="fr-FR" sz="1200" dirty="0"/><a:t>Salut</a:t>' in _anchor(tmp_path / 'saved.xlsx', 'French')


@pytest.mark.parametrize('name', ['Pic', 'Tuned', 'Clear'])
@pytest.mark.parametrize(('member', 'value'), [('Brightness', '0.7'), ('Contrast', '0.3'), ('ColorType', '2')])
def test_a_pictures_settings_come_back_from_the_file(name: str, member: str, value: str, tmp_path: Path) -> None:
    source = 'Public Function Go(ByVal name As String) As String\n    With ActiveWorkbook.Worksheets(1).Shapes(name).PictureFormat\n'
    app = ExcelApplication.open(FIXTURES / 'loaded_shape_text.xlsx', with_vba=False)
    app.add_module(source + f'        .{member} = {value}\n        Go = .{member}\n    End With\nEnd Function\n', name='Probe')
    assert str(app.run('Probe.Go', name)) == value
    app.save(tmp_path / 'saved.xlsx')
    reopened = ExcelApplication.open(tmp_path / 'saved.xlsx', with_vba=False)
    reopened.add_module(source + f'        Go = .{member}\n    End With\nEnd Function\n', name='Probe')
    assert str(reopened.run('Probe.Go', name)) == value


def test_changing_a_loaded_connectors_type_reports_itself() -> None:
    """The file would keep its old geometry."""
    app = ExcelApplication.open(FIXTURES / 'loaded_shape_text.xlsx', with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets(1).Shapes("Joined").ConnectorFormat.Type = 2\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go')
