"""Font on a chart's text, everything tests/test_excel_chart_fonts.py left refused, replayed from live Excel: every value each setting
takes and what the chart area's font does to the rest (scripts/measure_chart_font_gaps.py, tests/fixtures/chart_font_gaps.json), Font
on charts Excel saved (scripts/measure_chart_fonts_saved.py), and Subscript, which is not the same from run to run
(scripts/measure_chart_subscript.py).

The model's Font on cells is checked beside it: a setting a macro makes reads the same way wherever the Font came from."""
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
RECORD: dict[str, Any] = json.loads((FIXTURES / 'chart_font_gaps.json').read_text(encoding='utf-8'))
SAVED: dict[str, Any] = json.loads((FIXTURES / 'chart_fonts_saved.json').read_text(encoding='utf-8'))
SUBSCRIPT: dict[str, Any] = json.loads((FIXTURES / 'chart_subscript.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Not a font's own gap: Characters and TextFrame2 on a chart's text, ClearToMatchStyle over labels, and a new type over a look.
REFUSED = [index for index, name in enumerate(NAMES) if '.Characters(' in name or 'TextFrame2' in name or 'ClearToMatchStyle' in name or (' ch.ChartType = xlLine' in name and 'ChartArea' in name)]
#: Excel's Subscript writes one of two things at random (see the test below), so what it wrote is not compared.
RANDOM = [index for index, name in enumerate(NAMES) if '.Subscript =' in name]
#: Superscript over a subscript goes with what the subscript happened to write.
OVER = [index for index, name in enumerate(NAMES) if '.Subscript =' in name and '.Superscript' in name]
KEPT = [index for index in range(len(NAMES)) if index not in REFUSED and index not in OVER]
SAVES = [index for index in KEPT if index not in RANDOM and RECORD['cases'][index][2]]


def _plain(text: str) -> str:
    """A part without Excel's random ids, the axes' numbered as they first appear."""
    text = re.sub(r'\{[0-9A-F]{8}-[0-9A-F-]{27}\}', '{id}', text)
    seen: dict[str, str] = {}

    def numbered(found: re.Match[str]) -> str:
        seen.setdefault(found.group(2), str(len(seen)))
        return f'{found.group(1)}val="{seen[found.group(2)]}"'

    text = re.sub(r'(<c:(?:axId|crossAx) )val="(\d+)"', numbered, text)
    return re.sub(r'<c16r2:[^>]*>|</c16r2:[^>]*>|<c:date1904 val="0"/>|<c:lang [^>]*/>|<c:roundedCorners [^>]*/>', '', text)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_font_gaps')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    return {index: str(app.run('P.Probe', str(folder), index, index)) for index in KEPT}, folder


def _head(answer: str) -> str:
    """What a statement answered, without the fonts of the elements, for those where Superscript and the file vary."""
    return '|'.join(answer.split('|')[:2])


@pytest.mark.parametrize('index', KEPT, ids=[f'{index}: {NAMES[index][:70]}' for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    want = RECORD['cases'][index][1]
    if index in RANDOM:
        assert _head(run[0][index]) == _head(want)
    else:
        assert run[0][index] == want


@pytest.mark.parametrize('index', SAVES, ids=[f'{index}: {NAMES[index][:70]}' for index in SAVES])
def test_each_case_saves_the_part_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        ours = package.read('xl/charts/chart1.xml').decode('utf-8')
    assert _plain(ours) == _plain(next(iter(RECORD['cases'][index][2].values())))


@pytest.mark.parametrize('index', REFUSED, ids=[f'{index}: {NAMES[index][:70]}' for index in REFUSED])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    with pytest.raises(VBAUnsupportedError):
        app.run('P.Probe', str(tmp_path), index, index)


def test_excels_subscript_is_not_the_same_from_run_to_run() -> None:
    """Why it is not replayed: each element gave both answers, in the same run of the same statement. Superscript alone is steady."""
    for label, counts in SUBSCRIPT['tally'].items():
        if label.endswith('superscript'):
            assert list(counts) == ['False/True ']
        else:
            assert len(counts) == 2, label


@pytest.mark.parametrize(('step', 'answer', 'baseline'), [('.Subscript = True', 'True/False', '-25000'), ('.Subscript = False', 'False/False', '0'), ('.Subscript = True: f.Subscript = False', 'False/False', '0')])
def test_the_model_gives_what_was_asked_which_is_one_of_excels_two(tmp_path: Path, step: str, answer: str, baseline: str) -> None:
    label = {'.Subscript = True': 'on', '.Subscript = False': 'off'}.get(step, 'on off')
    for element, font in (('title', 'ch.ChartTitle.Font'), ('legend', 'ch.Legend.Font'), ('area', 'ch.ChartArea.Font')):
        assert f'{answer} {baseline}' in SUBSCRIPT['tally'][f'{element}: {label}']
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(f'''Public Function Go() As String
    Dim ws As Worksheet, ch As Object, f As Object
    Set ws = ActiveWorkbook.Worksheets(1)
    ws.Range("A1:B1").Value = Array("", "R")
    ws.Range("A2:B2").Value = Array("a", 1)
    Set ch = ws.ChartObjects.Add(48, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:B2")
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.HasLegend = True
    Set f = {font}
    f{step}
    Go = f.Subscript & "/" & f.Superscript
End Function
''', name='Sub')
        assert app.run('Sub.Go') == answer
        out = tmp_path / f'{element}.xlsx'
        app.save(out)
        with zipfile.ZipFile(out) as package:
            assert f'baseline="{baseline}"' in package.read('xl/charts/chart1.xml').decode('utf-8')


def test_excel_saved_charts_read_as_excel_reads_them(tmp_path: Path) -> None:
    copy = tmp_path / 'saved.xlsx'
    shutil.copy(FIXTURES / 'chart_fonts_saved.xlsx', copy)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SAVED['source'], name='Saved')
    assert app.run('Saved.Reads', str(copy)) == SAVED['answers']


@pytest.mark.parametrize('member', ['Size = 12', 'Bold = True', 'Color = 255', 'Name = "Arial"'])
def test_setting_a_font_on_an_excel_saved_chart_reports_itself(tmp_path: Path, member: str) -> None:
    copy = tmp_path / 'saved.xlsx'
    shutil.copy(FIXTURES / 'chart_fonts_saved.xlsx', copy)
    app = ExcelApplication.open(copy, with_vba=False)
    app.add_module(f'Public Sub Go()\n    ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart.ChartTitle.Font.{member}\nEnd Sub\n', name='Edited')
    with pytest.raises(VBAUnsupportedError, match='imported chart'):
        app.run('Edited.Go')


@pytest.mark.parametrize(('member', 'value'), [('Size', '14'), ('Bold', 'True'), ('Italic', 'True'), ('Color', '255'), ('ColorIndex', '3'), ('Name', '"Arial"'), ('Underline', 'xlUnderlineStyleDouble'), ('Strikethrough', 'True'), ('FontStyle', '"Italic"')])
def test_a_font_reads_what_it_was_given_on_a_cell_and_on_a_chart(member: str, value: str) -> None:
    """The Font rows of the checklist hold wherever a macro reaches one."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Function Go() As String
    Dim ws As Worksheet, ch As Object
    Set ws = ActiveWorkbook.Worksheets(1)
    ws.Range("A1:B1").Value = Array("", "R")
    ws.Range("A2:B2").Value = Array("a", 1)
    Set ch = ws.ChartObjects.Add(48, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:B2")
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ws.Range("D1").Font.{member} = {value}
    ch.ChartTitle.Font.{member} = {value}
    Go = ws.Range("D1").Font.{member} & "|" & ch.ChartTitle.Font.{member}
End Function
''', name='Both')
    cell, chart = str(app.run('Both.Go')).split('|')
    assert cell == chart
