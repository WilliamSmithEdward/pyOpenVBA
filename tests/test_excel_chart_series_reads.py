"""Series members read on each kind of chart, replayed from live Excel (scripts/measure_chart_series_reads.py,
tests/fixtures/chart_series_reads.json)."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_series_reads.json').read_text(encoding='utf-8'))
NAMES: list[str] = RECORD['names']
#: The members the model answers. The rest raise errors in Excel that were read only through CallByName: 440 for
#: the ApplyPictTo trio and 445 for the members of newer chart kinds.
ANSWERED = ['Parent', 'Application', 'Creator', 'AxisGroup', 'ChartType', 'Type', 'IsFiltered', 'HasErrorBars',
            'Has3DEffect', 'BarShape', 'PlotColorIndex', 'PictureType', 'PictureUnit2', 'InvertColorIndex',
            'InvertColor', 'Shadow']


#: Kinds of chart the model does not make yet.
UNMADE = {'xlColumnStacked'}
#: The probe as recorded, reading only the members the model answers, on the kinds it makes.
SOURCE = re.sub(r'    names = Array\(.*\n    names = Split\(.*\n',
                '    names = Array(' + ', '.join(f'"{name}"' for name in ANSWERED) + ')\n',
                re.sub(r'    s = s & "@@@" & Tried\((' + '|'.join(UNMADE) + r')\)\n', '    s = s & "@@@"\n',
                       RECORD['source']))


def _recorded(answer: str) -> list[dict[str, str]]:
    return [{name: value for name, value in zip(NAMES, one.split(','), strict=True) if name in ANSWERED}
            for one in answer.split('|')]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    assert SOURCE != RECORD['source']
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='SeriesReadsProbe')
    return str(app.run('SeriesReadsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_kind_as_excel_answered(answers: list[str], index: int) -> None:
    if RECORD['cases'][index][0] in UNMADE:
        pytest.skip('the model does not make this kind of chart')
    got = [dict(zip(ANSWERED, one.split(','), strict=True)) for one in answers[index].split('|')]
    assert got == _recorded(RECORD['cases'][index][1])
