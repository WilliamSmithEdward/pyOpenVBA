"""TableStyle.TableStyleElements on the built-in table styles, replayed from live Excel
(scripts/measure_table_style_elements.py, tests/fixtures/table_style_elements.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/table_style_elements.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What the model reports: the elements of a pivot, slicer or timeline style, and changes to a built-in style's font and borders.
REPORTED = ('PivotStyle', 'SlicerStyle', 'Font.Bold = True', 'LineStyle = xlContinuous')
#: The font's palette number, which the generated table does not keep, and an interior's pattern.
LEFT_OUT = ('Font.ColorIndex', 'Interior.Pattern', 'Creator', '.Count')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in REPORTED + LEFT_OUT)]
SHOWN = [index for index, name in enumerate(NAMES) if 'HasFormat' in name and any(text in name for text in REPORTED) or 'Font.Bold = True' in name or 'LineStyle = xlContinuous' in name]


def _run(index: int) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='StyleProbe')
    return str(app.run('StyleProbe.Tried', index, index))


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(index: int) -> None:
    assert _run(index) == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', SHOWN, ids=[NAMES[index] for index in SHOWN])
def test_what_is_not_followed_reports_itself(index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index)


def test_every_style_counts_43_elements() -> None:
    answers = [answer for name, answer in RECORD['cases'] if name.endswith('.TableStyleElements.Count')]
    assert set(answers) == {'ok Long:43'}
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function Go() As Long\n    Go = ActiveWorkbook.TableStyles("PivotStyleLight16").TableStyleElements.Count\nEnd Function\n', name='Counting')
    assert app.run('Counting.Go') == 43
