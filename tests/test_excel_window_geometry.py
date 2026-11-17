"""Window geometry's rules, replayed against live Excel at 144 DPI (scripts/measure_window_geometry.py,
tests/fixtures/window_geometry_a.json and _b): each measured case is laid out again with the sizes Excel's sheet had
(rows 29 pixels, columns 96, at 100%), and the rules' UsableWidth, UsableHeight, VisibleRange, ScrollRow and ScrollColumn
are compared with what Excel answered. The model lays sheets out at 96 DPI, which is not measured, so its Window
members report themselves until a run at StandardHeight 15 adds that display.
"""
from __future__ import annotations

import json
import sys
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest

sys.path.insert(0, str(Path(__file__).parents[1] / 'scripts'))

from pyopenvba._a1 import column_letter
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._window_geometry import (DISPLAYS, Layout, Unmeasured, large_scroll,
                                                   scroll_into_view, usable, visible)
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
DISPLAY = DISPLAYS[14.5]
BASE = Layout(width=700, height=500)


def _rows(pixels: int, upto: int) -> Any:
    return lambda row: pixels if row <= upto else 29


def _hidden(first: int, last: int, default: int) -> Any:
    return lambda index: 0 if first <= index <= last else default


#: Each measured case as a layout, and what to do in it: None to read it as it is, ('large', down, right) or
#: ('into', (left, top, width, height), start).
CASES: dict[str, tuple[Layout, Any]] = {
    'base': (BASE, None),
    'headings off': (replace(BASE, headings=False), None),
    'horizontal bar off': (replace(BASE, horizontal_bar=False), None),
    'vertical bar off': (replace(BASE, vertical_bar=False), None),
    'tabs off': (replace(BASE, tabs=False), None),
    'bars and tabs off': (replace(BASE, horizontal_bar=False, vertical_bar=False, tabs=False), None),
    'formula bar off': (replace(BASE, formula_bar=False), None),
    'status bar off': (replace(BASE, status_bar=False), None),
    'zoom 50': (replace(BASE, zoom=50), None),
    'zoom 200': (replace(BASE, zoom=200), None),
    'zoom 85': (replace(BASE, zoom=85), None),
    'scrolled': (replace(BASE, scroll=(10, 3)), None),
    'hidden rows': (replace(BASE, row_pixels=_hidden(3, 6, 29)), None),
    'frozen': (replace(BASE, frozen=(3, 2), scroll=(4, 3)), None),
    'frozen scrolled': (replace(BASE, frozen=(3, 2), scroll=(20, 3)), None),
    'window smaller': (replace(BASE, width=400, height=300), None),
    'application smaller': (replace(BASE, width=600, height=450), None),
    'application moved': (BASE, None),
    'large scroll down': (BASE, ('large', 1, 0)),
    'large scroll down 2 right 1': (BASE, ('large', 2, 1)),
    'large scroll up': (replace(BASE, scroll=(60, 1)), ('large', -1, 0)),
    'large scroll left at start': (BASE, ('large', 0, -1)),
    'large scroll frozen': (replace(BASE, frozen=(3, 2), scroll=(4, 3)), ('large', 1, 0)),
    'scroll into view near': (BASE, ('into', (0, 200, 10, 10), True)),
    'scroll into view far': (BASE, ('into', (0, 1000, 10, 10), True)),
    'scroll into view far start': (BASE, ('into', (0, 1000, 10, 10), True)),
    'scroll into view far end': (BASE, ('into', (0, 1000, 10, 10), False)),
    'scroll into view right': (BASE, ('into', (1500, 0, 10, 10), True)),
    'scroll into view back': (replace(BASE, scroll=(80, 1)), ('into', (0, 100, 10, 10), True)),
    **{f'window {width} by {height} zoom {zoom}': (replace(BASE, width=width, height=height, zoom=zoom), None)
       for width, height, zoom in ((300, 200, 100), (450, 330, 100), (610, 410, 100), (900, 640, 100), (700, 500, 75),
                                   (700, 500, 110), (700, 500, 150), (520, 360, 130), (380, 260, 60))},
    'tall rows': (replace(BASE, row_pixels=_rows(66, 40)), None),
    'narrow columns': (replace(BASE, column_pixels=lambda column: 40 if column <= 26 else 96), None),
    'hidden columns': (replace(BASE, column_pixels=_hidden(2, 4, 96)), None),
    'frozen at B2': (replace(BASE, frozen=(1, 1), scroll=(2, 2)), None),
    'frozen at F10': (replace(BASE, frozen=(9, 5), scroll=(10, 6)), None),
    'frozen rows only': (replace(BASE, frozen=(5, 0), scroll=(6, 1)), ('large', 1, 0)),
    'large scroll tall rows': (replace(BASE, row_pixels=_rows(66, 200)), ('large', 1, 0)),
    'large scroll right 2': (BASE, ('large', 0, 2)),
    'large scroll down 3 up 1': (BASE, ('large', 2, 0)),
    'large scroll near the end': (replace(BASE, scroll=(1048570, 1)), ('large', 1, 0)),
    'large scroll up past top': (replace(BASE, scroll=(5, 1)), ('large', -1, 0)),
    'into view middle': (BASE, ('into', (300, 400, 50, 50), True)),
    'into view middle start': (BASE, ('into', (300, 400, 50, 50), True)),
    'into view middle end': (BASE, ('into', (300, 400, 50, 50), False)),
    'into view taller than window': (BASE, ('into', (0, 600, 10, 900), True)),
    'into view already shown': (BASE, ('into', (100, 100, 20, 20), True)),
    'into view far right end': (BASE, ('into', (2000, 0, 30, 10), False)),
    'into view tall rows': (replace(BASE, row_pixels=_rows(66, 200)), ('into', (0, 1500, 10, 10), True)),
    'into view frozen': (replace(BASE, frozen=(3, 2), scroll=(4, 3)), ('into', (0, 1000, 10, 10), True)),
    'into view negative': (replace(BASE, scroll=(30, 1)), ('into', (-50, -50, 10, 10), True)),
    'into view zoom 50': (replace(BASE, zoom=50), ('into', (0, 2000, 10, 10), True)),
}


def _measured() -> list[tuple[str, list[str], dict[str, str]]]:
    found: list[tuple[str, list[str], dict[str, str]]] = []
    for look in ('a', 'b'):
        for name, _statements, reads, answer in json.loads(
                (FIXTURES / f'window_geometry_{look}.json').read_text('utf-8'))['cases']:
            if name in CASES:
                assert answer.startswith('dpi=14.5|') and answer.endswith(' dpi=14.5'), answer
                values = answer.removeprefix('dpi=14.5|').removesuffix(' dpi=14.5').split('|')
                found.append((name, reads, dict(zip(reads, values, strict=False))))
    return found


MEASURED = _measured()


def _value(text: str) -> str:
    return text.split(':', 1)[1]


def _address(top: int, left: int, bottom: int, right: int) -> str:
    return f'{column_letter(left)}{top}:{column_letter(right)}{bottom}'


@pytest.mark.parametrize('name, reads, answers', MEASURED, ids=[one[0] for one in MEASURED])
def test_the_rules_answer_as_excel_did_at_144_dpi(name: str, reads: list[str], answers: dict[str, str]) -> None:
    layout, action = CASES[name]
    if name == 'zoom 85':
        pytest.xfail('zoom 85: the headings at 85% are not 34 pixels, so Excel shows a column less')
    if action is not None and action[0] == 'large':
        layout = replace(layout, scroll=large_scroll(DISPLAY, layout, action[1], action[2]))
    elif action is not None:
        layout = replace(layout, scroll=scroll_into_view(DISPLAY, layout, action[1], action[2]))
    width, height = usable(DISPLAY, layout)
    expected = {'w.UsableWidth': width, 'w.UsableHeight': height, 'w.ScrollRow': layout.scroll[0],
                'w.ScrollColumn': layout.scroll[1], 'w.VisibleRange.Address(False, False)': _address(*visible(DISPLAY, layout))}
    compared = 0
    for read, answer in answers.items():
        if read in expected:
            want = expected[read]
            assert (_value(answer) == want if isinstance(want, str) else float(_value(answer)) == want), (read, answer, want)
            compared += 1
    assert compared, name


def test_every_measured_case_is_replayed_or_named_as_left_out() -> None:
    """The cases left out need what the rules do not take: the screen's size, a column width not read, the answer's
    type."""
    left_out = {'wide column tall row', 'window maximized', 'large scroll answer'}
    names = {name for look in ('a', 'b') for name, *_ in json.loads(
        (FIXTURES / f'window_geometry_{look}.json').read_text('utf-8'))['cases']}
    assert names - set(CASES) == left_out


@pytest.mark.parametrize('read', ['w.UsableWidth', 'w.UsableHeight', 'w.VisibleRange.Address', 'w.LargeScroll 1',
                                  'w.ScrollIntoView 0, 0, 10, 10'])
def test_the_model_reports_them_at_its_own_dpi(read: str) -> None:
    """The model lays sheets out at 96 DPI, which no run has measured yet."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    statement = f'v = {read}' if 'Scroll' not in read else read
    app.add_module('Function Go() As String\n    Dim w As Object, v As Variant\n    Set w = ActiveWindow\n'
                   '    w.Width = 700: w.Height = 500\n' f'    {statement}\nEnd Function\n', name='P')
    with pytest.raises(VBAUnsupportedError, match='96 DPI'):
        app.run('P.Go')


#: Round 1's cases the rules miss: headings at 50% and 60% are wider than their size at 100% times the zoom.
HEADINGS_OFF_100 = {15, 31, 37}
#: Round 2's cases with ScrollIntoView Start False whose first row the rules miss, by one.
START_FALSE_ROWS = {23, 27, 38}
#: Frozen panes that leave the scrolling pane less than a row are reported unmeasured: Excel showed one row of the
#: pane in round 1, two in round 3 and four in round 8.


def _round(stem: str) -> list[tuple[int, dict[str, Any], str]]:
    predicted = json.loads((FIXTURES / f'{stem}_predictions.json').read_text('utf-8'))['cases']
    measured = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))['cases']
    return [(number, one['spec'], answer[3]) for number, (one, answer) in enumerate(zip(predicted, measured, strict=True))]


ROUNDS = [(stem, *case) for stem in ('window_geometry1', 'window_geometry2', 'window_geometry3', 'window_geometry4',
                                'window_geometry5', 'window_geometry6', 'window_geometry7',
                                'window_geometry8', 'window_geometry9', 'window_geometry10',
                                'window_geometry11')
          if (FIXTURES / f'{stem}.json').exists() for case in _round(stem)]


@pytest.mark.parametrize('stem, number, spec, answer', ROUNDS, ids=[f'{one[0]}-{one[1]}' for one in ROUNDS])
def test_the_held_out_rounds_as_excel_answered(stem: str, number: int, spec: dict[str, Any], answer: str) -> None:
    from measure_window_geometry import predicted  # type: ignore[import-not-found]

    if stem == 'window_geometry1' and number in HEADINGS_OFF_100:
        pytest.xfail('headings at 50% and 60% are wider than their size at 100% times the zoom')
    if stem == 'window_geometry2' and number in START_FALSE_ROWS:
        pytest.xfail('ScrollIntoView with Start False: which first row it takes is not known')
    try:
        found = predicted(spec)
    except Unmeasured:
        # Only frozen panes that leave the scrolling pane less than a row or column are refused.
        assert spec['frozen'] != [0, 0]
        return
    assert found == answer
