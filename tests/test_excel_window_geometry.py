"""Window geometry's rules, replayed against live Excel at 144 DPI (scripts/measure_window_geometry.py,
tests/fixtures/window_geometry_a.json and _b): each measured case is laid out again with the sizes Excel's sheet had
(rows 29 pixels, columns 96, at 100%), and the rules' UsableWidth, UsableHeight, VisibleRange, ScrollRow and ScrollColumn
are compared with what Excel answered. The model lays sheets out at 96 DPI, which is not measured, so its Window
members report themselves until a run at StandardHeight 15 adds that display.
"""
from __future__ import annotations

import importlib.util
import json
from collections.abc import Callable
from dataclasses import replace
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba._a1 import column_letter
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._window_geometry import (DISPLAYS, Layout, Unmeasured, application_usable, large_scroll,
                                                   origin, range_from_point, screen_pixel, scroll_into_view, snapped,
                                                   usable, visible)
from pyopenvba.exceptions import VBAUnsupportedError


def _script() -> ModuleType:
    path = Path(__file__).parents[1] / 'scripts/measure_window_geometry.py'
    spec = importlib.util.spec_from_file_location('measure_window_geometry', path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()

FIXTURES = Path(__file__).parent / 'fixtures'
DISPLAY = DISPLAYS[14.5]
BASE = Layout(width=700, height=500)


def _rows(pixels: int, upto: int) -> Callable[[int], int]:
    def size(row: int) -> int:
        return pixels if row <= upto else 29
    return size


def _hidden(first: int, last: int, default: int) -> Callable[[int], int]:
    def size(index: int) -> int:
        return 0 if first <= index <= last else default
    return size


def _narrow(column: int) -> int:
    return 40 if column <= 26 else 96


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
    'narrow columns': (replace(BASE, column_pixels=_narrow), None),
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


#: Measured cases the rules miss, run so that a change that fixes one fails until its mark goes.
ZOOM_85 = 'zoom 85: the headings at 85% are not 34 pixels, so Excel shows a column less'


@pytest.mark.parametrize('name, reads, answers', [
    pytest.param(*one, id=one[0], marks=pytest.mark.xfail(strict=True, reason=ZOOM_85)) if one[0] == 'zoom 85'
    else pytest.param(*one, id=one[0]) for one in MEASURED])
def test_the_rules_answer_as_excel_did_at_144_dpi(name: str, reads: list[str], answers: dict[str, str]) -> None:
    layout, action = CASES[name]
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
                                  'w.ScrollIntoView 0, 0, 10, 10', 'w.PointsToScreenPixelsX(0)',
                                  'w.PointsToScreenPixelsY(0)', 'w.RangeFromPoint(100, 100)', 'Application.Width',
                                  'Application.Height', 'Application.Left', 'Application.Top', 'Application.UsableWidth',
                                  'Application.UsableHeight', 'Application.Width = 600 ', 'Application.Top = 10 '])
def test_the_model_reports_them_at_its_own_dpi(read: str) -> None:
    """The model lays sheets out at 96 DPI, which no run has measured yet."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    statement = read if 'Scroll' in read or read.endswith(' ') else f'v = {read}'
    if 'RangeFromPoint' in read:
        statement = f'Set v = {read}'
    app.add_module('Function Go() As String\n    Dim w As Object, v As Variant\n    Set w = ActiveWindow\n'
                   '    w.Width = 700: w.Height = 500: w.Left = 0: w.Top = 0\n' f'    {statement}\nEnd Function\n',
                   name='P')
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


def _missed(stem: str, number: int) -> str | None:
    if stem == 'window_geometry1' and number in HEADINGS_OFF_100:
        return 'headings at 50% and 60% are wider than their size at 100% times the zoom'
    if stem == 'window_geometry2' and number in START_FALSE_ROWS:
        return 'ScrollIntoView with Start False: which first row it takes is not known'
    return None


@pytest.mark.parametrize('stem, number, spec, answer', [
    pytest.param(*one, id=f'{one[0]}-{one[1]}', marks=pytest.mark.xfail(strict=True, reason=reason))
    if (reason := _missed(one[0], one[1])) is not None else pytest.param(*one, id=f'{one[0]}-{one[1]}')
    for one in ROUNDS])
def test_the_held_out_rounds_as_excel_answered(stem: str, number: int, spec: dict[str, Any], answer: str) -> None:
    try:
        found = str(SCRIPT.predicted(spec))
    except Unmeasured:
        # Only frozen panes that leave the scrolling pane less than a row or column are refused.
        assert spec['frozen'] != [0, 0]
        return
    assert found == answer


# -- Application's frame and screen pixels (tests/fixtures/window_geometry_c.json)

LOOK_C = {name: (reads, answer) for name, _statements, reads, answer in json.loads(
    (FIXTURES / 'window_geometry_c.json').read_text('utf-8'))['cases']}


def _values(name: str) -> dict[str, str]:
    reads, answer = LOOK_C[name]
    assert answer.startswith('dpi=14.5|') and answer.endswith(' dpi=14.5'), answer
    return dict(zip(reads, answer.removeprefix('dpi=14.5|').removesuffix(' dpi=14.5').split('|'), strict=True))


#: Each frame case: what Application's frame is set to, (width, height), and whether the bars show.
FRAMES: dict[str, tuple[float, float, bool, bool]] = {
    'frame as set': (700, 500, True, True),
    **{f'application width {value}': (value, 500, True, True) for value in (600, 600.25, 600.5, 600.75, 601)},
    **{f'application height {value}': (700, value, True, True) for value in (450, 450.25, 450.5, 450.75)},
    **{f'window width {value}': (value, 500, True, True) for value in (400, 400.25, 400.5, 401)},
    **{f'window height {value}': (700, value, True, True) for value in (300, 300.25, 300.5, 301)},
    'formula bar off': (700, 500, False, True),
    'status bar off': (700, 500, True, False),
    'both bars off': (700, 500, False, False),
}


def _number(text: str) -> float:
    return float(text.split(':', 1)[1])


@pytest.mark.parametrize('name', list(FRAMES))
def test_application_s_frame_reads_as_excel_did(name: str) -> None:
    """Application's size reads back down to a whole pixel, the window's usable area follows it, and
    Application.UsableWidth and UsableHeight are that size less the frame, ribbon and bars."""
    width, height, formula_bar, status_bar = FRAMES[name]
    values = _values(name)
    assert _number(values['Application.Width']) == snapped(DISPLAY, width)
    assert _number(values['Application.Height']) == snapped(DISPLAY, height)
    assert (_number(values['Application.UsableWidth']), _number(values['Application.UsableHeight'])) == \
        application_usable(DISPLAY, width, height, formula_bar, status_bar)
    layout = replace(BASE, width=width, height=height, formula_bar=formula_bar, status_bar=status_bar)
    assert (_number(values['w.UsableWidth']), _number(values['w.UsableHeight'])) == usable(DISPLAY, layout)


#: Each screen-pixel case: the window's layout and its place in points.
PIXEL_CASES: dict[str, tuple[Layout, float, float]] = {
    'pixels': (BASE, 0, 0),
    'pixels moved': (BASE, 100, 60),
    'pixels headings off': (replace(BASE, headings=False), 0, 0),
    'pixels formula bar off': (replace(BASE, formula_bar=False), 0, 0),
    'pixels scrolled': (replace(BASE, scroll=(5, 2)), 0, 0),
    'pixels hidden': (replace(BASE, row_pixels=_hidden(1, 1, 29), column_pixels=_hidden(1, 1, 96)), 0, 0),
    'pixels shape': (BASE, 0, 0),
    'pixels rows 100': (replace(BASE, scroll=(100, 1)), 0, 0),
}
#: The points PointsToScreenPixelsX and Y were asked for, in the probe's order.
POINTS = [('x', 0), ('y', 0), ('x', 1), ('y', 1), ('x', 47), ('x', 48), ('x', 97), ('y', -30), ('x', 1000), ('x', 2.6),
          ('y', 2.4)]
#: The screen pixels RangeFromPoint was asked for: from A1's corner, or (None) on the screen itself.
SPOTS = [(0, 0, True), (-1, 0, True), (0, -1, True), (95, 0, True), (96, 0, True), (0, 28, True), (0, 29, True),
         (500, 300, True), (-45, 10, True), (10, -35, True), (0, 0, False), (-5, -5, False), (1300, 10, True),
         (10, 600, True)]


@pytest.mark.parametrize('name', list(PIXEL_CASES))
def test_screen_pixels_and_the_cells_under_them_as_excel_answered(name: str) -> None:
    """PointsToScreenPixelsX and Y, and RangeFromPoint, for a window at a whole pixel at 100%. RangeFromPoint's
    Nothing reads as error 91 in the probe, which took its Address."""
    layout, left, top = PIXEL_CASES[name]
    reads, answer = LOOK_C[name]
    values = answer.removeprefix('dpi=14.5|').removesuffix(' dpi=14.5').split('|')
    across, down = origin(DISPLAY, layout, left, top)
    for (axis, points), value in zip(POINTS, values, strict=False):
        assert value == f'Long:{screen_pixel(across if axis == "x" else down, points)}', (axis, points)
    for (x, y, relative), value in zip(SPOTS, values[len(POINTS):], strict=True):
        found = range_from_point(DISPLAY, layout, left, top, across + x if relative else x,
                                 down + y if relative else y)
        want = 'E91' if found is None else f'String:Range:{column_letter(found[1])}{found[0]}'
        assert value == want, (x, y, relative)
    assert len(reads) == len(POINTS) + len(SPOTS)


def test_the_third_look_s_cases_left_out_are_named() -> None:
    """Left out: Application's Left and Top, which snap in ways four cases each do not pin; a frame set below its
    least width (99pt); zooms other than 100%, where the headings come from a table of their own; a wide column whose
    width was not read; and frozen panes, whose screen pixels follow no rule found."""
    left_out = {f'application left {value}' for value in (30, 30.25, 30.5, 31)} | \
        {f'application top {value}' for value in (30, 30.25, 30.5, 31)} | \
        {f'window left {value}' for value in (10, 10.5, 11)} | {f'window top {value}' for value in (10, 10.5, 11)} | \
        {'application too small', 'pixels tall row', 'pixels frozen'} | \
        {f'pixels zoom {zoom}' for zoom in (50, 60, 75, 85, 110, 150, 200)}
    assert set(LOOK_C) - set(FRAMES) - set(PIXEL_CASES) == left_out
