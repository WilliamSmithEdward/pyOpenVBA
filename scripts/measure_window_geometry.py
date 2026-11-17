"""Window geometry: Application and Window Width, Height, Left, Top, UsableWidth and UsableHeight, Window.VisibleRange,
LargeScroll, ScrollIntoView, PointsToScreenPixelsX/Y and RangeFromPoint, read with Excel shown at a known size.

Each case runs on a new workbook. Excel is shown with ScreenUpdating on, the application window is set normal at
800 by 600 points at (20, 20) and the workbook's window normal at 700 by 500 at (0, 0); then the case's statements
run and the reads are taken. Afterwards the window state, size and place, DisplayFormulaBar and DisplayStatusBar
are put back exactly as they were, because Excel saves its window placement when it exits and the next session opens
with it. They are put back through ActiveWindow: Application.Top reads a point less than the window's Top, while
setting it sets the window's, so a restore through Application moves the window up a point each time. StandardHeight is read at the start and end of every case: geometry depends on the
display's DPI, and a case whose two readings differ is not to be trusted.

    python scripts/measure_window_geometry.py [_a]   # a development look

writes tests/fixtures/window_geometry<look>.json. Window.Hwnd and Window.PrintPreview are not probed.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests/fixtures'

#: What every case reads.
READS = ['Application.Width', 'Application.Height', 'Application.Left', 'Application.Top', 'Application.UsableWidth',
         'Application.UsableHeight', 'w.Width', 'w.Height', 'w.Left', 'w.Top', 'w.UsableWidth', 'w.UsableHeight',
         'w.VisibleRange.Address(False, False)', 'w.ScrollRow', 'w.ScrollColumn', 'w.Zoom',
         'w.PointsToScreenPixelsX(0)', 'w.PointsToScreenPixelsY(0)', 'w.PointsToScreenPixelsX(100)',
         'w.PointsToScreenPixelsY(100)']
POINTS = ['w.RangeFromPoint(w.PointsToScreenPixelsX(0) + 5, w.PointsToScreenPixelsY(0) + 5).Address(False, False)',
          'w.RangeFromPoint(w.PointsToScreenPixelsX(100), w.PointsToScreenPixelsY(100)).Address(False, False)',
          'TypeName(w.RangeFromPoint(0, 0))', 'TypeName(w.RangeFromPoint(w.PointsToScreenPixelsX(0) - 5, '
          'w.PointsToScreenPixelsY(0) - 5))']

CASES_A: list[tuple[str, list[str], list[str]]] = [
    ('base', [], READS + POINTS),
    ('headings off', ['w.DisplayHeadings = False'], READS),
    ('horizontal bar off', ['w.DisplayHorizontalScrollBar = False'], READS),
    ('vertical bar off', ['w.DisplayVerticalScrollBar = False'], READS),
    ('tabs off', ['w.DisplayWorkbookTabs = False'], READS),
    ('bars and tabs off', ['w.DisplayHorizontalScrollBar = False', 'w.DisplayVerticalScrollBar = False',
                           'w.DisplayWorkbookTabs = False'], READS),
    ('formula bar off', ['Application.DisplayFormulaBar = False'], READS),
    ('status bar off', ['Application.DisplayStatusBar = False'], READS),
    ('zoom 50', ['w.Zoom = 50'], READS),
    ('zoom 200', ['w.Zoom = 200'], READS),
    ('zoom 85', ['w.Zoom = 85'], READS),
    ('scrolled', ['w.ScrollRow = 10', 'w.ScrollColumn = 3'], READS + POINTS),
    ('wide column tall row', ['ActiveSheet.Columns("C").ColumnWidth = 30', 'ActiveSheet.Rows(5).RowHeight = 40'], READS),
    ('hidden rows', ['ActiveSheet.Rows("3:6").Hidden = True'], READS),
    ('frozen', ['ActiveSheet.Range("C4").Select', 'w.FreezePanes = True'], READS),
    ('frozen scrolled', ['ActiveSheet.Range("C4").Select', 'w.FreezePanes = True', 'w.ScrollRow = 20'], READS),
    ('window smaller', ['w.Width = 400', 'w.Height = 300'], READS),
    ('window maximized', ['w.WindowState = xlMaximized'], READS),
    ('application smaller', ['Application.Width = 600', 'Application.Height = 450'], READS),
    ('application moved', ['Application.Left = 60', 'Application.Top = 40'], READS),
    ('large scroll down', ['w.LargeScroll Down:=1'], ['w.ScrollRow', 'w.ScrollColumn', 'w.VisibleRange.Address(False, False)']),
    ('large scroll down 2 right 1', ['w.LargeScroll Down:=2, ToRight:=1'],
     ['w.ScrollRow', 'w.ScrollColumn', 'w.VisibleRange.Address(False, False)']),
    ('large scroll up', ['w.ScrollRow = 60', 'w.LargeScroll Up:=1'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('large scroll left at start', ['w.LargeScroll ToLeft:=1'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('large scroll answer', ['v = w.LargeScroll(Down:=1)'], ['TypeName(v)', 'w.ScrollRow']),
    ('large scroll frozen', ['ActiveSheet.Range("C4").Select', 'w.FreezePanes = True', 'w.LargeScroll Down:=1'],
     ['w.ScrollRow', 'w.ScrollColumn', 'w.VisibleRange.Address(False, False)']),
    ('scroll into view near', ['w.ScrollIntoView 0, 200, 10, 10'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('scroll into view far', ['w.ScrollIntoView 0, 1000, 10, 10'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('scroll into view far start', ['w.ScrollIntoView 0, 1000, 10, 10, True'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('scroll into view far end', ['w.ScrollIntoView 0, 1000, 10, 10, False'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('scroll into view right', ['w.ScrollIntoView 1500, 0, 10, 10'], ['w.ScrollRow', 'w.ScrollColumn']),
    ('scroll into view back', ['w.ScrollRow = 80', 'w.ScrollIntoView 0, 100, 10, 10'], ['w.ScrollRow', 'w.ScrollColumn']),
]


#: Each case's sheet dimensions in points, for the rules that turn a usable area into cells.
SIZES = ['ActiveSheet.Columns(1).Width', 'ActiveSheet.Rows(1).Height']
GRID = ['w.UsableWidth', 'w.UsableHeight', 'w.VisibleRange.Address(False, False)', 'w.ScrollRow', 'w.ScrollColumn']


def _sized(width: int, height: int, zoom: int) -> tuple[str, list[str], list[str]]:
    return (f'window {width} by {height} zoom {zoom}', [f'w.Width = {width}', f'w.Height = {height}', f'w.Zoom = {zoom}'],
            SIZES + GRID + ['w.Width', 'w.Height'])


#: The second look: window sizes and zooms, scroll targets, frozen splits and odd rows and columns, to pin the rules
#: that turn a usable area into cells.
CASES_B: list[tuple[str, list[str], list[str]]] = [
    _sized(300, 200, 100), _sized(450, 330, 100), _sized(610, 410, 100), _sized(900, 640, 100),
    _sized(700, 500, 75), _sized(700, 500, 110), _sized(700, 500, 150), _sized(520, 360, 130), _sized(380, 260, 60),
    ('tall rows', ['ActiveSheet.Rows("1:40").RowHeight = 33'], SIZES + GRID),
    ('narrow columns', ['ActiveSheet.Columns("A:Z").ColumnWidth = 3'], SIZES + GRID),
    ('hidden columns', ['ActiveSheet.Columns("B:D").Hidden = True'], SIZES + GRID),
    ('frozen at B2', ['ActiveSheet.Range("B2").Select', 'w.FreezePanes = True'], SIZES + GRID),
    ('frozen at F10', ['ActiveSheet.Range("F10").Select', 'w.FreezePanes = True'], SIZES + GRID),
    ('frozen rows only', ['ActiveSheet.Range("A6").Select', 'w.FreezePanes = True', 'w.LargeScroll Down:=1'],
     SIZES + GRID),
    ('large scroll tall rows', ['ActiveSheet.Rows("1:200").RowHeight = 33', 'w.LargeScroll Down:=1'], SIZES + GRID),
    ('large scroll right 2', ['w.LargeScroll ToRight:=2'], SIZES + GRID),
    ('large scroll down 3 up 1', ['w.LargeScroll Down:=3, Up:=1'], SIZES + GRID),
    ('large scroll near the end', ['w.ScrollRow = 1048570', 'w.LargeScroll Down:=1'], ['w.ScrollRow']),
    ('large scroll up past top', ['w.ScrollRow = 5', 'w.LargeScroll Up:=1'], ['w.ScrollRow']),
    ('into view middle', ['w.ScrollIntoView 300, 400, 50, 50'], SIZES + GRID),
    ('into view middle start', ['w.ScrollIntoView 300, 400, 50, 50, True'], SIZES + GRID),
    ('into view middle end', ['w.ScrollIntoView 300, 400, 50, 50, False'], SIZES + GRID),
    ('into view taller than window', ['w.ScrollIntoView 0, 600, 10, 900'], SIZES + GRID),
    ('into view already shown', ['w.ScrollIntoView 100, 100, 20, 20'], SIZES + GRID),
    ('into view far right end', ['w.ScrollIntoView 2000, 0, 30, 10, False'], SIZES + GRID),
    ('into view tall rows', ['ActiveSheet.Rows("1:200").RowHeight = 33', 'w.ScrollIntoView 0, 1500, 10, 10'],
     SIZES + GRID),
    ('into view frozen', ['ActiveSheet.Range("C4").Select', 'w.FreezePanes = True', 'w.ScrollIntoView 0, 1000, 10, 10'],
     SIZES + GRID),
    ('into view negative', ['w.ScrollRow = 30', 'w.ScrollIntoView -50, -50, 10, 10'], SIZES + GRID),
    ('into view zoom 50', ['w.Zoom = 50', 'w.ScrollIntoView 0, 2000, 10, 10'], SIZES + GRID),
]


def source(cases: list[tuple[str, list[str], list[str]]]) -> str:
    lines = ['Public Function Go(ByVal which As Long) As String',
             '    Dim wb As Workbook, s As String, state As Long, wide As Double, high As Double, across As Double',
             '    Dim down As Double, formulas As Boolean, status As Boolean',
             '    Application.Visible = True', '    Application.ScreenUpdating = True',
             '    state = Application.WindowState', '    Application.WindowState = xlNormal',
             '    wide = ActiveWindow.Width: high = ActiveWindow.Height',
             '    across = ActiveWindow.Left: down = ActiveWindow.Top',
             '    formulas = Application.DisplayFormulaBar: status = Application.DisplayStatusBar',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    wb.Activate',
             '    s = "dpi=" & wb.Worksheets(1).StandardHeight',
             '    Application.Width = 800: Application.Height = 600: Application.Left = 20: Application.Top = 20',
             '    wb.Windows(1).WindowState = xlNormal',
             '    wb.Windows(1).Left = 0: wb.Windows(1).Top = 0: wb.Windows(1).Width = 700: wb.Windows(1).Height = 500',
             '    On Error Resume Next', '    Select Case which']
    lines += [f'    Case {number}: s = s & Case{number}(wb.Windows(1))' for number in range(len(cases))]
    lines += ['    End Select', '    Err.Clear', '    s = s & " dpi=" & wb.Worksheets(1).StandardHeight',
              '    wb.Close False', '    Application.DisplayFormulaBar = formulas',
              '    Application.DisplayStatusBar = status', '    Application.WindowState = xlNormal',
              '    ActiveWindow.Width = wide: ActiveWindow.Height = high',
              '    ActiveWindow.Left = across: ActiveWindow.Top = down', '    Application.WindowState = state',
              '    Go = s', 'End Function']
    for number, (_name, statements, reads) in enumerate(cases):
        lines += ['', f'Private Function Case{number}(w As Object) As String',
                  '    Dim v As Variant, r As String', '    On Error Resume Next']
        for statement in statements:
            lines += ['    Err.Clear', f'    {statement}',
                      '    If Err.Number <> 0 Then r = r & "[" & Err.Number & ":" & Err.Description & "]"']
        for read in reads:
            lines += ['    Err.Clear', '    v = Empty', f'    v = {read}',
                      '    If Err.Number <> 0 Then r = r & "|E" & Err.Number '
                      'Else r = r & "|" & TypeName(v) & ":" & CStr(v)']
        lines += [f'    Case{number} = r', 'End Function']
    return '\n'.join(lines) + '\n'


def placement() -> str:
    """Excel's saved window placement, read after Excel has exited."""
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office\16.0\Excel\Options') as key:
        return f'Pos={winreg.QueryValueEx(key, "Pos")[0]} Maximized={winreg.QueryValueEx(key, "Maximized")[0]}'


def held_out(round_: str) -> list[dict[str, Any]]:
    """Fresh cases for a round: a window size, a measured zoom, headings and bars on or off, rows made taller, panes
    frozen or not, a scroll, then nothing, LargeScroll or ScrollIntoView; each with the layout the rules read it as."""
    import random

    pick = random.Random('window_geometry' + round_)
    found: list[dict[str, Any]] = []
    for number in range(40):
        width, height = pick.randint(300, 900), pick.randint(220, 650)
        zoom = 100 if round_ not in ('1',) else pick.choice([50, 60, 75, 100, 100, 110, 130, 150, 200])
        spec: dict[str, Any] = {'width': width, 'height': height, 'zoom': zoom,
                                'headings': pick.random() > 0.2, 'vertical_bar': pick.random() > 0.2,
                                'horizontal_bar': pick.random() > 0.2, 'tabs': pick.random() > 0.2,
                                'tall': pick.choice([0, 0, 0, 22, 40]), 'frozen': [0, 0], 'scroll': [1, 1]}
        statements = [f'w.Width = {width}', f'w.Height = {height}', f'w.Zoom = {zoom}']
        statements += [f'w.{name} = False' for name, key in (('DisplayHeadings', 'headings'),
                       ('DisplayVerticalScrollBar', 'vertical_bar'), ('DisplayHorizontalScrollBar', 'horizontal_bar'),
                       ('DisplayWorkbookTabs', 'tabs')) if not spec[key]]
        if spec['tall']:
            statements.append(f'ActiveSheet.Rows("1:600").RowHeight = {spec["tall"]}')
        if pick.random() < 0.3:
            row, column = pick.randint(1, 6), pick.randint(1, 4)
            if row > 1 or column > 1:
                spec['frozen'], spec['scroll'] = [row - 1, column - 1], [row, column]
                statements += [f'ActiveSheet.Cells({row}, {column}).Select', 'w.FreezePanes = True']
        if pick.random() < 0.5:
            spec['scroll'] = [spec['scroll'][0] + pick.randint(0, 150), spec['scroll'][1] + pick.randint(0, 20)]
            statements += [f'w.ScrollRow = {spec["scroll"][0]}', f'w.ScrollColumn = {spec["scroll"][1]}']
        action = pick.choice(['none', 'large', 'large', 'into', 'into'])
        if action == 'large':
            down, right = pick.randint(-2, 3), pick.randint(-1, 2)
            spec['action'] = ['large', down, right]
            parts = [f'Down:={down}' if down > 0 else f'Up:={-down}' if down < 0 else '',
                     f'ToRight:={right}' if right > 0 else f'ToLeft:={-right}' if right < 0 else '']
            statements.append('w.LargeScroll ' + ', '.join(one for one in parts if one) if any(parts)
                              else 'w.LargeScroll')
        elif action == 'into':
            rectangle = [pick.randint(-20, 3000), pick.randint(-20, 4000), pick.randint(1, 300), pick.randint(1, 300)]
            plain = spec['scroll'] == [1, 1] and spec['frozen'] == [0, 0]
            start = pick.choice(['', ', True', ', False'] if plain and round_ in ('1', '2') else ['', ', True'])
            spec['action'] = ['into', rectangle, start != ', False']
            statements.append('w.ScrollIntoView ' + ', '.join(str(one) for one in rectangle) + start)
        found.append({'call': [f'held {round_}.{number}', statements, SIZES + GRID], 'spec': spec})
    return found


def predicted(spec: dict[str, Any]) -> str:
    """What the rules say the case reads, spelled as the probe writes it, at StandardHeight 14.5."""
    from dataclasses import replace as replaced

    from pyopenvba._a1 import column_letter
    from pyopenvba.apps.excel._window_geometry import DISPLAYS, Layout, large_scroll, scroll_into_view, usable, visible

    display = DISPLAYS[14.5]
    tall = spec['tall']
    layout = Layout(width=spec['width'], height=spec['height'], zoom=spec['zoom'], headings=spec['headings'],
                    vertical_bar=spec['vertical_bar'], horizontal_bar=spec['horizontal_bar'], tabs=spec['tabs'],
                    scroll=(spec['scroll'][0], spec['scroll'][1]), frozen=(spec['frozen'][0], spec['frozen'][1]),
                    row_pixels=(lambda row: int(tall * 2) if row <= 600 else 29) if tall else (lambda _row: 29))
    action = spec.get('action')
    if action and action[0] == 'large':
        layout = replaced(layout, scroll=large_scroll(display, layout, action[1], action[2]))
    elif action:
        layout = replaced(layout, scroll=scroll_into_view(display, layout, tuple(action[1]), action[2]))
    width, height = usable(display, layout)
    top, left, bottom, right = visible(display, layout)

    def number(value: float) -> str:
        return 'Double:' + (str(int(value)) if value == int(value) else str(value))

    return ('dpi=14.5|' + '|'.join([number(48), number(tall or 14.5), number(width), number(height),
                                     f'String:{column_letter(left)}{top}:{column_letter(right)}{bottom}',
                                     number(layout.scroll[0]), number(layout.scroll[1])]) + ' dpi=14.5')


def predict(round_: str) -> None:
    cases = [{'call': one['call'], 'spec': one['spec'], 'model': predicted(one['spec'])} for one in held_out(round_)]
    (FIXTURES / f'window_geometry{round_}_predictions.json').write_text(json.dumps({'cases': cases}, indent=1) + '\n',
                                                                       encoding='utf-8', newline='\n')
    print(len(cases), 'predicted')


REPAIR = '''Public Function Repair(ByVal across As Double, ByVal down As Double, ByVal wide As Double, _
        ByVal high As Double) As String
    Application.WindowState = xlNormal
    ActiveWindow.Left = ActiveWindow.Left + across: ActiveWindow.Top = ActiveWindow.Top + down
    ActiveWindow.Width = ActiveWindow.Width + wide: ActiveWindow.Height = ActiveWindow.Height + high
    Repair = ActiveWindow.Left & "," & ActiveWindow.Top & "," & ActiveWindow.Width & "," & ActiveWindow.Height
End Function
'''


def repaired(before: str, after: str) -> str:
    """Put Excel's saved placement back where it was, through ActiveWindow, when a run moved it: points read back a
    pixel off at some sizes, so a restore through them can land a pixel away. Pos is in screen pixels, and a point is
    as many pixels as StandardHeight's rows say (2 at 144 DPI, 4/3 at 96)."""
    from pyvbaharness import ExcelSession, HarnessConfig

    def numbers(text: str) -> list[int]:
        return [int(one) for one in text.split()[0].removeprefix('Pos=').split(',')]

    for _ in range(3):
        if after == before:
            break
        old, new = numbers(before), numbers(after)
        if before.split()[1] != after.split()[1]:
            raise SystemExit(f'Maximized moved: {before} -> {after}; put it back by hand')
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            scale = excel.run_vba('Public Function Dpi() As Double\n    Dpi = ActiveSheet.StandardHeight\nEnd Function\n',
                                  'Dpi', (), timeout=60.0, line_numbers=False)
            pixels = 2.0 if scale.value == 14.5 else 4.0 / 3.0
            change = tuple((one - two) / pixels for one, two in zip(old, new, strict=True))
            excel.run_vba(REPAIR, 'Repair', change, timeout=60.0, line_numbers=False)
        after = placement()
    return after


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    look = sys.argv[1] if len(sys.argv) > 1 else '_a'
    predictions = FIXTURES / f'window_geometry{look}_predictions.json'
    expected = json.loads(predictions.read_text('utf-8'))['cases'] if predictions.exists() else None
    cases = ([tuple(one['call']) for one in expected] if expected is not None
             else {'_a': CASES_A, '_b': CASES_B}[look])
    text = source(cases)
    before = placement()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for number in range(len(cases)):
            result = excel.run_vba(text, 'Go', (number,), timeout=120.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs
                           else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
    after = placement()
    (FIXTURES / f'window_geometry{look}.json').write_text(json.dumps({
        'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text,
        'cases': [[*case, answer] for case, answer in zip(cases, answers, strict=True)]},
        indent=1) + '\n', encoding='utf-8', newline='\n')
    print(len(answers), 'asked,', sum(answer.startswith('ERROR') for answer in answers), 'errors')
    print('placement before', before, 'after', after)
    if expected is not None:
        wrong = [(one['call'][0], one['model'], answer) for one, answer in zip(expected, answers, strict=True)
                 if one['model'] != answer]
        print(f'{len(answers) - len(wrong)} right, {len(wrong)} wrong of {len(answers)}')
        for one in wrong[:12]:
            print(one)
    if before != after:
        after = repaired(before, after)
        print('placement repaired to', after)
        if after != before:
            print('PLACEMENT STILL MOVED: put it back through ActiveWindow before anyone else runs', file=sys.stderr)
            sys.exit(2)


def model_answer(cases: list[Any], number: int) -> str | None:
    """Not built yet: the model's answer, for the replay once it is."""
    raise NotImplementedError((cases, number))


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
