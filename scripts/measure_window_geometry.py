"""Window geometry: Application and Window Width, Height, Left, Top, UsableWidth and UsableHeight, Window.VisibleRange,
LargeScroll, ScrollIntoView, PointsToScreenPixelsX/Y and RangeFromPoint, read with Excel shown at a known size.

Each case runs on a new workbook. Excel is shown with ScreenUpdating on, the application window is set normal at
800 by 600 points at (20, 20) and the workbook's window normal at 700 by 500 at (0, 0); then the case's statements
run and the reads are taken. Afterwards the window state, size and place, DisplayFormulaBar and DisplayStatusBar
are put back exactly as they were, through ActiveWindow: Application.Top reads a point less than the window's Top
while setting it sets the window's, so a restore through Application moves the window up a point each time, because Excel saves its window placement when it exits and the
next session opens with it. StandardHeight is read at the start and end of every case: geometry depends on the
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


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    look = sys.argv[1] if len(sys.argv) > 1 else '_a'
    cases = {'_a': CASES_A}[look]
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


def model_answer(cases: list[Any], number: int) -> str | None:
    """Not built yet: the model's answer, for the replay once it is."""
    raise NotImplementedError((cases, number))


if __name__ == '__main__':
    main()
