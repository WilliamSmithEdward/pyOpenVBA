"""What a cell shift does to merged cells it cuts through, with DisplayAlerts on: Excel's prompt, each of its answers, and after.

Nobody is at the screen, so the harness's worker runs here with its watcher replaced by a driver, as in
scripts/measure_dialogs.py. The driver only touches windows of the Excel the worker owns. For each of Excel's own alerts it logs
the text, the buttons and the default button, then presses the button the case asks for (a file the probe writes before each
case: "ok", "cancel" or "default", which presses Return), so the probe sees what each answer does.

Each case builds a new workbook: numbers in A1:F12, a merged area (its value kept in its top-left cell), DisplayAlerts on or
off, then one Insert or Delete with a Shift. It reads the error, every merged area in A1:H16, the cells A1:H16 and the alert
the driver logged, then saves the book as .xlsx, opens it again and reads the merged areas and cells once more.

    python scripts/measure_merge_shifts.py            # the first look: tests/fixtures/merge_shifts.json
    python scripts/measure_merge_shifts.py predict 2  # tests/fixtures/merge_shifts2_predictions.json, from the model
    python scripts/measure_merge_shifts.py measure 2  # asks Excel those and scores them: tests/fixtures/merge_shifts2.json
"""
from __future__ import annotations

import json
import os
import random
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures'
SOURCE = r'''Private Function Merges(ws As Object) As String
    Dim c As Object, out As String
    For Each c In ws.Range("A1:H16").Cells
        If c.MergeCells Then
            If c.Address = c.MergeArea.Cells(1, 1).Address Then out = out & c.MergeArea.Address(False, False) & ";"
        End If
    Next
    Merges = out
End Function

Private Function Cells16(ws As Object) As String
    Dim c As Object, out As String
    For Each c In ws.Range("A1:H16").Cells
        out = out & c.Formula & ","
    Next
    Cells16 = out
End Function

Public Function Go(ByVal merged As String, ByVal alerts As Boolean, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, i As Long, j As Long, out As String, path As String, one As Variant
    Dim parts() As String, target As String, how As String
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wt = ws
    For i = 1 To 12
        For j = 1 To 6
            ws.Cells(i, j).Value = i * 10 + j
        Next
    Next
    Application.DisplayAlerts = False
    For Each one In Split(merged, ";")
        If one <> "" Then ws.Range(one).Merge
    Next
    Application.DisplayAlerts = alerts
    parts = Split(edit, " ")
    target = parts(1)
    how = parts(2)
    On Error Resume Next
    If typed Then
        If parts(0) = "insert" Then
            wt.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            wt.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    Else
        If parts(0) = "insert" Then
            ws.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            ws.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    End If
    out = "err=" & Err.Number & ":" & Err.Description & " alerts=" & Application.DisplayAlerts & " merges=" & Merges(ws) & " cells=" & Cells16(ws)
    Err.Clear
    Application.DisplayAlerts = False
    path = Environ("TEMP") & "\pyovba_merge_shift.xlsx"
    Kill path
    Err.Clear
    wb.SaveAs path, 51
    wb.Close False
    Set wb = Workbooks.Open(path)
    out = out & " ~ merges=" & Merges(wb.Worksheets(1)) & " cells=" & Cells16(wb.Worksheets(1))
    wb.Close False
    Kill path
    Application.DisplayAlerts = True
    Go = out
End Function
'''


def _drive() -> None:
    """Run as the harness's worker, with the driver in the watcher's place."""
    import ctypes
    from ctypes import wintypes as wt

    from pyvbaharness.worker import __main__ as worker
    from pyvbaharness.worker import watcher

    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.SendMessageTimeoutW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM, wt.UINT, wt.UINT, ctypes.POINTER(ctypes.c_size_t)]
    user32.SendMessageTimeoutW.restype = ctypes.c_size_t
    user32.PostMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
    log, action = Path(os.environ['MERGE_LOG']), Path(os.environ['MERGE_ACTION'])
    DM_GETDEFID, BM_CLICK, WM_KEYDOWN, WM_KEYUP, VK_RETURN = 0x0400, 0x00F5, 0x0100, 0x0101, 0x0D

    def mine(hwnd: int, pid: int) -> bool:
        return watcher._window_pid(hwnd) == pid

    def children(hwnd: int) -> list[int]:
        found: list[int] = []
        proc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)(lambda child, _: found.append(child) or True)
        user32.EnumChildWindows(hwnd, proc, 0)
        return found

    def send(hwnd: int, pid: int, message: int) -> int:
        result = ctypes.c_size_t(0)
        if mine(hwnd, pid):
            user32.SendMessageTimeoutW(hwnd, message, 0, 0, 2, 1000, ctypes.byref(result))
        return result.value

    answered: set[int] = set()

    def answer_office_form(hwnd: int) -> None:
        import comtypes
        import comtypes.client

        comtypes.CoInitialize()
        comtypes.client.GetModule('UIAutomationCore.dll')
        from comtypes.gen.UIAutomationClient import (CUIAutomation, IUIAutomation, IUIAutomationInvokePattern,
                                                     IUIAutomationLegacyIAccessiblePattern)

        if hwnd in answered:
            return
        automation = comtypes.client.CreateObject(CUIAutomation, interface=IUIAutomation)
        found = automation.ElementFromHandle(hwnd).FindAll(4, automation.CreateTrueCondition())  # TreeScope_Descendants
        texts: list[str] = []
        buttons: list[tuple[str, bool, object]] = []
        for index in range(found.Length):
            element = found.GetElement(index)
            name, control = element.CurrentName or '', element.CurrentControlType
            if control == 50000:  # a button; the default one carries STATE_SYSTEM_DEFAULT, or has the focus as the form opens
                legacy = element.GetCurrentPattern(10018).QueryInterface(IUIAutomationLegacyIAccessiblePattern)
                buttons.append((name, bool(legacy.CurrentState & 0x100) or bool(element.CurrentHasKeyboardFocus), element))
            elif name and name not in texts:
                texts.append(name)
        wanted = action.read_text(encoding='ascii').strip()
        chosen = next((one for one in buttons if (one[1] if wanted == 'default' else one[0].lower() == wanted)), None)
        with log.open('a', encoding='utf-8') as file:
            file.write(' / '.join(text.replace('\r', ' ').replace('\n', ' ') for text in texts)
                       + f' | buttons={[(name, default) for name, default, _ in buttons]} pressed={wanted}:{chosen[0] if chosen else None}\n')
        if chosen is not None:
            answered.add(hwnd)
            chosen[2].GetCurrentPattern(10000).QueryInterface(IUIAutomationInvokePattern).Invoke()

    def scan(self: watcher.DialogWatcher) -> None:
        for hwnd in watcher._enum_top_level(self.pid):
            kind = watcher._window_class(hwnd)
            if not user32.IsWindowVisible(hwnd) or kind in ('XLMAIN', 'XLDESK', 'EXCEL7'):
                continue
            if kind == 'NUIDialog' and mine(hwnd, self.pid):
                # Office's own message form draws its text and buttons, so UI Automation reads and presses them.
                answer_office_form(hwnd)
                continue
            if kind != '#32770':
                with log.open('a', encoding='utf-8') as file:
                    file.write(f'other {kind} {watcher._window_text(hwnd)!r}\n')
                continue
            kids = children(hwnd)
            texts = [watcher._window_text(kid) for kid in kids if watcher._window_text(kid)]
            if 'ShellView' in texts:
                continue
            buttons = [(user32.GetDlgCtrlID(kid), watcher._window_text(kid)) for kid in kids if 'Button' in watcher._window_class(kid)]
            default = send(hwnd, self.pid, DM_GETDEFID) & 0xFFFF
            wanted = action.read_text(encoding='ascii').strip()
            with log.open('a', encoding='utf-8') as file:
                file.write(' / '.join(text.replace('\r', ' ').replace('\n', ' ') for text in texts) + f' | buttons={buttons} default={default} pressed={wanted}\n')
            if wanted == 'default':
                user32.PostMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0)
                user32.PostMessageW(hwnd, WM_KEYUP, VK_RETURN, 0)
                continue
            control = 1 if wanted == 'ok' else 2
            for kid in kids:
                if user32.GetDlgCtrlID(kid) == control and 'Button' in watcher._window_class(kid):
                    send(kid, self.pid, BM_CLICK)
                    break

    watcher.DialogWatcher._scan = scan  # type: ignore[method-assign]
    sys.exit(worker.main())


#: A merged area A3:B4 (or one other), cut or carried by each kind of shift, with each answer and with alerts off.
MERGES = ('A3:B4', 'B5:C6', 'A3:A5')
EDITS = ('insert A3:A4 down', 'insert A3 down', 'insert A2:B2 down', 'insert A1:F1 down', 'delete A3 up', 'delete A2:B2 up', 'delete A3:B4 up',
         'insert A3 right', 'insert A3:A4 right', 'delete A3 left', 'delete A3:A4 left', 'insert C3 down', 'insert B4 down', 'delete A5 up',
         'insert A10 down', 'insert H3 right')


def first_look() -> list[list[object]]:
    cases: list[list[object]] = []
    for merged in MERGES:
        for edit in EDITS:
            for alerts, answer in ((True, 'ok'), (True, 'cancel'), (False, 'ok')):
                cases.append([merged, alerts, edit, True, answer])
    return cases


def held_out(round_: str) -> list[list[object]]:
    chooser = random.Random('merge-shifts' + round_)
    columns = 'ABCDEF'
    cases: list[list[object]] = []
    while len(cases) < 160:
        left, top = chooser.randrange(5), chooser.randint(1, 9)
        merged = f'{columns[left]}{top}:{columns[left + chooser.randint(0, 1)]}{top + chooser.randint(0, 2)}'
        if merged.split(':')[0] == merged.split(':')[1]:
            continue
        el, et = chooser.randrange(6), chooser.randint(1, 11)
        target = f'{columns[el]}{et}:{columns[min(5, el + chooser.choice((0, 0, 1, 2)))]}{et + chooser.choice((0, 0, 1, 2))}'
        verb = chooser.choice(('insert', 'delete'))
        how = chooser.choice(('down', 'right')) if verb == 'insert' else chooser.choice(('up', 'left'))
        alerts = chooser.random() < .8
        answer = chooser.choice(('ok', 'cancel')) if alerts else 'ok'
        cases.append([merged, alerts, f'{verb} {target} {how}', chooser.random() < .5, answer])
    return cases


def _ask(cases: list[list[object]]) -> tuple[list[str], list[str]]:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: list[str] = []
    alerts: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        log, action = Path(temp) / 'driver.log', Path(temp) / 'action.txt'
        log.touch()
        action.write_text('ok', encoding='ascii')
        config = HarnessConfig(exclusive=False, worker_argv=['uv', 'run', '--no-project', '--python', sys.executable, '--with', 'comtypes==1.4.8', '--with', 'pywin32==311', 'python', str(Path(__file__).resolve()), '--worker'],
                               extra_env={'MERGE_LOG': str(log), 'MERGE_ACTION': str(action)})
        with ExcelSession(config) as excel:
            excel.new_document()
            for case in cases:
                action.write_text(str(case[4]), encoding='ascii')
                start = log.stat().st_size
                result = excel.run_vba(SOURCE, 'Go', tuple(case[:4]), timeout=60.0, line_numbers=False)
                answers.append(str(result.value) if result.ok else f'ERROR {str(result)[:200]}')
                with log.open(encoding='utf-8') as file:
                    file.seek(start)
                    alerts.append(file.read().strip())
    return answers, alerts


def _write(path: Path, cases: list[list[object]], answers: list[str], alerts: list[str]) -> None:
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases, 'answers': answers,
                                'alerts': alerts}, indent=0), encoding='utf-8', newline='\n')


def predict(round_: str) -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.apps.excel._merges import UNMERGE_PROMPT
    from pyopenvba.exceptions import VBAUnsupportedError

    found = []
    for case in held_out(round_):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(SOURCE, name='Probe')
        app.answer_dialogs(2 if case[4] == 'cancel' else 1)
        try:
            model: str | None = str(app.run('Probe.Go', *case[:4]))
        except VBAUnsupportedError:
            model = None
        asked = any(one[1] == UNMERGE_PROMPT for one in app.interpreter.dialogs)  # type: ignore[union-attr]
        found.append({'call': case, 'model': model, 'asks': asked})
    (FIXTURES / f'merge_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused,', sum(one['asks'] for one in found), 'asking')


def measure(round_: str) -> None:
    predicted = json.loads((FIXTURES / f'merge_shifts{round_}_predictions.json').read_text('utf-8'))['cases']
    cases = [one['call'] for one in predicted]
    answers, alerts = _ask(cases)
    _write(FIXTURES / f'merge_shifts{round_}.json', cases, answers, alerts)
    wrong = [(one['call'], one['model'], answer) for one, answer, alert in zip(predicted, answers, alerts)
             if one['model'] is not None and (one['model'] != answer or one['asks'] != ('unmerge' in alert))]
    refused = sum(one['model'] is None for one in predicted)
    print(f'{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}')
    for one in wrong[:20]:
        print(str(one)[:500])


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] in ('predict', 'measure'):
        {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
    elif len(sys.argv) > 1 and sys.argv[1] == '--worker':
        del sys.argv[1]
        _drive()
    elif len(sys.argv) > 2 and sys.argv[1] == 'try':
        # Ask some first-look cases by index and print what comes back, writing nothing.
        picked = [first_look()[int(one)] for one in sys.argv[2:]]
        for case, answer, alert in zip(picked, *_ask(picked)):
            print(case, '\n   ', answer[:400], '\n    alert:', alert)
    elif len(sys.argv) == 1:
        found = first_look()
        _write(FIXTURES / 'merge_shifts.json', found, *_ask(found))
        print(len(found), 'measured')
