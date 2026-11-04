"""What Excel's dialogs take and give back: InputBox, Application.InputBox, GetOpenFilename, GetSaveAsFilename, FileDialog, Dialogs.

Nobody is at the screen, so the harness's worker runs here with its watcher replaced by a driver. The driver only touches
windows of the Excel the worker owns, found by process id and checked again before each message, and uses no foreground
window, keyboard or mouse. A dialog says what to do with it in its own title: `drive|cancel`, `drive|ok`, or
`drive|ok|text` to fill in the answer first. Excel's own alerts get OK, and their text is kept with the case.

    python scripts/measure_dialogs.py

writes tests/fixtures/dialogs.json, which tests/test_excel_dialogs.py replays. A run stops after about eight minutes and the
next takes up where it stopped; delete the fixture to measure everything again.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/dialogs.json'


def _drive() -> None:
    """Run as the harness's worker, with the driver in the watcher's place."""
    import ctypes
    from ctypes import wintypes as wt

    from pyvbaharness.worker import __main__ as worker
    from pyvbaharness.worker import watcher

    user32 = ctypes.WinDLL('user32', use_last_error=True)
    user32.SendMessageTimeoutW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPCWSTR, wt.UINT, wt.UINT, ctypes.POINTER(ctypes.c_size_t)]
    user32.SendMessageTimeoutW.restype = ctypes.c_size_t
    user32.PostMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
    log = Path(os.environ['DIALOG_LOG'])
    WM_SETTEXT, WM_CLOSE, WM_KEYDOWN, WM_KEYUP, BM_CLICK, VK_RETURN, VK_ESCAPE = 0x000C, 0x0010, 0x0100, 0x0101, 0x00F5, 0x0D, 0x1B
    handled: set[int] = set()

    def note(line: str) -> None:
        with log.open('a', encoding='utf-8') as file:
            file.write(line.replace('\r', ' ').replace('\n', ' ') + '\n')

    def mine(hwnd: int, pid: int) -> bool:
        return watcher._window_pid(hwnd) == pid

    def children(hwnd: int) -> list[int]:
        found: list[int] = []
        proc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)(lambda child, _: found.append(child) or True)
        user32.EnumChildWindows(hwnd, proc, 0)
        return found

    def send(hwnd: int, pid: int, message: int, text: str | None = None) -> None:
        if mine(hwnd, pid):
            result = ctypes.c_size_t(0)
            user32.SendMessageTimeoutW(hwnd, message, 0, text, 2, 1000, ctypes.byref(result))

    def press(hwnd: int, pid: int, kids: list[int], control: int, key: int) -> None:
        for kid in kids:
            if user32.GetDlgCtrlID(kid) == control and 'Button' in watcher._window_class(kid):
                send(kid, pid, BM_CLICK)
                return
        if mine(hwnd, pid):  # Excel's own forms draw their buttons, so they get the key
            user32.PostMessageW(hwnd, WM_KEYDOWN, key, 0)
            user32.PostMessageW(hwnd, WM_KEYUP, key, 0)

    def scan(self: watcher.DialogWatcher) -> None:
        for hwnd in watcher._enum_top_level(self.pid):
            kind = watcher._window_class(hwnd)
            if not user32.IsWindowVisible(hwnd) or not (kind == '#32770' or kind.startswith('bosa_sdm_') or kind in ('NUIDialog', 'MsoDialog')):
                continue
            title = watcher._window_text(hwnd)
            kids = children(hwnd)
            if title.startswith('drive|'):
                _, action, *text = title.split('|', 2)
                if hwnd in handled:  # the answer was refused and the form is still up
                    action, text = 'cancel', []
                handled.add(hwnd)
                note('form ' + kind + ' ' + ' '.join(sorted({watcher._window_class(kid) for kid in kids})))
                if text:
                    edits = [kid for kid in kids if watcher._window_class(kid) in ('Edit', 'EDTBX')]
                    if edits:
                        send(edits[0], self.pid, WM_SETTEXT, text[0])
                press(hwnd, self.pid, kids, 1 if action == 'ok' else 2, VK_RETURN if action == 'ok' else VK_ESCAPE)
            elif kind == '#32770':
                texts = [watcher._window_text(kid) for kid in kids if watcher._window_text(kid)]
                note('alert ' + ('file dialog' if 'ShellView' in texts else ' / '.join(texts)))  # its pane names the machine's folders
                if 'ShellView' in texts:  # a form of Excel's own: OK would save or open something
                    press(hwnd, self.pid, kids, 2, VK_ESCAPE)
                    continue
                press(hwnd, self.pid, kids, 1, VK_RETURN)
                if mine(hwnd, self.pid) and user32.IsWindowVisible(hwnd):
                    press(hwnd, self.pid, kids, 2, VK_ESCAPE)
            else:
                note(f'other {kind} {title}')
                if mine(hwnd, self.pid):
                    user32.PostMessageW(hwnd, WM_KEYDOWN, VK_ESCAPE, 0)
                    user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)

    watcher.DialogWatcher._scan = scan  # type: ignore[method-assign]
    sys.exit(worker.main())


#: Each in a new Excel, what one session's dialogs keep. `s` collects the answers, `d(k)` is FileDialog(k).
FRESH = {
    'defaults': 'For k = 1 To 4: Set x = d(k): s = s & k & ":" & x.Title & "|" & x.ButtonName & "|" & x.AllowMultiSelect & "|" & x.InitialView & "|" & x.InitialFileName & "|" & x.FilterIndex & "|" & x.Filters.Count & "|" & x.SelectedItems.Count & ";": Next',
    'filters open': 'Set x = d(1): For k = 1 To x.Filters.Count: s = s & x.Filters(k).Description & "=" & x.Filters(k).Extensions & ";": Next',
    'filters save': 'Set x = d(2): For k = 1 To x.Filters.Count: s = s & x.Filters(k).Description & "=" & x.Filters(k).Extensions & ";": Next',
    'filters pickers': 'Set x = d(3): For k = 1 To x.Filters.Count: s = s & x.Filters(k).Description & "=" & x.Filters(k).Extensions & ";": Next: s = s & "/" & d(4).Filters.Count',
    'title not shared': 'd(1).Title = "t": For k = 1 To 4: s = s & d(k).Title & ";": Next: d(3).Title = "": s = s & d(3).Title & ";": d(4).Title = "": s = s & d(4).Title & ";": d(1).Title = "": s = s & d(1).Title & ";": d(2).Title = "": s = s & d(2).Title',
    'button not shared': 'd(1).ButtonName = "b": For k = 1 To 4: s = s & d(k).ButtonName & ";": Next: d(1).ButtonName = "": s = s & d(1).ButtonName & ";"',
    'view shared': 'd(1).InitialView = 3: For k = 1 To 4: s = s & d(k).InitialView & ";": Next: d(1).InitialView = 99: s = s & d(1).InitialView',
    'index shared': 'd(1).FilterIndex = 3: For k = 1 To 4: s = s & d(k).FilterIndex & ";": Next: d(1).FilterIndex = 99: s = s & d(1).FilterIndex & ";": d(1).FilterIndex = 0: s = s & d(1).FilterIndex & ";": d(1).FilterIndex = -1: s = s & d(1).FilterIndex',
    'name shared': 'd(1).InitialFileName = "@DIR@\\b.txt": For k = 1 To 4: s = s & d(k).InitialFileName & ";": Next: d(1).InitialFileName = "@DIR@\\": s = s & d(1).InitialFileName & ";": d(1).InitialFileName = "@DIR@": s = s & d(1).InitialFileName & ";": d(1).InitialFileName = "@DIR@\\*.xlsx": s = s & d(1).InitialFileName & ";": d(1).InitialFileName = "": s = s & "[" & d(1).InitialFileName & "];": d(1).InitialFileName = "@DIR@\\no such\\b.txt": s = s & d(1).InitialFileName & ";": d(1).InitialFileName = "http://x.test/a": s = s & d(1).InitialFileName',
    'multi not shared': 'd(1).AllowMultiSelect = False: For k = 1 To 4: s = s & d(k).AllowMultiSelect & ";": Next: d(3).AllowMultiSelect = False: s = s & d(3).AllowMultiSelect',
    'filters open and picker': 'd(1).Filters.Clear: For k = 1 To 4: s = s & d(k).Filters.Count & ";": Next',
    'filters picker and open': 'd(3).Filters.Add "A", "*.a": For k = 1 To 4: s = s & d(k).Filters.Count & ";": Next',
    'filter edits': 'Set x = d(1).Filters: x.Clear: x.Add "A", "*.a": s = s & x.Count & ";": x.Add "B", "*.b;*.bb": s = s & x.Count & ";": x.Add "C", "*.c", 1: s = s & x(1).Description & x(2).Description & x(3).Description & x(3).Extensions & ";": s = s & d(1).FilterIndex & ";": x.Delete 2: s = s & x.Count & x(2).Description & ";": x.Delete: s = s & x.Count & ";"',
    'filter empty extension': 'd(1).Filters.Add "A", ""',
    'filter empty name': 's = d(1).Filters.Add("", "*.a").Description',
    'filter far position': 'Set x = d(1): s = x.Filters.Add("A", "*.a", 99).Description',
    'filter position zero': 'Set x = d(1): s = x.Filters.Add("A", "*.a", 0).Description',
    'filter delete far': 'Set x = d(1): x.Filters.Delete 99',
    'filter plain name': 'Set x = d(1): s = x.Filters.Add("A", "a.b").Extensions',
    'filter star': 'Set x = d(1): s = x.Filters.Add("A", "*").Extensions',
    'filter comma': 'Set x = d(1): s = x.Filters.Add("A", "*.a,*.b").Extensions',
    'filter parents': 'Set x = d(1): s = TypeName(x.Filters(1)) & TypeName(x.Filters.Item(1)) & TypeName(x.Filters(1).Parent) & TypeName(x.Filters.Parent) & TypeName(x.SelectedItems.Parent)',
    'filter lifetime': 'd(3).Filters.Add "A", "*.a": s = s & d(3).Filters.Count & ";": Set x = d(3): x.Filters.Add "A", "*.a": s = s & x.Filters.Count & ";" & d(3).Filters.Count & ";": Set y = x.Filters: y.Add "B", "*.b": s = s & y.Count & x.Filters.Count & d(3).Filters.Count & ";"',
    'filter lifetime open': 'd(1).Filters.Clear: d(1).Filters.Add "A", "*.a": s = s & d(1).Filters.Count & ";": Set x = d(1): s = s & x.Filters.Count & ";": x.Filters.Add "B", "*.b": s = s & x.Filters.Count & ";"',
    'title after show': 'Set x = d(1): x.Title = "drive|cancel": x.ButtonName = "Go": x.AllowMultiSelect = False: x.Show: s = x.Title & ";" & x.ButtonName & ";" & x.AllowMultiSelect & ";" & x.FilterIndex & ";" & x.InitialView & ";" & x.Filters.Count',
    'filter index empty': 'Set x = d(1): x.Filters.Clear: x.FilterIndex = 2: s = x.FilterIndex & ";": x.Filters.Add "A", "*.a": x.FilterIndex = 1: s = s & x.FilterIndex',
    'selection kept': 'Set x = d(1): x.Title = "drive|ok|@DIR@\\pick.txt": s = s & x.Show & x.SelectedItems.Count & ";": x.Title = "drive|cancel": s = s & x.Show & x.SelectedItems.Count & ";": s = s & d(3).SelectedItems.Count & ";" & x.InitialFileName & ";" & x.Title',
    'selection walk': 'Set x = d(1): x.Title = "drive|ok|@DIR@\\pick.txt": x.Show: For Each y In x.SelectedItems: s = s & TypeName(y) & ";": Next: s = s & x.SelectedItems.Item(1) & ";": On Error Resume Next: s = s & x.SelectedItems(0) & "e" & Err.Number & ";": Err.Clear: s = s & x.SelectedItems(2) & "e" & Err.Number',
    'folder after pick': 'ChDir "@DIR@\\sub": s = s & CurDir & ";": r = Application.GetOpenFilename(Title:="drive|ok|@DIR@\\pick.txt"): s = s & CurDir & ";": ChDir "@DIR@\\sub": r = Application.GetSaveAsFilename(Title:="drive|ok|@DIR@\\new.txt"): s = s & CurDir & ";": ChDir "@DIR@\\sub": Set x = d(1): x.Title = "drive|ok|@DIR@\\pick.txt": x.Show: s = s & CurDir & ";": ChDir "@DIR@\\sub": r = Application.GetOpenFilename(Title:="drive|cancel"): s = s & CurDir & ";": d(1).InitialFileName = "abc": s = s & d(1).InitialFileName',
    'not opened': 'n = Workbooks.Count: r = Application.GetOpenFilename(Title:="drive|ok|@DIR@\\book.xlsx"): s = s & r & Workbooks.Count - n & ";": Set x = d(1): x.Title = "drive|ok|@DIR@\\book.xlsx": x.Show: s = s & Workbooks.Count - n & ";": s = s & Dir("@DIR@\\new2.xlsx") & ";"',
    'one object': 'd(1).Title = "t": Set x = d(1): Set y = d(2): s = x.Title & ";" & x.DialogType & ";" & y.Title & ";" & (x Is y) & ";" & x.SelectedItems.Count & ";": s = s & d(1).Title',
    'back to a type': 'Set x = d(1): x.Filters.Clear: x.ButtonName = "b": x.AllowMultiSelect = False: d(2): Set x = d(1): s = x.Filters.Count & ";" & x.ButtonName & ";" & x.AllowMultiSelect',
    'selection and type': 'Set x = d(1): x.Title = "drive|ok|@DIR@\\pick.txt": x.Show: s = x.SelectedItems.Count & ";": Set y = d(1): s = s & y.SelectedItems.Count & ";": Set y = d(3): s = s & x.SelectedItems.Count',
    'execute open': 'n = Workbooks.Count: Set x = d(1): x.Title = "drive|ok|@DIR@\\book.xlsx": x.Show: r = x.Execute: s = s & TypeName(r) & Workbooks.Count - n & ActiveWorkbook.Name & ";" & x.SelectedItems.Count',
    'execute save names': 'Dim b As Workbook: Set b = Workbooks.Add: k = Workbooks.Count: Set x = d(2): x.Title = "drive|ok|@DIR@\\saved3.xlsx": x.Show: x.Execute: s = TypeName(ActiveWorkbook) & "/" & (ActiveWorkbook Is b) & "/" & b.Name & "/" & (ThisWorkbook.Name = ActiveWorkbook.Name) & "/" & (Workbooks.Count - k) & "/" & b.Saved & "/" & (Dir("@DIR@\\saved3.xlsx") <> "")',
    'dialog walk bare': 'For Each y In Application.Dialogs: s = s & TypeName(y) & ";": Exit For: Next',
    'execute save': 'Workbooks.Add: Set x = d(2): x.Title = "drive|ok|@DIR@\\saved.xlsx": x.Show: r = x.Execute: s = s & TypeName(r) & ActiveWorkbook.Name & ";" & Dir("@DIR@\\saved.xlsx") & ";": ActiveWorkbook.Close False',
    'execute pickers': 'For k = 3 To 4: Set x = d(k): x.Title = "drive|ok|@DIR@" & IIf(k = 3, "\\book.xlsx", ""): x.Show: n = Workbooks.Count: On Error Resume Next: x.Execute: s = s & Err.Number & "/" & Workbooks.Count - n & ";": Err.Clear: On Error GoTo 0: Next',
    'dialog numbers': 'On Error Resume Next: For k = -2 To 3000: Err.Clear: Set y = Nothing: Set y = Application.Dialogs(k): If Err.Number = 0 Then s = s & k & ","\nNext',
    'dialog walk': 'On Error Resume Next: For Each y In Application.Dialogs: m = m + 1: If m < 3 Then s = s & TypeName(y) & ";"\nNext: s = s & m & "e" & Err.Number',
}


def _cases(folder: str) -> list[str]:
    cases = [
        'r = InputBox("p", "drive|cancel", "d")', 'r = StrPtr(InputBox("p", "drive|cancel", "d")) = 0', 'r = InputBox("p", "drive|ok", "d")', 'r = InputBox("p", "drive|ok|typed", "d")',
        'r = StrPtr(InputBox("p", "drive|ok|", "")) = 0',
        'r = Application.InputBox("p", "drive|cancel")', 'r = Application.InputBox("p", "drive|ok", "d")', 'r = Application.InputBox("p", "drive|ok", 5)', 'r = Application.InputBox("p", "drive|ok")', 'r = Application.InputBox("p", "drive|ok|typed")', 'r = Application.InputBox("p", "drive|ok|12")', 'r = Application.InputBox("p", "drive|ok|=1+2")', 'r = Application.InputBox("p", "drive|ok|TRUE")',
    ]
    for kind in (0, 1, 2, 4, 8, 16, 64, 3, 5, 128):
        cases.append(f'r = Application.InputBox("p", "drive|cancel", Type:={kind})')
        for answer in ('12', '1.5', 'abc', 'TRUE', '=1+2', '=A1', 'A1', '#N/A', '{{1,2}}', ''):
            cases.append(f'r = Application.InputBox("p", "drive|ok|{answer}", Type:={kind})'.replace('{{', '{').replace('}}', '}'))
    for kind in (1, 2):
        cases += [f'r = Application.InputBox("p", "drive|ok|{answer}", Type:={kind})' for answer in ('50%', '1,000', '$5', '1/2/2020', '1e3', ' 12 ', '=1/0', '=NA()', '=A1:A2', 'FALSE', 'false')]
    cases += [f'r = Application.InputBox("p", "drive|ok|{answer}", Type:=4)' for answer in ('FALSE', 'false', '0', '=1=2')]
    cases += [f'r = Application.InputBox("p", "drive|ok|{answer}", Type:=0)' for answer in ('=B2', '=$A$1', '=Sheet1!A1', '=SUM(A1:A2)')]
    cases += [f'r = Application.InputBox("p", "drive|ok|{answer}", Type:=64): r = TypeName(r) & "/" & LBound(r) & "-" & UBound(r) & "/" & r(LBound(r))' for answer in ('12', 'TRUE', '=1+2', '=A1', 'A1')]
    cases += [f'r = Application.InputBox("p", "drive|ok|{answer}", Type:=64): r = TypeName(r) & "/" & LBound(r, 1) & "-" & UBound(r, 1) & "/" & LBound(r, 2) & "-" & UBound(r, 2) & "/" & r(1, 1)' for answer in ('12', 'A1', 'A1:B2', '{1,2}', '{1,2;3,4}')]
    cases += [
        'r = Application.InputBox("p", "drive|ok", 5, Type:=1)', 'r = Application.InputBox("p", "drive|ok", ActiveSheet.Range("A1"), Type:=8).Address', 'r = Application.InputBox("p", "drive|ok", ActiveSheet.Range("A1"))', 'r = Application.InputBox("p", "drive|ok", "=A1", Type:=2)', 'r = Application.InputBox("p", "drive|ok", True, Type:=4)',
        'Set x = Application.InputBox("p", "drive|ok|Sheet9!A1", Type:=8): r = x.Address', 'Set x = Application.InputBox("p", "drive|ok|A1,C3", Type:=8): r = x.Address', 'ActiveWorkbook.Names.Add "Nm", "=Sheet1!$A$1": Set x = Application.InputBox("p", "drive|ok|Nm", Type:=8): r = x.Address',
        'Set x = Application.InputBox("p", "drive|cancel", Type:=8)', 'Set x = Application.InputBox("p", "drive|ok|A1:B2", Type:=8): r = x.Address', 'Set x = Application.InputBox("p", "drive|ok|Sheet1!C3", Type:=8): r = x.Address', 'r = Application.InputBox("p", "drive|cancel", Type:="x")', 'r = Application.InputBox("p", "drive|cancel", Type:=-1)', 'r = Application.InputBox("p", "drive|cancel", Left:=1, Top:=1)', 'r = Application.InputBox(1, "drive|cancel")', 'r = Application.InputBox("", "drive|cancel")',
        'r = Application.GetOpenFilename(Title:="drive|cancel")', 'r = Application.GetOpenFilename(Title:="drive|cancel", MultiSelect:=True)', 'r = TypeName(Application.GetOpenFilename(Title:="drive|cancel", MultiSelect:=True))',
        f'r = Application.GetOpenFilename(Title:="drive|ok|{folder}\\pick.txt")', f'r = Application.GetOpenFilename(Title:="drive|ok|{folder}\\pick.txt", MultiSelect:=True)(1)', f'r = Application.GetOpenFilename(Title:="drive|ok|{folder}\\pick.txt", MultiSelect:=True): r = LBound(r) & "-" & UBound(r) & TypeName(r)',
        'r = Application.GetOpenFilename("abc", Title:="drive|cancel")', 'r = Application.GetOpenFilename("Text,*.txt", 5, "drive|cancel")', 'r = Application.GetOpenFilename("Text,*.txt,All,*.*", 0, "drive|cancel")', 'r = Application.GetOpenFilename("", Title:="drive|cancel")', 'r = Application.GetOpenFilename("a,b,c", Title:="drive|cancel")', 'r = Application.GetOpenFilename(1, Title:="drive|cancel")',
        'r = Application.GetSaveAsFilename(Title:="drive|cancel")', f'r = Application.GetSaveAsFilename(Title:="drive|ok|{folder}\\new.txt")', f'r = Application.GetSaveAsFilename(Title:="drive|ok|{folder}\\new", FileFilter:="Text,*.txt")', f'r = Application.GetSaveAsFilename("{folder}\\given.csv", "CSV,*.csv", , "drive|ok")', 'r = Application.GetSaveAsFilename(FileFilter:="abc", Title:="drive|cancel")',
    ]
    for kind in (1, 2, 3, 4):
        with_box = f'Set x = Application.FileDialog({kind}): '
        for member in ('TypeName(x)', 'x.DialogType', 'x.Title', 'x.ButtonName', 'x.AllowMultiSelect', 'x.InitialView', 'x.InitialFileName', 'x.FilterIndex', 'x.Filters.Count', 'x.Filters(1).Description', 'x.Filters(1).Extensions', 'x.SelectedItems.Count', 'x.Item', 'TypeName(x.Parent)', 'x.Creator', 'TypeName(x.Application)'):
            cases.append(with_box + f'r = {member}')
        cases += [with_box + 'x.Title = "drive|cancel": r = x.Show', with_box + 'x.Title = "drive|cancel": x.Show: r = x.SelectedItems.Count', with_box + 'x.Title = "t": r = x.Title', with_box + 'x.Title = "": r = x.Title', with_box + 'x.ButtonName = "Go": r = x.ButtonName', with_box + 'x.AllowMultiSelect = True: r = x.AllowMultiSelect', with_box + 'x.InitialFileName = "abc": r = x.InitialFileName', with_box + 'x.InitialView = 2: r = x.InitialView', with_box + 'x.FilterIndex = 2: r = x.FilterIndex', with_box + 'x.Filters.Clear: r = x.Filters.Count', with_box + 'r = x.Filters.Add("D", "*.d").Description', with_box + 'r = x.Filters.Add("D", "*.d", 1).Extensions & x.Filters(1).Description',
                  with_box + 'x.Filters.Delete 1: r = x.Filters.Count', with_box + 'r = x.Filters.Add("D", "bad").Extensions', with_box + 'r = x.Filters(99).Description', with_box + 'r = x.SelectedItems(1)', with_box + 'x.Execute']
        name = 'pick.txt' if kind in (1, 3) else 'new.txt' if kind == 2 else ''
        target = f'{folder}\\{name}' if name else folder
        cases += [with_box + f'x.Title = "drive|ok|{target}": r = x.Show & "/" & x.SelectedItems.Count & "/" & x.SelectedItems(1) & "/" & x.Item', with_box + f'x.Title = "drive|ok|{target}": x.Show: Set y = x.SelectedItems: r = TypeName(y) & y.Count & y.Item(1)']
    cases += [
        'Set x = Application.FileDialog(1): x.Title = "t": Set y = Application.FileDialog(1): r = y.Title', 'r = Application.FileDialog(1) Is Application.FileDialog(1)', 'r = TypeName(Application.FileDialog(5))', 'r = TypeName(Application.FileDialog(0))', 'r = TypeName(Application.FileDialog("x"))',
        'r = TypeName(Application.Dialogs)', 'r = Application.Dialogs.Count', 'r = TypeName(Application.Dialogs(xlDialogOpen))', 'r = TypeName(Application.Dialogs(1))', 'r = Application.Dialogs(1).Show', 'r = TypeName(Application.Dialogs(99999))', 'r = TypeName(Application.Dialogs("x"))', 'r = TypeName(Application.Dialogs.Item(xlDialogOpen))', 'r = TypeName(Application.Dialogs(xlDialogOpen).Application)', 'r = TypeName(Application.Dialogs(xlDialogOpen).Parent)', 'r = Application.Dialogs(xlDialogOpen).Creator', 'r = Application.Dialogs(xlDialogOpen).Show', 'r = Application.Dialogs(xlDialogSaveAs).Show', 'r = Application.Dialogs(xlDialogFontProperties).Show',
    ]
    cancelled = 'Dim s As String: s = InputBox("p", "drive|cancel"): '
    cases += [
        'Dim s As String: r = StrPtr(s) = 0', 'Dim s As String: s = "": r = StrPtr(s) = 0', 'Dim s As String: s = vbNullString: r = StrPtr(s) = 0', 'Dim s As String: s = "a": s = "": r = StrPtr(s) = 0',
        'r = StrPtr(vbNullString) = 0', 'r = StrPtr("") = 0', 'r = StrPtr("a") = 0', 'r = TypeName(StrPtr("a"))', 'r = StrPtr(InputBox("p", "drive|ok|typed")) = 0',
        cancelled + 'r = StrPtr(s) = 0', cancelled + 'Dim t As String: t = s: r = StrPtr(t) = 0', cancelled + 'r = StrPtr(s & "") = 0', cancelled + 'r = StrPtr(Trim$(s)) = 0', cancelled + 'r = StrPtr(Left$(s, 0)) = 0', cancelled + 'r = StrPtr(CStr(s)) = 0', cancelled + 'r = StrPtr(UCase$(s)) = 0',
        cancelled + 'r = (s = "") & Len(s) & (s = vbNullString) & LenB(s)', cancelled + 'Dim v As Variant: v = s: r = StrPtr(v) = 0', cancelled + 's = s: r = StrPtr(s) = 0', cancelled + 'Dim a(1) As String: a(1) = s: r = StrPtr(a(1)) = 0',
        cancelled + 'Dim c As New Collection: c.Add s: r = StrPtr(c(1)) = 0',
        'Dim v As Variant: v = InputBox("p", "drive|cancel"): r = StrPtr(v) = 0', 'Dim v As Variant: v = InputBox("p", "drive|cancel"): r = TypeName(v) & (v = "")', 'Dim v As Variant: v = vbNullString: r = StrPtr(v) = 0', 'Dim v As Variant: v = "": r = StrPtr(v) = 0',
        'Dim v As Variant: r = StrPtr(v) = 0', 'r = StrPtr(Null) = 0', 'r = StrPtr(12) = 0', 'r = StrPtr(True) = 0', 'r = StrPtr(Application.InputBox("p", "drive|cancel")) = 0', 'Dim a(1) As String: r = StrPtr(a(1)) = 0', 'Dim s As String: r = StrPtr(s & "") = 0', 'Dim s As String: r = StrPtr(Trim$(s)) = 0',
        'r = StrPtr(ActiveSheet.Range("B9").Value) = 0', 'r = StrPtr(Left$("abc", 0)) = 0', 'Dim s As String: s = "a": r = (StrPtr(s) = StrPtr(s))', 'Dim s As String * 2: r = StrPtr(s) = 0', 'r = StrPtr() = 0',
    ]
    return cases


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        folder = temp
        (Path(temp) / 'pick.txt').write_text('x', encoding='ascii')
        log = Path(temp) / 'driver.log'
        log.touch()
        cases = _cases(folder)
        selected = [int(arg) for arg in sys.argv[1:]] or list(range(len(cases)))
        config = HarnessConfig(exclusive=False, worker_argv=[sys.executable, str(Path(__file__).resolve()), '--worker'], extra_env={'DIALOG_LOG': str(log)})

        def module(statement: str) -> str:
            return ('Option Explicit\nPrivate Function Shown(ByVal v As Variant) As String\n    On Error Resume Next\n    If IsError(v) Then\n        Shown = "Error:" & CStr(v)\n    ElseIf IsObject(v) Then\n        Shown = "object " & TypeName(v)\n    ElseIf IsArray(v) Then\n        Shown = "array " & TypeName(v)\n    Else\n        Shown = TypeName(v) & ":" & CStr(v)\n    End If\nEnd Function\n'
                    'Public Function Tried() As String\n    Dim r As Variant, x As Object, y As Object, wb As Workbook\n    Application.DisplayAlerts = True\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n    wb.Worksheets(1).Range("A1").Value = 7\n    On Error GoTo Failed\n'
                    f'    {statement}\n    Tried = "ok " & Shown(r)\nAnswered:\n    On Error Resume Next\n    Application.DisplayAlerts = False\n    wb.Close False\n    Exit Function\nFailed:\n    Tried = "Err " & Err.Number & " " & Err.Description\n    Resume Answered\nEnd Function\n')

        done: dict[str, list[object]] = {}
        if not sys.argv[1:] and OUT.exists():
            done = {case: rest for case, *rest in json.loads(OUT.read_text(encoding='utf-8'))['cases']}
        answers: dict[int, list[object]] = {index: done[shown] for index in selected if (shown := cases[index].replace(folder, '<DIR>')) in done}
        pending = [index for index in selected if index not in answers]
        ends = time.monotonic() + 500.0
        while pending and time.monotonic() < ends:
            with ExcelSession(config) as excel:
                excel.new_document()
                while pending and time.monotonic() < ends:
                    index = pending.pop(0)
                    start = log.stat().st_size
                    result = excel.run_vba(module(cases[index]), 'Tried', (), timeout=20.0)
                    with log.open(encoding='utf-8') as file:
                        file.seek(start)
                        driven = file.read().splitlines()
                    answer = str(result.value) if result.ok else f'breaks: {result.outcome}'
                    alerts = [line.removeprefix('alert ') for line in driven if line.startswith(('alert ', 'other '))]
                    answers[index] = [answer.replace(folder, '<DIR>'), [line.replace(folder, '<DIR>') for line in alerts]]
                    print(index, cases[index].replace(folder, '<DIR>'), '=>', answers[index], [line for line in driven if line.startswith('form ')], flush=True)
                    if not result.ok:
                        break
        if not sys.argv[1:]:
            fresh: dict[str, str] = {name: pair[1] for name, pair in json.loads(OUT.read_text(encoding='utf-8')).get('fresh', {}).items() if name in FRESH} if OUT.exists() else {}
            (Path(temp) / 'sub').mkdir()
            prepared = False
            for name, statement in FRESH.items():
                if name in fresh or time.monotonic() >= ends:
                    continue
                with ExcelSession(config) as excel:
                    excel.new_document()
                    if not prepared:
                        excel.run_vba(f'Public Sub Made()\n    With Workbooks.Add(xlWBATWorksheet)\n        .SaveAs "{folder}\\book.xlsx", xlOpenXMLWorkbook\n        .Close False\n    End With\nEnd Sub\n', 'Made', (), timeout=30.0)
                        prepared = True
                    source = ('Option Explicit\nPrivate Function d(ByVal k As Long) As Object\n    Set d = Application.FileDialog(k)\nEnd Function\nPublic Function Tried() As String\n    Dim s As String, k As Long, n As Long, m As Long, x As Object, y As Variant, r As Variant\n    On Error GoTo Failed\n'
                              + '    ' + statement.replace('@DIR@', folder).replace('\n', '\n    ')
                              + '\n    Tried = s\n    Exit Function\nFailed:\n    Tried = s & "Err " & Err.Number & " " & Err.Description\nEnd Function\n')
                    result = excel.run_vba(source, 'Tried', (), timeout=40.0)
                    fresh[name] = (str(result.value) if result.ok else f'breaks: {result.outcome}').replace(folder, '<DIR>').replace(str(Path.home()), '<HOME>')
                    print(name, '=>', fresh[name], flush=True)
            record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'fresh': {name: [statement, fresh[name]] for name, statement in FRESH.items() if name in fresh},
                      'cases': [[cases[index].replace(folder, '<DIR>'), *answers[index]] for index in selected if index in answers]}
            OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
            print(len(answers), 'of', len(cases), flush=True)


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--worker':
        del sys.argv[1]
        _drive()
    else:
        main()
