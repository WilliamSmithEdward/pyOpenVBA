"""Which button a message box comes back with when nobody clicks and Enter is pressed (issue #250): the answer
Windows' MessageBox gives for each MsgBox style, VBA's vbDefaultButton2, 3 and 4 among them.

VBA's MsgBox hands its buttons and vbDefaultButtonN, which are MB_DEFBUTTON1 to 4, to MessageBoxW. Each box is opened
on a daemon thread of its own, so nothing outlives the script; this thread finds it by its caption, records the
default button DM_GETDEFID reports, and presses Enter on the control holding the focus, which is what the box
answers. A box still open three seconds on is closed by its buttons' commands and recorded as a failure.

Predictions written before measuring: the answer is the button vbDefaultButtonN names, counted from the first; a
number past the box's last button gives the first.

    python scripts/measure_msgbox_default.py

writes tests/fixtures/msgbox_default.json, which tests/test_vba_msgbox_default.py replays. Windows only; it shows
each box on the screen for a moment.
"""
from __future__ import annotations

import ctypes
import json
import threading
import time
from ctypes import wintypes
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/msgbox_default.json'
#: The button sets, vbOKOnly to vbRetryCancel, and the defaults, none to vbDefaultButton4.
STYLES = range(6)
DEFAULTS = (0, 256, 512, 768)
DM_GETDEFID, WM_COMMAND, WM_KEYDOWN, WM_KEYUP, VK_RETURN, DC_HASDEFID = 0x0400, 0x0111, 0x0100, 0x0101, 0x0D, 0x534B

user32 = ctypes.WinDLL('user32', use_last_error=True)
user32.MessageBoxW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
user32.FindWindowW.restype = wintypes.HWND
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [('cbSize', wintypes.DWORD), ('flags', wintypes.DWORD), ('hwndActive', wintypes.HWND),
                ('hwndFocus', wintypes.HWND), ('hwndCapture', wintypes.HWND), ('hwndMenuOwner', wintypes.HWND),
                ('hwndMoveSize', wintypes.HWND), ('hwndCaret', wintypes.HWND), ('rcCaret', wintypes.RECT)]


user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]


def answer_of(style: int) -> tuple[int, int | None]:
    """What the box opened with ``style`` answers to Enter, and the default button DM_GETDEFID reported."""
    caption = f'pyOpenVBA probe {style}'
    answer: list[int] = []
    shown = threading.Thread(target=lambda: answer.append(user32.MessageBoxW(None, 'probe', caption, style)),
                             daemon=True)
    shown.start()
    window = None
    for _ in range(250):
        window = user32.FindWindowW('#32770', caption)
        if window:
            break
        time.sleep(0.02)
    if not window:
        raise RuntimeError(f'no box appeared for style {style}')
    time.sleep(0.15)
    found = user32.SendMessageW(window, DM_GETDEFID, 0, 0)
    reported = found & 0xFFFF if (found >> 16) & 0xFFFF == DC_HASDEFID else None
    info = GUITHREADINFO(cbSize=ctypes.sizeof(GUITHREADINFO))
    thread = user32.GetWindowThreadProcessId(window, None)
    focus = info.hwndFocus if user32.GetGUIThreadInfo(thread, ctypes.byref(info)) and info.hwndFocus else window
    user32.PostMessageW(focus, WM_KEYDOWN, VK_RETURN, 0)
    user32.PostMessageW(focus, WM_KEYUP, VK_RETURN, 0)
    shown.join(3)
    if shown.is_alive():
        for button in range(1, 8):
            user32.SendMessageW(window, WM_COMMAND, button, 0)
            shown.join(0.3)
            if not shown.is_alive():
                break
        raise RuntimeError(f'the box for style {style} did not answer Enter')
    return answer[0], reported


def main() -> None:
    answers: dict[str, int] = {}
    reported: dict[str, int | None] = {}
    for buttons in STYLES:
        for default in DEFAULTS:
            style = buttons + default
            answers[str(style)], reported[str(style)] = answer_of(style)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers, 'reported': reported}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for style, answer in answers.items():
        print(style, '->', answer, 'DM_GETDEFID', reported[style])


if __name__ == '__main__':
    main()
