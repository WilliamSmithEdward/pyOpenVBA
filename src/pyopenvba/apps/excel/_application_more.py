"""Members of Application a macro meets with no screen and no one at it, measured in live Excel
(scripts/measure_application_undo.py, tests/fixtures/application_undo.json).

Nothing is left to undo or repeat when a macro runs, so Undo and Repeat raise 1004. OnUndo and OnRepeat name what the
menu would offer, and OnKey a key's macro, and with no menu or keyboard here they are checked and held nowhere."""
from __future__ import annotations

import re

from pyopenvba.apps.excel._creator import Created
from pyopenvba.interpreter._objects import member
from pyopenvba.interpreter._values import MISSING, VBAInt, error, to_text

_KEY_NAMES = frozenset({
    'BACKSPACE', 'BS', 'BREAK', 'CAPSLOCK', 'CLEAR', 'DELETE', 'DEL', 'DOWN', 'END', 'ENTER', 'ESCAPE', 'ESC', 'HELP',
    'HOME', 'INSERT', 'INS', 'LEFT', 'NUMLOCK', 'PGDN', 'PGUP', 'RETURN', 'RIGHT', 'SCROLLLOCK', 'TAB', 'UP', '~',
    *(f'F{number}' for number in range(1, 16))})


def required(*values: object) -> None:
    """A required argument left out: 449."""
    if any(value is MISSING for value in values):
        raise error(449, 'Argument not optional')


def checked_key(key: object) -> None:
    """A key as OnKey's Key spells it: shift, control and alt signs, then one character or a name in braces."""
    required(key)
    found = re.fullmatch(r'[+^%]*(?:\{(.+)\}|([^{}]))', to_text(key))
    if found is None or (found.group(1) is not None and found.group(1).upper() not in _KEY_NAMES):
        raise error(1004, 'OnKey method of Application class failed')


class EmptyCollection(Created):
    """A collection the model never holds anything in."""

    def __init__(self, type_name: str) -> None:
        self.vba_type_name = type_name

    @member
    def Count(self) -> object:
        return VBAInt(0, 'Long')
