"""A defined name's macro metadata: MacroType, ShortcutKey and Category, as Excel keeps and saves them
(tests/fixtures/name_macro.json, scripts/measure_name_macro.py).

A name is an ordinary one until MacroType makes it an Excel 4 command (2) or function (1); then it takes a
shortcut key, which commands save, and a function category, which the file keeps as a functionGroupId. Excel
leaves such names out of a workbook saved without macros.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import NameEntry

#: Excel's function categories in functionGroupId order, 1 first.
CATEGORIES = ('Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text',
              'Logical', 'Information', 'Commands', 'Customizing', 'Macro Control', 'DDE/External', 'User Defined',
              'Engineering', 'Cube', 'Compatibility', 'Web')
#: The group a category Excel does not have is saved under: User Defined's.
_OTHER = CATEGORIES.index('User Defined') + 1
#: What MacroType reads for an ordinary name: xlNone, though xlNotXLM cannot be set.
NOT_MACRO = -4142


def macro_type(entry: NameEntry) -> int:
    return entry.macro_type or NOT_MACRO


def set_macro_type(entry: NameEntry, value: object) -> None:
    """Only xlFunction and xlCommand are taken; xlNotXLM and anything else are refused with 1004."""
    wanted = int(to_integer(value))
    if wanted not in (1, 2):
        raise error(1004)
    if not entry.macro_type:
        entry.category, entry.group = 'User Defined', _OTHER
    entry.macro_type = wanted
    entry.macro_changed = True


def _macro(entry: NameEntry) -> None:
    if not entry.macro_type:
        raise error(1004)


def shortcut_key(entry: NameEntry) -> str:
    _macro(entry)
    return entry.shortcut_key


def set_shortcut_key(entry: NameEntry, value: object) -> None:
    """The first character, a letter; "" takes the key away."""
    _macro(entry)
    text = to_text(value)[:1]
    if text and not ('a' <= text.lower() <= 'z'):
        if text.isalpha():
            raise VBAUnsupportedError('A shortcut key outside A to Z is not implemented')
        raise error(1004)
    entry.shortcut_key = text
    entry.macro_changed = True


def category(entry: NameEntry) -> str:
    _macro(entry)
    return entry.category


def set_category(entry: NameEntry, value: object) -> None:
    """One of Excel's categories in its own spelling, matched ignoring case; any other text as given."""
    _macro(entry)
    text = to_text(value)
    at = next((at for at, one in enumerate(CATEGORIES, start=1) if one.lower() == text.lower()), None)
    entry.category, entry.group = (CATEGORIES[at - 1], at) if at is not None else (text, _OTHER)
    entry.macro_changed = True


def added(entry: NameEntry, kind: object, key: object, named: object, local: object) -> None:
    """Names.Add's MacroType, ShortcutKey, Category and CategoryLocal: a name it redefines starts ordinary again,
    and without a MacroType of xlFunction or xlCommand the rest are ignored."""
    from pyopenvba.interpreter._values import MISSING

    entry.macro_type, entry.shortcut_key, entry.category, entry.group = 0, '', '', 0
    entry.macro_changed = True
    if kind is MISSING or int(to_integer(kind)) not in (1, 2):
        return
    set_macro_type(entry, kind)
    if key is not MISSING:
        set_shortcut_key(entry, key)
    for given in (named, local):
        if given is not MISSING:
            set_category(entry, given)


def loaded(entry: NameEntry, attributes: dict[str, str]) -> None:
    """A name read from a file: Excel's Category reads All for it, whatever group it was saved in."""
    if attributes.get('xlm', '0') not in ('1', 'true'):
        return
    entry.macro_type = 1 if attributes.get('function', '0') in ('1', 'true') else 2
    entry.shortcut_key = attributes.get('shortcutKey', '')
    entry.category = 'All'
    group = attributes.get('functionGroupId', '')
    entry.group = int(group) if group.isdecimal() else 0


def saved(entry: NameEntry) -> dict[str, str | None]:
    """The attributes a changed name is saved with, None for one it has not."""
    if not entry.macro_type:
        return {'function': None, 'xlm': None, 'functionGroupId': None, 'shortcutKey': None}
    if not entry.category:
        raise VBAUnsupportedError('Saving a function name with an empty Category is not implemented')
    command = entry.macro_type == 2
    return {'function': None if command else '1', 'xlm': '1', 'functionGroupId': str(entry.group),
            'shortcutKey': entry.shortcut_key or None if command else None}
