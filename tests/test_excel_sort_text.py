"""Text sorted as Windows' word sort orders it, which Excel's sort does, beyond ASCII too.

tests/fixtures/sort_keys.json holds the sort keys Windows' LCMapStringEx
(en-US, word sort) made of 1,500 strings drawn from Latin and its
extensions, combining marks, Greek, Cyrillic, punctuation, symbols and
control characters, ignoring case and matching it
(scripts/bake_sort_weights.py, which bakes the weights _sort.text_key puts
together). tests/fixtures/sort_text.json is the order live Excel's
Range.Sort put 400 such strings in, the spellings of one word that differ
in accents, case, hyphens and apostrophes among them, ignoring case and
matching it (scripts/measure_sort_text.py). tests/fixtures/style_order.json
is what live Excel did with cell styles named so: the names it refused,
the styles it found by other spellings, the Styles collection's order and
the saved file's (scripts/measure_style_order.py).
"""
from __future__ import annotations

import ctypes
import html
import json
import random
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._sort import ordering_key, text_key
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._locale import EN_US, JA_JP, Locale

FIXTURES = Path(__file__).parent / "fixtures"
KEYS: list[list[str]] = json.loads((FIXTURES / "sort_keys.json").read_text(encoding="utf-8"))["cases"]
ORDER: dict[str, object] = json.loads((FIXTURES / "sort_text.json").read_text(encoding="utf-8"))
STYLES: dict[str, Any] = json.loads((FIXTURES / "style_order.json").read_text(encoding="utf-8"))

#: ``ctypes`` seen as untyped: its Windows-only names (``WinDLL``) are not
#: in the stubs off Windows, and the tests are type-checked everywhere.
_windows: Any = ctypes


def test_every_key_is_the_one_windows_made() -> None:
    missed = [(text, ignoring, text_key(text, False).hex()) for text, ignoring, _matching in KEYS
              if text_key(text, False).hex() != ignoring]
    missed += [(text, matching, text_key(text, True).hex()) for text, _ignoring, matching in KEYS
               if text_key(text, True).hex() != matching]
    assert not missed, missed[:3]


@pytest.mark.parametrize("matching", [False, True], ids=["ignoring case", "matching case"])
def test_excels_sort_orders_text_as_the_model_does(matching: bool) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(ORDER["code"]), name="Measure")
    ignoring, matched = (part for part in str(app.run("Measure")).split("|"))
    found = [int(place) for place in (matched if matching else ignoring).split(",") if place]
    assert found == ORDER["matching_case" if matching else "ignoring_case"]


def _decoded(coded: str) -> str:
    return "".join(chr(int(part, 16)) for part in coded.split(".") if part)


def test_styles_named_beyond_ascii_are_refused_found_listed_and_saved_as_in_excel(tmp_path: Path) -> None:
    """A ligature is the letters it is made of and ss a sharp s, an accent or a hyphen tells two names apart, and
    case does not: as Excel refused and found styles, and in the order it listed and saved them."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(STYLES["code"], name="Measure")
    saved = tmp_path / "style_order.xlsx"
    added, found, listed = str(app.run("Measure", str(saved))).split("|")
    assert [int(number) for number in added.split(",") if number] == STYLES["added"]
    lookups = [[int(number), _decoded(name)] for number, name in (one.split(":") for one in found.split(",") if one)]
    assert lookups == STYLES["lookups"]
    assert [_decoded(name) for name in listed.split(",") if name] == STYLES["listed"]
    with zipfile.ZipFile(saved) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    assert [html.unescape(name) for name in re.findall(r'<cellStyle name="([^"]*)"', styles)] == STYLES["saved"]


@pytest.mark.parametrize("text", ["\u0301a", "\u4e2d", "\u0385", "\x02\u0301a", "-\u0301b", "a-\u0301",
                                  "\u03cd\u0528\u0309", "\u1f84\u0359"],
                         ids=["leading mark", "CJK", "dialytika tonos", "mark after a control character",
                              "mark after a hyphen", "mark after a letter and a hyphen", "mark after an ignored character",
                              "marks past 255"])
def test_text_the_model_does_not_sort_reports_itself(text: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        ordering_key(text, False)


def test_text_compare_string_orders_inconsistently_reports_itself() -> None:
    """CompareString puts U+1FE0 after U+03CD U+0528 U+0309 but before U+03CD U+0309, which it calls equal to that,
    and live Excel's sort did the same (U+0528 is a character word sort ignores): no key orders that."""
    with pytest.raises(VBAUnsupportedError):
        ordering_key("\u03cd\u0528\u0309", False)
    assert ordering_key("\u1fe0", False) < ordering_key("\u03cd\u0309", False)


def _orderable(text: str) -> bool:
    try:
        ordering_key(text, False)
    except VBAUnsupportedError:
        return False
    return True


@pytest.mark.skipif(sys.platform != "win32", reason="CompareStringEx is Windows'")
def test_fresh_strings_sort_as_compare_string_orders_them() -> None:
    compare = _windows.WinDLL("kernel32").CompareStringEx
    compare.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_wchar_p,
                        ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
    pick = random.Random(20261030)
    alphabet = sorted({character for text, _ignoring, _matching in KEYS for character in text})
    missed: list[tuple[str, str, bool]] = []
    for _ in range(20000):
        first, second = ("".join(pick.choice(alphabet) for _ in range(pick.randint(1, 6))) for _ in range(2))
        if not (_orderable(first) and _orderable(second)):
            # Text the model reports itself unable to order (see the tests above).
            continue
        for matching in (False, True):
            mine = (ordering_key(first, matching) > ordering_key(second, matching)) - \
                (ordering_key(first, matching) < ordering_key(second, matching))
            theirs = compare("en-US", 0 if matching else 1, first, -1, second, -1, None, None, None) - 2
            if mine != theirs:
                missed.append((first, second, matching))
    assert not missed, missed[:3]


# --- kana and ideographs ------------------------------------------------------------------------------------------

#: tests/fixtures/sort_keys_cjk.json: 1,500 strings of kana, ideographs, CJK symbols, full-width forms and ASCII
#: with the keys Windows made of them under en-US and ja-JP, ignoring case and matching it
#: (scripts/bake_sort_weights.py).
CJK_KEYS: list[list[str]] = json.loads((FIXTURES / "sort_keys_cjk.json").read_text(encoding="utf-8"))["cases"]
LOCALES = {"en-US": EN_US, "ja-JP": JA_JP}


def _reported(text: str, matching: bool, locale: Locale) -> bytes | None:
    try:
        return ordering_key(text, matching, locale)
    except VBAUnsupportedError:
        return None


@pytest.mark.parametrize("locale", sorted(LOCALES))
def test_kana_and_ideographs_weigh_as_windows_weighs_them(locale: str) -> None:
    """Every key the model makes is the one Windows made; the few it does not make are text it reports: a prolonged
    sound mark after something not a kana, an iteration mark, a combining mark with no kana before it."""
    column = 1 if locale == "en-US" else 3
    missed: list[tuple[str, bool]] = []
    reported = 0
    for case in CJK_KEYS:
        for offset, matching in ((0, False), (1, True)):
            mine = _reported(case[0], matching, LOCALES[locale])
            if mine is None:
                reported += 1
            elif mine.hex() != case[column + offset]:
                missed.append((case[0], matching))
    assert not missed, missed[:3]
    assert reported < len(CJK_KEYS) // 5


def test_ideographs_weigh_by_the_locale_and_names_weigh_none() -> None:
    """ja-JP weighs the common ideographs in Japanese order, en-US in Unicode's; names, compared without a locale,
    still report text beyond Latin, Greek and Cyrillic."""
    assert ordering_key("亜", False, JA_JP) < ordering_key("一", False, JA_JP)
    assert ordering_key("一", False, EN_US) < ordering_key("亜", False, EN_US)
    with pytest.raises(VBAUnsupportedError):
        text_key("あ", False)


@pytest.mark.skipif(sys.platform != "win32", reason="CompareStringEx is Windows'")
@pytest.mark.parametrize("locale", sorted(LOCALES))
def test_kana_and_ideographs_sort_as_compare_string_orders_them(locale: str) -> None:
    compare = _windows.WinDLL("kernel32").CompareStringEx
    compare.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_wchar_p,
                        ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
    pick = random.Random(20261031)
    texts = [case[0] for case in CJK_KEYS]
    missed: list[tuple[str, str, bool]] = []
    for _ in range(20000):
        first, second = pick.choice(texts), pick.choice(texts)
        for matching in (False, True):
            mine_first = _reported(first, matching, LOCALES[locale])
            mine_second = _reported(second, matching, LOCALES[locale])
            if mine_first is None or mine_second is None:
                continue
            mine = (mine_first > mine_second) - (mine_first < mine_second)
            theirs = compare(locale, 0 if matching else 1, first, -1, second, -1, None, None, None) - 2
            if mine != theirs:
                missed.append((first, second, matching))
    assert not missed, missed[:3]


LIVE_SORTS: dict[str, Any] = json.loads((FIXTURES / "vba_locale_live.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("locale", sorted(LIVE_SORTS))
@pytest.mark.parametrize("matching", [False, True], ids=["ignoring case", "matching case"])
def test_range_sort_orders_kana_and_kanji_as_live_excel_did(locale: str, matching: bool) -> None:
    """Range.Sort of 25 kana, kanji, full-width and ASCII strings, as live Excel put them under each Regional format
    (tests/fixtures/vba_locale_live.json): the kanji in Japanese order under ja-JP, in Unicode's under en-US."""
    kind = "sort-case" if matching else "sort"
    texts, order = next((payload, answer) for _label, case_kind, payload, answer in LIVE_SORTS[locale]["cases"]
                        if case_kind == kind)
    app = ExcelApplication(locale=locale)
    app.add_workbook()
    filled = "".join(f'    Cells({place + 1}, 3).Value = '
                     + " & ".join(f"ChrW(&H{ord(char):04X})" for char in text)
                     + f"\n    Cells({place + 1}, 4).Value = {place}\n" for place, text in enumerate(texts))
    source = (f'Function P() As String\n{filled}    Range("C1:D{len(texts)}").Sort Key1:=Range("C1"), '
              f'Order1:=xlAscending, Header:=xlNo, MatchCase:={matching}\n    Dim i As Long, s As String\n'
              f'    For i = 1 To {len(texts)}\n        s = s & "," & Cells(i, 4).Value\n    Next\n'
              '    P = Mid(s, 2)\nEnd Function\n')
    app.add_module(source, name="M")
    assert [int(place) for place in str(app.run("P")).split(",")] == order
