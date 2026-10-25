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
from pyopenvba.apps.excel._sort import text_key
from pyopenvba.exceptions import VBAUnsupportedError

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


@pytest.mark.parametrize("text", ["\u0301a", "\u4e2d", "\u0385"], ids=["leading mark", "CJK", "dialytika tonos"])
def test_text_the_model_does_not_sort_reports_itself(text: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        text_key(text, False)


@pytest.mark.skipif(sys.platform != "win32", reason="CompareStringEx is Windows'")
def test_fresh_strings_sort_as_compare_string_orders_them() -> None:
    compare = _windows.WinDLL("kernel32").CompareStringEx
    compare.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_wchar_p, ctypes.c_int, ctypes.c_wchar_p,
                        ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
    from pyopenvba.apps.excel._sort_weights import WEIGHTS

    pick = random.Random()
    alphabet = sorted({character for text, _ignoring, _matching in KEYS for character in text})
    # Text beginning with what Windows weighs as a combining mark, a diacritic and no primary weight, reports itself
    # unsupported (see the test above).
    marks = {character for character in alphabet if not WEIGHTS[ord(character)][0] and WEIGHTS[ord(character)][1]}
    missed: list[tuple[str, str, bool]] = []
    for _ in range(20000):
        first, second = ("".join(pick.choice(alphabet) for _ in range(pick.randint(1, 6))) for _ in range(2))
        if first[0] in marks or second[0] in marks:
            continue
        for matching in (False, True):
            mine = (text_key(first, matching) > text_key(second, matching)) - \
                (text_key(first, matching) < text_key(second, matching))
            theirs = compare("en-US", 0 if matching else 1, first, -1, second, -1, None, None, None) - 2
            if mine != theirs:
                missed.append((first, second, matching))
    assert not missed, missed[:3]
