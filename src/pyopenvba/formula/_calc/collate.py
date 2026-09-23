"""How Excel orders and matches text in a filter.

Measured against Excel's own filters rather than taken from a sort
specification: a column of words, a ``>`` criterion per word, and the rows
Excel kept recorded, so every pair's order is Excel's answer. What that
showed is Windows' word sort with case ignored.

**Letters by their base, accents only to break a tie.** ``e < é < f``
and ``resume < résume < résumé``: an accent matters only between words
that are otherwise the same, and then from the left. Case never matters,
and a ligature is its letters: ``æ`` is ``ae`` and ``ß`` is ``ss``, equal
and not merely adjacent.

**Symbols first, then digits, then letters.** Space sorts lowest, then
the no-break space, then ASCII punctuation in the order below, then the
common non-ASCII symbols, all before ``0``. Greek follows Latin.

**Apostrophe and the hyphens are ignored until the very end.** ``coop <
co-op < cop``: the hyphen does not separate ``co`` from ``op``. It only
decides between words that are otherwise equal, and there ``'`` comes
before ``-``, then the en dash, then the em dash. So ``coop`` and ``co-op``
are not equal, which matters to a value list.

**Wildcards match character by character.** ``*`` any run, ``?`` one
character, ``~`` escapes either or itself. Case is ignored; accents and
ligatures are not, so ``a*e`` matches ``ae`` but not ``æ``, and ``a*`` does
not match ``ábc``.

Trimming is not collation: Excel trims a cell's text before comparing it
and compares a criterion as it is written. That happens in the filter, not
here.
"""

from __future__ import annotations

import functools
import unicodedata

#: Symbols in the order Excel sorts them, all below the digits. Space and
#: the no-break space lead; the rest was measured one pair at a time.
_SYMBOLS = (
    " \u00a0!\"#$%&()*,./:;?@[\\]^_`{|}~\u00bf\u2019\u00a3\u20ac+<=>\u00b1\u00ab\u00a7\u00a9\u00b0\u2026"
)
_SYMBOL_WEIGHT = {symbol: index + 1 for index, symbol in enumerate(_SYMBOLS)}

#: Ignored at every level but the last, in this order there.
_IGNORABLE = {"'": 1, "-": 2, "\u2010": 2, "\u2011": 2, "\u2013": 3, "\u2014": 4}
#: Ignored entirely.
_INVISIBLE = frozenset({"\u00ad", "\u200b", "\u200c", "\u200d", "\ufeff"})

#: Letters that are other letters: ligatures expand, and letters with a
#: stroke are their base with an accent heavier than any combining mark.
#: The stroked ones are o, d, l, h, t and b with a stroke, and dotless i.
_EXPANSIONS = {"æ": "ae", "œ": "oe", "ß": "ss", "ĳ": "ij"}
_STROKED = {"\u00f8": "o", "\u0111": "d", "\u0142": "l", "\u0127": "h", "\u0167": "t", "\u0180": "b", "\u0131": "i"}
_STROKE_WEIGHT = 100

#: Combining marks in the Unicode default order; those not listed follow.
_MARKS = [
    "\u0301", "\u0300", "\u0306", "\u0302", "\u030c", "\u030a", "\u0308", "\u030b", "\u0303",
    "\u0307", "\u0327", "\u0328", "\u0304",
]
_MARK_WEIGHT = {mark: index + 1 for index, mark in enumerate(_MARKS)}

# Primary weight bands, each above the last.
_UNLISTED_SYMBOL = 1_000
_DIGIT = 2_000_000
_LATIN = 3_000_000
_OTHER_LETTER = 4_000_000

#: A text's place in the order: primary weights, then accents per primary,
#: then the ignorables and where they stood.
SortKey = tuple[tuple[int, ...], tuple[tuple[int, ...], ...], tuple[tuple[int, int], ...]]


@functools.lru_cache(maxsize=65536)
def sort_key(text: str) -> SortKey:
    """The key Excel's text order sorts by; equal keys are equal texts."""
    primary: list[int] = []
    accents: list[list[int]] = []
    tail: list[tuple[int, int]] = []
    for char in unicodedata.normalize("NFD", text):
        if char in _INVISIBLE:
            continue
        if char in _IGNORABLE:
            tail.append((len(primary), _IGNORABLE[char]))
            continue
        if unicodedata.combining(char):
            if accents:
                accents[-1].append(_MARK_WEIGHT.get(char, len(_MARKS) + 1 + ord(char)))
            continue
        # A capital can lower to a letter and a mark, as "İ" does: keep the letter.
        lowered = char.lower()[0]
        if lowered in _EXPANSIONS:
            for letter in _EXPANSIONS[lowered]:
                primary.append(_weight(letter))
                accents.append([])
            continue
        if lowered in _STROKED:
            primary.append(_weight(_STROKED[lowered]))
            accents.append([_STROKE_WEIGHT])
            continue
        primary.append(_weight(lowered))
        accents.append([])
    return tuple(primary), tuple(tuple(marks) for marks in accents), tuple(tail)


def _weight(char: str) -> int:
    if char in _SYMBOL_WEIGHT:
        return _SYMBOL_WEIGHT[char]
    if char.isdigit():
        return _DIGIT + unicodedata.digit(char, 0)
    if "a" <= char <= "z":
        return _LATIN + ord(char) - ord("a")
    if char.isalpha():
        return _OTHER_LETTER + ord(char)
    return _UNLISTED_SYMBOL + ord(char)


def compare(left: str, right: str) -> int:
    """-1, 0 or 1 as ``left`` sorts before, level with or after ``right``."""
    a, b = sort_key(left), sort_key(right)
    return (a > b) - (a < b)


def equal(left: str, right: str) -> bool:
    """Whether Excel treats two texts as the same value: every level of
    the order agrees, so case differs and nothing else does."""
    return sort_key(left) == sort_key(right)


def has_wildcards(pattern: str) -> bool:
    """Whether a criterion has an unescaped ``*`` or ``?``."""
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "~":
            index += 2
            continue
        if char in "*?":
            return True
        index += 1
    return False


def unescape(pattern: str) -> str:
    """A criterion with its ``~`` escapes resolved, for use as a value."""
    out: list[str] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "~" and index + 1 < len(pattern) and pattern[index + 1] in "*?~":
            out.append(pattern[index + 1])
            index += 2
            continue
        out.append(char)
        index += 1
    return "".join(out)


def wildcard_match(pattern: str, text: str) -> bool:
    """Whether ``text`` matches ``pattern`` the way a filter does: all of
    it, ``*`` any run, ``?`` one character, case ignored."""
    tokens = _pattern_tokens(pattern)
    lowered = text.lower()
    # Positions in the text each token sequence prefix can end at.
    ends: set[int] = {0}
    for kind, char in tokens:
        if not ends:
            return False
        if kind == "star":
            ends = set(range(min(ends), len(lowered) + 1))
        elif kind == "one":
            ends = {end + 1 for end in ends if end < len(lowered)}
        else:
            ends = {end + 1 for end in ends if end < len(lowered) and lowered[end] == char}
    return len(lowered) in ends


def _pattern_tokens(pattern: str) -> list[tuple[str, str]]:
    tokens: list[tuple[str, str]] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "~" and index + 1 < len(pattern) and pattern[index + 1] in "*?~":
            tokens.append(("char", pattern[index + 1].lower()))
            index += 2
            continue
        if char == "*":
            if not tokens or tokens[-1][0] != "star":
                tokens.append(("star", ""))
        elif char == "?":
            tokens.append(("one", ""))
        else:
            tokens.append(("char", char.lower()))
        index += 1
    return tokens


__all__ = ["SortKey", "compare", "equal", "has_wildcards", "sort_key", "unescape", "wildcard_match"]
