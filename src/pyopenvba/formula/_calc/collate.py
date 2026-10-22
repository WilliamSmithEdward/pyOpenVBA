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
from collections.abc import Iterator

#: Symbols in the order Excel sorts them, all below the digits. Space and
#: the no-break space lead; the rest was measured one pair at a time.
_SYMBOLS = (
    " \u00a0!\"#$%&()*,./:;?@[\\]^_`{|}~\u00bf\u2019\u00a3\u20ac+<=>\u00b1\u00ab\u00a7\u00a9\u00b0\u2026"
)
_SYMBOL_WEIGHT = {symbol: index + 1 for index, symbol in enumerate(_SYMBOLS)}

#: Ignored at every level but the last, in this order there.
_IGNORABLE = {"'": 1, "-": 2, "\u2010": 2, "\u2011": 2, "\u2013": 3, "\u2014": 4}
#: Ignored entirely. The zero-width space is not: measured, a filter tells
#: ``a\u200bb`` from ``ab`` and a sort puts it first.
_INVISIBLE = frozenset({"\u00ad", "\u200c", "\u200d", "\ufeff"})
#: Letters that are another letter whatever their case: measured, a filter
#: keeping ``\u03c3`` keeps the final ``\u03c2`` and ``\u03a3`` too.
_FOLDED = {"\u03c2": "\u03c3"}
#: Forms that are their base character drawn another way, full width or
#: raised or lowered: measured, a filter tells ``\uff11`` from ``1`` and ``\u00b2``
#: from ``2``, and a sort puts each just after its base, as it puts a
#: letter with a stroke after its letter.
_VARIANT_TAGS = ("<wide>", "<super>", "<sub>")
_VARIANT_WEIGHT = 200

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
CollationKey = tuple[tuple[int, ...], tuple[tuple[int, ...], ...], tuple[tuple[int, int], ...]]


@functools.lru_cache(maxsize=65536)
def sort_key(text: str) -> CollationKey:
    """The key Excel's text order sorts by; equal keys are equal texts."""
    primary: list[int] = []
    accents: list[list[int]] = []
    tail: list[tuple[int, int]] = []
    for char in _characters(text):
        if char in _INVISIBLE:
            continue
        if char in _IGNORABLE:
            tail.append((len(primary), _IGNORABLE[char]))
            continue
        if unicodedata.combining(char):
            if accents:
                accents[-1].append(_MARK_WEIGHT.get(char, len(_MARKS) + 1 + ord(char)))
            continue
        base = _variant_base(char)
        lowered = _lowered(char if base is None else base)
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
        accents.append([] if base is None else [_VARIANT_WEIGHT])
    return tuple(primary), tuple(tuple(marks) for marks in accents), tuple(tail)


def _lowered(char: str) -> str:
    # A capital can lower to a letter and a mark, as "İ" does: keep the letter.
    lowered = char.lower()[0]
    return _FOLDED.get(lowered, lowered)


def _variant_base(char: str) -> str | None:
    """The character a full-width, raised or lowered form is drawn from."""
    if not unicodedata.decomposition(char).startswith(_VARIANT_TAGS):
        return None
    base = unicodedata.normalize("NFKD", char)
    return base if len(base) == 1 else None


def _characters(text: str) -> Iterator[str]:
    """``text`` decomposed as the order reads it: canonically, and a letter
    that is other letters, a ligature or a digraph, as those letters.
    Measured in Excel's sort, ``ﬁ`` is ``fi`` and ``ǅ`` is ``dž``."""
    for char in unicodedata.normalize("NFD", text):
        if char.isalpha() and unicodedata.decomposition(char).startswith("<compat>"):
            yield from unicodedata.normalize("NFKD", char)
        else:
            yield char


#: A text's place in the order when case matters: :data:`CollationKey`'s
#: levels, with each letter's case between the accents and the ignorables.
CaseCollationKey = tuple[tuple[int, ...], tuple[tuple[int, ...], ...], tuple[int, ...], tuple[tuple[int, int], ...]]


@functools.lru_cache(maxsize=65536)
def case_sort_key(text: str) -> CaseCollationKey:
    """The key Excel's sort orders text by when told to match case.

    Measured with Sort's MatchCase: case decides only between texts equal
    in their letters and accents, from the left, a lowercase letter before
    a title-case one, as ``ǅ`` is, before a capital. It decides before the
    hyphens and apostrophes do, so ``ab < a-b < aB < Ab``.
    """
    primary, accents, tail = sort_key(text)
    cases: list[int] = []
    for char in _characters(text):
        if char in _INVISIBLE or char in _IGNORABLE or unicodedata.combining(char):
            continue
        # Measured, final sigma comes between its lowercase and its capital,
        # as a title-case letter does.
        weight = 2 if char.isupper() else 1 if char.istitle() or char in _FOLDED else 0
        base = _variant_base(char)
        lowered = _lowered(char if base is None else base)
        cases.extend([weight] * len(_EXPANSIONS.get(lowered, lowered)))
    return primary, accents, tuple(cases), tail


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


__all__ = [
    "CaseCollationKey",
    "CollationKey",
    "case_sort_key",
    "compare",
    "equal",
    "has_wildcards",
    "sort_key",
    "unescape",
    "wildcard_match",
]
