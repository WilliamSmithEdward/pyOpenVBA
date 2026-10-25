"""Whole-string filter wildcards with bounded, non-backtracking transitions.

Callers retain their own type, text-admission and criterion-length policies.
In :func:`matches` single-character literal comparisons retain Python's Unicode
IGNORECASE rules; only escaped literal characters are compiled as regular
expressions. :func:`matches_by` compares them by a key the caller gives.
"""
from __future__ import annotations

import re
from collections.abc import Callable, Hashable
from functools import lru_cache


@lru_cache(maxsize=64)
def _parsed(pattern: str, escape_tilde: bool) -> tuple[int, int, tuple[tuple[str, int], ...], int]:
    """The places of a pattern's stars and question marks as bits, each literal character with the bits of its
    places, and the bit past the last place."""
    stars, questions, count = 0, 0, 0
    literals: dict[str, int] = {}
    index = 0
    escaped = '*?~' if escape_tilde else '*?'
    while index < len(pattern):
        char = pattern[index]
        literal = char == '~' and index + 1 < len(pattern) and pattern[index + 1] in escaped
        if literal:
            index += 1
            char = pattern[index]
        if not literal and char == '*':
            if count and stars & (1 << (count - 1)):
                index += 1
                continue
            stars |= 1 << count
        elif not literal and char == '?':
            questions |= 1 << count
        else:
            literals[char] = literals.get(char, 0) | (1 << count)
        count += 1
        index += 1
    return stars, questions, tuple(literals.items()), 1 << count


@lru_cache(maxsize=64)
def _compile(pattern: str, escape_tilde: bool) -> tuple[int, int, tuple[tuple[re.Pattern[str], int], ...], int]:
    stars, questions, literals, final = _parsed(pattern, escape_tilde)
    compiled = tuple((re.compile(re.escape(char), re.IGNORECASE), mask) for char, mask in literals)
    return stars, questions, compiled, final


def _literal_mask(char: str, literals: tuple[tuple[re.Pattern[str], int], ...]) -> int:
    mask = 0
    for literal, positions in literals:
        if literal.fullmatch(char) is not None:
            mask |= positions
    return mask


def _run(text: str, stars: int, questions: int, final: int, mask_of: Callable[[str], int]) -> bool:
    states = 1
    states |= (states & stars) << 1
    masks: dict[str, int] = {}
    for char in text:
        if char not in masks:
            masks[char] = mask_of(char)
        states = ((states & (questions | masks[char])) << 1) | (states & stars)
        states |= (states & stars) << 1
        if not states:
            return False
    return bool(states & final)


def matches(text: str, pattern: str, *, escape_tilde: bool = False) -> bool:
    """Match * and ? over Unicode characters, including newlines.

    Adjacent stars collapse, so one epsilon transition closes every active star.
    Each input character advances the bitset once. Literal comparisons are cached
    per distinct input character for this call, without retaining source strings.
    AutoFilter escapes only * and ?; other callers may explicitly escape ~ too.
    """
    stars, questions, literals, final = _compile(pattern, escape_tilde)
    return _run(text, stars, questions, final, lambda char: _literal_mask(char, literals))


def matches_by(text: str, pattern: str, key: Callable[[str], Hashable], *, escape_tilde: bool = False) -> bool:
    """As :func:`matches`, a literal standing for every character with its ``key``."""
    stars, questions, literals, final = _parsed(pattern, escape_tilde)
    keyed: dict[Hashable, int] = {}
    for char, mask in literals:
        found = key(char)
        keyed[found] = keyed.get(found, 0) | mask
    return _run(text, stars, questions, final, lambda char: keyed.get(key(char), 0))
