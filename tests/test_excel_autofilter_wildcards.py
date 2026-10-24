"""Wildcard compatibility and bounded-work regressions for AutoFilter."""
from __future__ import annotations

import itertools
import os
import re
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.apps.excel import _autofilter, _filter_wildcards
from pyopenvba.exceptions import VBAUnsupportedError


def legacy(text: str, pattern: str) -> bool:
    parts: list[str] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == '~' and index + 1 < len(pattern) and pattern[index + 1] in '*?':
            parts.append(re.escape(pattern[index + 1]))
            index += 2
            continue
        parts.append('.*' if char == '*' else '.' if char == '?' else re.escape(char))
        index += 1
    return re.fullmatch(''.join(parts), text, re.IGNORECASE | re.DOTALL) is not None


def test_short_patterns_preserve_legacy_literal_escape_and_wildcard_semantics() -> None:
    texts = ['', 'a', 'A', 'b', '*', '?', '~', '~~', '~*', '~a', 'aa', 'ab', 'a\nb', '\n']
    for length in range(5):
        for chars in itertools.product('a*?~', repeat=length):
            pattern = ''.join(chars)
            for text in texts:
                assert _filter_wildcards.matches(text, pattern) == legacy(text, pattern), (text, pattern)


@pytest.mark.parametrize('text', ['[a].+($)^|', '\r\n', '\x00', '\U0001f600', '\u2160', '\u2170', '\u0345', '\u03b9', '\u0130', 'i', '\u017f', 's', '\u00df'])
def test_unicode_literal_case_and_character_width_match_legacy(text: str) -> None:
    for pattern in [text, text.lower(), text.upper(), '*', '?', '??', '*' + text + '*']:
        assert _filter_wildcards.matches(text, pattern) == legacy(text, pattern), (text, pattern)


@pytest.mark.parametrize('shown,text,pattern,expected', [
    ('2', None, '2', True), ('2', None, '*2', False), ('TRUE', None, '*', False),
    ('#N/A', None, '#N/A', True), ('#N/A', None, '*', False), ('', None, '', True),
    ('*', None, '~*', True), ('~*', None, '~~*', True), ('a\nb', 'a\nb', 'a?b', True),
    ('a' * 300, 'a' * 300, 'a' * 300, True),
])
def test_autofilter_type_gating_newlines_and_no_new_length_cap(
        shown: str, text: str | None, pattern: str, expected: bool) -> None:
    seen = _autofilter.Seen(shown, None, text, False)
    assert getattr(_autofilter, '_like')(seen, pattern) is expected


def test_autofilter_keeps_existing_non_ascii_letter_admission() -> None:
    seen = _autofilter.Seen('caf\u00e9', None, 'caf\u00e9', False)
    with pytest.raises(VBAUnsupportedError):
        getattr(_autofilter, '_like')(seen, '*')


def test_optional_tilde_escape_is_explicit() -> None:
    assert _filter_wildcards.matches('~anything', '~~*', escape_tilde=True)
    assert not _filter_wildcards.matches('~anything', '~~*', escape_tilde=False)
    assert _filter_wildcards.matches('~*', '~~*', escape_tilde=False)


def test_repeated_star_literal_masks_are_evaluated_once_per_input_character(monkeypatch: pytest.MonkeyPatch) -> None:
    original = cast('Callable[[str, tuple[tuple[re.Pattern[str], int], ...]], int]',
                    getattr(_filter_wildcards, '_literal_mask'))
    calls: list[str] = []

    def count(char: str, literals: tuple[tuple[re.Pattern[str], int], ...]) -> int:
        calls.append(char)
        return original(char, literals)

    monkeypatch.setattr(_filter_wildcards, '_literal_mask', count)
    assert not _filter_wildcards.matches('a' * 2048, '*a' * 60 + 'b')
    assert calls == ['a']


def test_repeated_star_miss_finishes_in_isolated_process() -> None:
    source = ("from pyopenvba.apps.excel._autofilter import Seen, _like;"
              "assert not _like(Seen('a'*2048,None,'a'*2048,False),'*a'*60+'b')")
    environment = dict(os.environ, PYTHONPATH=str(Path(__file__).parents[1] / 'src'), PYTHONDONTWRITEBYTECODE='1')
    result = subprocess.run([sys.executable, '-B', '-c', source], env=environment,
                            capture_output=True, text=True, timeout=10, check=False)
    assert result.returncode == 0, result.stderr
