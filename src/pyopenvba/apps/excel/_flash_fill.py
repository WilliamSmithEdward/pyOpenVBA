"""Range.FlashFill, for the family of rules tests/fixtures/flash_fill.json measured: each example is built from whole words of the columns to
the left, their first letters, a change of case, and fixed text. One rule has to account for every example, or it is error 1004.

Excel's own inference is not published. A rule here is only what these parts can say, so a row a rule cannot be applied to, one missing
the word it asks for, says it is not implemented rather than guess what Excel would make of it.
"""
from __future__ import annotations

import itertools
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

_WORD = re.compile(r"[^\W_]+")
_MOST_RULES = 256


@dataclass(frozen=True)
class _Part:
    """Fixed text, or a word of a column: counted from the start or, negative, from the end; all of it or its first letters."""

    text: str = ""
    column: int = -1
    word: int = 0
    letters: int = 0
    case: str = ""

    def applied(self, words: list[list[str]]) -> str | None:
        if self.column < 0:
            return self.text
        source = words[self.column]
        if not -len(source) <= self.word < len(source):
            return None
        found = source[self.word]
        found = found[:self.letters] if self.letters else found
        return found.upper() if self.case == "upper" else found.lower() if self.case == "lower" else found


def _case_between(found: str, wanted: str) -> str | None:
    if found == wanted:
        return ""
    if found.upper() == wanted:
        return "upper"
    return "lower" if found.lower() == wanted else None


def _word_at(words: list[list[str]], example: str, at: int) -> list[list[_Part]] | None:
    """The ways a word's whole text, or failing that its first letters, can account for the example at this place; each with how long."""
    remaining = example[at:]
    best: list[_Part] = []
    longest = 0
    for column, source in enumerate(words):
        for place, word in enumerate(source):
            common = 0
            while common < min(len(word), len(remaining)) and word[common].lower() == remaining[common].lower():
                common += 1
            if not common:
                continue
            whole = common == len(word)
            # Part of a word is its first letter, as in initials, or runs to the end of what it stands for.
            if 1 < common < len(word) and common < len(remaining) and remaining[common].isalnum():
                continue
            size = common * 2 + (1 if whole else 0)
            case = _case_between(word[:common], remaining[:common])
            if case is None or size < longest:
                continue
            if size > longest:
                best, longest = [], size
            letters = 0 if whole else common
            forms = [_Part(column=column, word=place, letters=letters, case=case)]
            if place == len(source) - 1 and place:
                forms.insert(0, _Part(column=column, word=place - len(source), letters=letters, case=case))
            best.extend(forms)
    return [best] if best else None


def _rules_from(words: list[list[str]], example: str) -> list[list[list[_Part]]]:
    """What the example could be built from, each place's choices in order."""
    slots: list[list[_Part]] = []
    at = 0
    while at < len(example):
        found = _word_at(words, example, at) if example[at].isalnum() else None
        if found is None:
            if slots and len(slots[-1]) == 1 and slots[-1][0].column < 0:
                slots[-1] = [_Part(text=slots[-1][0].text + example[at])]
            else:
                slots.append([_Part(text=example[at])])
            at += 1
            continue
        choice = found[0]
        slots.append(choice)
        size = choice[0].letters or len(words[choice[0].column][choice[0].word])
        at += size
    return [slots]


def _built(rule: tuple[_Part, ...], words: list[list[str]]) -> str | None:
    pieces: list[str] = []
    for part in rule:
        piece = part.applied(words)
        if piece is None:
            return None
        pieces.append(piece)
    return "".join(pieces)


def _rule_for(examples: list[tuple[list[list[str]], str]]) -> tuple[_Part, ...] | None:
    first_words, first_example = examples[0]
    for slots in _rules_from(first_words, first_example):
        for rule in itertools.islice(itertools.product(*slots), _MOST_RULES):
            if all(_built(rule, words) == example for words, example in examples):
                return rule
    return None


def _text(sheet: Worksheet, row: int, column: int) -> str:
    from pyopenvba.apps.excel._model import Range

    return str(Range(sheet, [Area(row, column, row, column, sheet.name)]).vba_get("Text"))


def flash_fill(target: Range) -> object:
    from pyopenvba.apps.excel._model import Range

    sheet, area = target.sheet, target.first
    if len(target.areas) != 1 or area.columns != 1:
        raise VBAUnsupportedError("FlashFill on anything but one column is not implemented")
    column = area.left
    columns = list(range(1, column))
    while columns and not any(position[1] == columns[0] for position in sheet.cells_):
        columns.pop(0)
    if not columns:
        raise error(1004, "FlashFill method of Range class failed")

    def words_of(row: int) -> list[list[str]]:
        return [_WORD.findall(_text(sheet, row, one)) for one in columns]

    def has_input(row: int) -> bool:
        return any(_text(sheet, row, one) for one in columns)

    bottom = area.bottom
    if area.rows == 1:
        # One cell runs as far as the data beside it does.
        while has_input(bottom + 1):
            bottom += 1
    examples = [(words_of(row), _text(sheet, row, column)) for row in range(area.top, bottom + 1) if _text(sheet, row, column)]
    if not examples:
        raise error(1004, "FlashFill method of Range class failed")
    rule = _rule_for(examples)
    if rule is None:
        raise error(1004, "FlashFill method of Range class failed")
    for row in range(area.top, bottom + 1):
        if _text(sheet, row, column) or not has_input(row):
            continue
        made = _built(rule, words_of(row))
        if made is None:
            raise VBAUnsupportedError("FlashFill on a row without the word its rule asks for is not implemented")
        Range(sheet, [Area(row, column, row, column, sheet.name)]).vba_set("Value", made)
    return EMPTY
