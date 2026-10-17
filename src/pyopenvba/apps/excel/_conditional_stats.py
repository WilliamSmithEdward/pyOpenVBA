"""Sparse statistical rule populations, shared only within one display read."""
from __future__ import annotations

import datetime as dt
import math
import unicodedata
from collections import Counter
from collections.abc import Generator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from itertools import pairwise
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._engine_book import scalar
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc import collate
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.numbers import number_text
from pyopenvba.formula._calc.values import Area, Empty, Scalar, text_to_number

if TYPE_CHECKING:
    from pyopenvba._a1 import Area as ModelArea
    from pyopenvba.apps.excel._conditional import Rule
    from pyopenvba.apps.excel._model import Worksheet

Key = tuple[str, str]


@dataclass
class Population:
    numbers: tuple[float, ...]
    counts: Counter[Key]
    text_counts: Counter[str]
    pattern_counts: dict[str, int]
    today: dt.date
    signature: tuple[object, ...]
    average: float
    deviation: float


_READ: ContextVar[dict[tuple[int, int], Population | None] | None] = ContextVar('conditional_statistics_read', default=None)


@contextmanager
def conditional_read() -> Generator[None, None, None]:
    """Share populations in a synchronous property read, never across reads."""
    if _READ.get() is not None:
        yield
        return
    token = _READ.set({})
    try:
        yield
    finally:
        _READ.reset(token)


def _fold(value: str) -> str:
    # Merge explicit VBA surrogate pairs before canonical Unicode comparison.
    text = value.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'surrogatepass')
    return unicodedata.normalize('NFC', text).casefold()


def _units(value: str) -> str:
    """Wildcard question marks consume UTF-16 units, including half a pair."""
    if not any(ord(char) > 0xffff for char in value):
        return value
    encoded = value.encode('utf-16-le', 'surrogatepass')
    return ''.join(chr(encoded[index] | encoded[index + 1] << 8) for index in range(0, len(encoded), 2))


def _wildcard(pattern: str, text: str) -> bool:
    """Match raw UTF-16 units; canonical/expanding equality is a different API."""
    pattern, text = _units(pattern), _units(text)
    tokens: list[tuple[str, str]] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == '~' and index + 1 < len(pattern) and pattern[index + 1] in '*?~':
            index += 1
            tokens.append(('literal', pattern[index]))
        elif char == '*':
            tokens.append(('star', ''))
        elif char == '?':
            tokens.append(('one', ''))
        else:
            tokens.append(('literal', char))
        index += 1
    ends: set[int] = {0}
    for kind, char in tokens:
        if not ends:
            return False
        if kind == 'star':
            ends = set(range(min(ends), len(text) + 1))
        elif kind == 'one':
            ends = {end + 1 for end in ends if end < len(text)}
        else:
            ends = {end + 1 for end in ends if end < len(text) and text[end].casefold() == char.casefold()}
    return len(text) in ends


def _key(value: Scalar, today: dt.date) -> Key:
    if isinstance(value, bool):
        return ('logical', str(value))
    if isinstance(value, CellError):
        return ('error', value.code)
    if isinstance(value, Empty):
        return ('value', '')
    if isinstance(value, str):
        number = text_to_number(value, today)
        if number is None or not math.isfinite(number):
            return ('value', _fold(value))
    else:
        number = float(value)
    return ('number', number_text(number) if number else '0')


def _area_size(areas: list[ModelArea]) -> int:
    """Union size without expanding even whole-column or whole-sheet ranges."""
    boundaries = sorted({value for area in areas for value in (area.top, area.bottom + 1)})
    total = 0
    for top, bottom in pairwise(boundaries):
        intervals = sorted((area.left, area.right + 1) for area in areas if area.top <= top <= area.bottom)
        edge, width = 0, 0
        for left, right in intervals:
            width += max(0, right - max(edge, left))
            edge = max(edge, right)
        total += (bottom - top) * width
    return total


def _population(sheet: Worksheet, rule: Rule) -> Population:
    signature = _signature(rule)
    engine = sheet.book.calculator.engine_book
    held: dict[tuple[int, int], Scalar] = {}
    for area in rule.areas:
        for row, column, value in engine.cells(Area(sheet.name, area.top, area.left, area.bottom, area.right)):
            held[(row, column)] = value
    numbers = sorted(value for value in held.values() if isinstance(value, float))
    today = sheet.book.calculator.now().date()
    counts = Counter(_key(value, today) for value in held.values())
    counts[_key(Empty(), today)] += _area_size(rule.areas) - len(held)
    texts = Counter(value for value in held.values() if isinstance(value, str))
    average, deviation = 0.0, 0.0
    if numbers:
        # Scale before squaring to avoid overflowing finite native doubles.
        scale = max(abs(numbers[0]), abs(numbers[-1]))
        if scale:
            center = math.fsum(value / scale for value in numbers) / len(numbers)
            average = center * scale
            deviation = math.sqrt(math.fsum((value / scale - center) ** 2 for value in numbers) / len(numbers)) * scale
    if _signature(rule) != signature:
        raise VBAUnsupportedError('Changing conditional rules during a statistical display read is not implemented')
    return Population(tuple(numbers), counts, texts, {}, today, signature, average, deviation)


def _signature(rule: Rule) -> tuple[object, ...]:
    return tuple(rule.areas), tuple(sorted(rule.root.attrib.items())), rule.deleted


def _pattern_count(value: str, population: Population) -> int:
    """Only raw wildcard characters activate COUNTIF-style operator parsing."""
    operator, pattern = '=', value
    for candidate in ('<=', '>=', '<>', '=', '<', '>'):
        if value.startswith(candidate):
            operator, pattern = candidate, value[len(candidate):]
            break
    if operator in ('=', '<>'):
        count = sum(n for text, n in population.text_counts.items() if _wildcard(pattern, text))
        return count if operator == '=' else sum(population.counts.values()) - count
    count = 0
    for text, number in population.text_counts.items():
        order = collate.compare(text, pattern)
        if ((operator == '<' and order < 0) or (operator == '>' and order > 0)
                or (operator == '<=' and order <= 0) or (operator == '>=' and order >= 0)):
            count += number
    return count


def matches(sheet: Worksheet, rule: Rule, row: int, column: int) -> bool:
    """Native top/bottom, average/deviation and duplicate/unique predicates."""
    bank = _READ.get()
    if bank is None:
        with conditional_read():
            return matches(sheet, rule, row, column)
    key = id(sheet), id(rule)
    population = bank.get(key)
    if population is None:
        if key in bank:
            raise VBAUnsupportedError('Recursive statistical conditional formatting evaluation is not implemented')
        bank[key] = None
        try:
            population = _population(sheet, rule)
        except BaseException:
            bank.pop(key, None)
            raise
        bank[key] = population
    if _signature(rule) != population.signature:
        raise VBAUnsupportedError('Changing conditional rules during a statistical display read is not implemented')
    value = scalar(sheet.book.calculator.value_of(sheet.name, row, column))
    if _signature(rule) != population.signature:
        raise VBAUnsupportedError('Changing conditional rules during a statistical display read is not implemented')
    kind = rule.root.get('type', '')
    if kind in ('uniqueValues', 'duplicateValues'):
        if isinstance(value, Empty):
            return False  # Blank cells count as partners but never paint themselves.
        if isinstance(value, str) and any(char in value for char in '*?'):
            pattern = value
            count = population.pattern_counts.get(pattern)
            if count is None:
                count = _pattern_count(pattern, population)
                population.pattern_counts[pattern] = count
        else:
            count = population.counts[_key(value, population.today)]
        return count == 1 if kind == 'uniqueValues' else count > 1
    if not isinstance(value, float) or not population.numbers:
        return False
    if kind == 'top10':
        rank = int(rule.root.get('rank', '10'))
        if rule.root.get('percent', '0') in ('1', 'true'):
            rank = max(1, len(population.numbers) * rank // 100)
        rank = min(len(population.numbers), max(1, rank))
        if rule.root.get('bottom', '0') in ('1', 'true'):
            return value <= population.numbers[rank - 1]
        return value >= population.numbers[-rank]
    above = rule.root.get('aboveAverage', '1') in ('1', 'true')
    deviation = rule.root.get('stdDev')
    if deviation is not None:
        distance = int(deviation) * population.deviation
        return value >= population.average + distance if above else value <= population.average - distance
    if rule.root.get('equalAverage', '0') in ('1', 'true'):
        return value >= population.average if above else value <= population.average
    return value > population.average if above else value < population.average
