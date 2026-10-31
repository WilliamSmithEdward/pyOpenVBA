"""PivotTable.GetData: the value a pivot table shows for a name that picks a data field and items.

Measured in live Excel (scripts/measure_pivot_get_data.py, tests/fixtures/pivot_get_data.json): the name is words
and single-quoted names, in any order and case. A run of plain words reads as the longest name it starts with, so
``Sum of v a`` is the data field Sum of v and the item a. An item is named alone, by the one field that has it, or
as ``field[item]``, the item quoted or not; ``Grand Total`` picks nothing. A name that picks no data field where
the table has several, names what the table lacks, or picks a cell the table does not show is error 1004.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc import collate
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivot_reports import PivotReport


@dataclass(frozen=True)
class _Token:
    text: str
    quoted: bool = False
    #: For ``field[item]``, the field.
    field: str | None = None


def _quoted(text: str, at: int) -> tuple[str, int]:
    """The name quoted from ``at``, a doubled quote standing for one, and where it ends."""
    found: list[str] = []
    at += 1
    while at < len(text):
        if text[at] == "'":
            if text[at + 1:at + 2] == "'":
                found.append("'")
                at += 2
                continue
            return ''.join(found), at + 1
        found.append(text[at])
        at += 1
    raise error(1004)


def _tokens(text: str) -> list[_Token]:
    tokens: list[_Token] = []
    at = 0
    while at < len(text):
        if text[at] == ' ':
            at += 1
        elif text[at] == "'":
            name, at = _quoted(text, at)
            tokens.append(_Token(name, quoted=True))
        else:
            end = at
            while end < len(text) and text[end] not in ' [':
                end += 1
            word = text[at:end]
            if end < len(text) and text[end] == '[':
                if text[end + 1:end + 2] == "'":
                    item, close = _quoted(text, end + 1)
                else:
                    close = text.find(']', end + 1)
                    if close < 0:
                        raise error(1004)
                    item = text[end + 1:close]
                if close >= len(text) or text[close] != ']':
                    raise error(1004)
                tokens.append(_Token(item, field=word))
                at = close + 1
            else:
                tokens.append(_Token(word))
                at = end
    return tokens


def _item(report: PivotReport, name: str) -> tuple[int, int] | None:
    """The field among the table's that has an item so named, and the item; None where none has."""
    found = [(field, item) for field in (*report.rows, *report.columns, *(page for page, _ in report.pages))
             if (item := report.item_index(field, name)) is not None]
    if len(found) > 1:
        raise VBAUnsupportedError('PivotTable.GetData of an item more than one field has is not implemented')
    return found[0] if found else None


def picked(report: PivotReport, text: str) -> tuple[int, dict[int, int]]:
    """The data field and the items a name picks, by field."""
    tokens = _tokens(text)
    value: int | None = None
    chosen: dict[int, int] = {}

    def choose(field: int, item: int) -> None:
        if chosen.get(field, item) != item:
            raise error(1004)
        chosen[field] = item

    at = 0
    while at < len(tokens):
        token = tokens[at]
        if token.field is not None:
            field = report.field_index(token.field)
            item = report.item_index(field, token.text) if field is not None else None
            if field is None or item is None:
                raise error(1004)
            choose(field, item)
            at += 1
            continue
        # A run of plain words is the longest name it starts with.
        run = 1
        if not token.quoted:
            while at + run < len(tokens) and not tokens[at + run].quoted and tokens[at + run].field is None:
                run += 1
        for length in range(run, 0, -1):
            name = ' '.join(one.text for one in tokens[at:at + length])
            if value is None and (found := report.value_index(name)) is not None:
                value = found
            elif collate.equal(name, 'Grand Total'):
                pass
            elif (pair := _item(report, name)) is not None:
                choose(*pair)
            else:
                continue
            at += length
            break
        else:
            raise error(1004)
    if value is None:
        if len(report.values) == 1:
            raise VBAUnsupportedError('PivotTable.GetData naming no data field of a table with one is not implemented')
        raise error(1004)
    return value, chosen
