"""A calculated field's formula, as Excel reads it, spells it back and works it out.

A calculated field's formula is a worksheet formula over the table's fields. Measured with
scripts/measure_pivot_calculated_fields.py (tests/fixtures/pivot_calculated_fields/):

- A field is named as the table shows it, its caption, in apostrophes where its name needs them. Excel takes the
  apostrophes away where it does not and spells the name as the field is named: 'amount' is Amount. A cell
  reference, a name no field has, a field that leads back to itself and a function written apart from its bracket
  are refused.
- Excel spells the formula back with each function in capitals and each number as a formula spells it, 0.10 as 0.1
  and 1E3 as 1000. The space before a field moves to just after it, in place of any there, and goes where that is a
  comma or the end. Any other space stays, but one before a comma or at the end. The file keeps the formula with
  each field named as its cache names it, the table's captions aside.
- Each field stands for the sum of its values in the cell's records, a text field's being 0, and a calculated one
  for its own formula worked out there, whatever function its data field shows; a cell no record reaches works it
  out over zeros. TRUE comes out as 1 and FALSE as 0, text that reads as a number as that number, and any other text
  as #VALUE!.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._parse import literal
from pyopenvba.formula._r1c1 import is_cell, refused_in_a1
from pyopenvba.formula._values import ExcelError, number_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.formula._calc.nodes import Node
    from pyopenvba.formula._calc.values import Scalar


class Refused(Exception):
    """A formula Excel does not take for a calculated field."""


@dataclass(frozen=True, slots=True)
class Piece:
    """One token of a formula: its kind, its text as written, the white space before it, and for a field the place
    of the cache field it names."""

    kind: str
    text: str
    space: str = ''
    field: int = -1


_PIECE = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<number>(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)"
    r'|(?P<string>"(?:[^"]|"")*")'
    r"|(?P<quoted>'(?:[^']|'')+')"
    r"|(?P<error>#NULL!|#DIV/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N/A)"
    r"|(?P<word>[^\W\d][\w.]*)"
    r"|(?P<op><>|<=|>=|[-+*/^&=<>%])"
    r"|(?P<open>\()|(?P<close>\))|(?P<comma>,)", re.IGNORECASE)
_BARE = re.compile(r"[^\W\d][\w.]*")
#: The binary operators; % follows its operand and a leading + or - is a sign.
_BINARY = {'+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>='}
#: What a field stands for in the formula the engine works out: a name no formula can spell for itself.
_FIELD = '_PVF{}'


def read(text: str, resolve: Callable[[str], int | None]) -> list[Piece]:
    """A formula's pieces, each field found by ``resolve`` from the name the formula gives it, ending in an 'end'
    piece that holds the space at the end; Refused where Excel refuses the formula."""
    body = text[1:] if text.startswith('=') else text
    pieces: list[Piece] = []
    space = ''
    position = 0
    while position < len(body):
        found = _PIECE.match(body, position)
        if found is None:
            raise Refused(f'cannot read the formula at {body[position:position + 12]!r}')
        position = found.end()
        kind, written = found.lastgroup or '', found.group()
        if kind == 'ws':
            space += written
            continue
        if kind == 'word' and body[position:position + 1] != '(':
            index = resolve(written)
            if index is None and written.upper() not in ('TRUE', 'FALSE'):
                raise Refused(f'no field is named {written}')
            kind = 'field' if index is not None else 'bool'
            pieces.append(Piece(kind, written, space, -1 if index is None else index))
        elif kind == 'quoted':
            index = resolve(written[1:-1].replace("''", "'"))
            if index is None:
                raise Refused(f'no field is named {written}')
            pieces.append(Piece('field', written, space, index))
        else:
            pieces.append(Piece('function' if kind == 'word' else kind, written, space))
        space = ''
    pieces.append(Piece('end', '', space))
    if _expression(pieces, 0) != len(pieces) - 1:
        raise Refused('the formula does not read as one expression')
    return pieces


def _expression(pieces: list[Piece], at: int) -> int:
    at = _operand(pieces, at)
    while pieces[at].kind == 'op' and pieces[at].text in _BINARY:
        at = _operand(pieces, at + 1)
    return at


def _operand(pieces: list[Piece], at: int) -> int:
    while pieces[at].kind == 'op' and pieces[at].text in ('+', '-'):
        at += 1
    piece = pieces[at]
    if piece.kind in ('number', 'string', 'bool', 'error', 'field'):
        at += 1
    elif piece.kind == 'function' and pieces[at + 1].kind == 'open':
        at += 2
        while pieces[at].kind != 'close':
            if pieces[at].kind != 'comma':
                at = _expression(pieces, at)
            if pieces[at].kind == 'comma':
                at += 1
            elif pieces[at].kind != 'close':
                raise Refused('a function is not closed')
        at += 1
    elif piece.kind == 'open':
        at = _expression(pieces, at + 1)
        if pieces[at].kind != 'close':
            raise Refused('a bracket is not closed')
        at += 1
    else:
        raise Refused(f'{piece.text or "the end"} where a value should be')
    while pieces[at].kind == 'op' and pieces[at].text == '%':
        at += 1
    return at


def fields(pieces: list[Piece]) -> set[int]:
    """The cache fields a formula names."""
    return {piece.field for piece in pieces if piece.kind == 'field'}


def quoted(name: str) -> str:
    """A field's name as a formula spells it: in apostrophes where it is not a name a formula reads bare."""
    if _BARE.fullmatch(name) and not is_cell(name) and not refused_in_a1(name) \
            and name.upper() not in ('TRUE', 'FALSE'):
        return name
    return "'" + name.replace("'", "''") + "'"


def spelled(pieces: list[Piece], name_of: Callable[[int], str], *, given: bool = True) -> str:
    """The formula as Excel spells it back, without its =, each field named by ``name_of``: one ``given`` as a macro
    gives it, else one spelled already, as its cache keeps it, which keeps its spaces where they are."""
    out: list[str] = []
    moved: str | None = None
    for piece in pieces:
        space, moved = (piece.space if moved is None else moved), None
        if piece.kind == 'field' and given:
            # A field's space goes after it, in place of whatever space is there.
            space, moved = '', space
        if piece.kind in ('comma', 'end'):
            space = ''
        if piece.kind == 'number':
            text = number_text(literal(piece.text), formula=True)
        elif piece.kind in ('bool', 'error'):
            text = piece.text.upper()
        elif piece.kind == 'function':
            text = piece.text.upper() if is_excel_function(piece.text) else piece.text
        elif piece.kind == 'field':
            text = quoted(name_of(piece.field))
        else:
            text = piece.text
        out.append(space + text)
    return ''.join(out)


def without(pieces: list[Piece], gone: int) -> list[Piece]:
    """A formula once the field at ``gone`` is deleted: #NAME? where it named the field (tests/fixtures/
    pivot_calculated_fields/ delete_referenced)."""
    return [Piece('error', '#NAME?', piece.space) if piece.kind == 'field' and piece.field == gone else piece
            for piece in pieces]


def engine_node(pieces: list[Piece], sheet: Worksheet) -> Node:
    """The formula read into the engine's tree, each field a name bound to its value when the formula is worked out;
    Refused where the engine cannot read it."""
    from pyopenvba.formula._calc.lexer import FormulaSyntaxError

    text = ''.join(_FIELD.format(piece.field) if piece.kind == 'field' else piece.text for piece in pieces)
    try:
        return sheet.book.calculator.engine_book.read('=' + text)
    except FormulaSyntaxError as failure:
        raise Refused(str(failure)) from None


def worked_out(node: Node, values: dict[int, object], sheet: Worksheet) -> object:
    """What a formula comes to with each field at ``values``, a number or an error, as a calculated field's cell
    shows it."""
    from pyopenvba.apps.excel._evaluate import as_number
    from pyopenvba.formula._calc.cells import CellError

    def scalar(value: object) -> Scalar:
        if isinstance(value, ExcelError):
            return CellError(value.name)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return float(value)
        return 0.0

    bound = {_FIELD.format(index).upper(): scalar(value) for index, value in values.items()}
    found = as_number(sheet, node, bound)
    return ExcelError(found.code) if isinstance(found, CellError) else found
