"""Reading an Excel formula.

The grammar is not VBA's: ``&`` is the only concatenation, ``^`` binds
tighter than unary minus in one direction and looser in the other,
``%`` follows its operand, a reference is a value, and TRUE is a
constant rather than a keyword.

What is parsed here is the text after the leading ``=``.  A formula
that this cannot read raises :class:`FormulaError`, which is a compile
error rather than a cell error: Excel refuses to accept such a formula
at all rather than showing an error value.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Final

from pyopenvba.exceptions import PyOpenVBAError
from pyopenvba.formula._values import ERRORS, ExcelError


class FormulaError(PyOpenVBAError):
    """Raised for a formula Excel itself would not accept."""


# --- the tree -----------------------------------------------------------------------


@dataclass(slots=True)
class Node:
    pass


@dataclass(slots=True)
class Literal(Node):
    value: object = None


@dataclass(slots=True)
class Reference(Node):
    """``A1``, ``$A$1:$B$2``, ``Sheet2!A1`` -- a block of cells."""

    text: str = ""
    sheet: str = ""


@dataclass(slots=True)
class NameNode(Node):
    """A defined name, or a constant such as TRUE that is not one."""

    name: str = ""
    sheet: str = ""


@dataclass(slots=True)
class Call(Node):
    name: str = ""
    args: list[Node] = field(default_factory=lambda: [])


@dataclass(slots=True)
class Unary(Node):
    op: str = ""
    operand: Node | None = None


@dataclass(slots=True)
class Binary(Node):
    op: str = ""
    left: Node | None = None
    right: Node | None = None


@dataclass(slots=True)
class ArrayLiteral(Node):
    rows: list[list[Node]] = field(default_factory=lambda: [])


# --- tokens -------------------------------------------------------------------------

#: A sheet name in front of a reference, quoted or not.
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_WHOLE_COLUMNS = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
_WHOLE_ROWS = r"\$?[0-9]{1,7}:\$?[0-9]{1,7}"

_TOKEN: Final = re.compile(
    rf"""
    (?P<ws>\s+)
  | (?P<text>"(?:[^"]|"")*")
  | (?P<error>(?:{_SHEET})?(?:\#N/A|\#NULL!|\#DIV/0!|\#VALUE!|\#REF!|\#NAME\?|\#NUM!|\#SPILL!|\#CALC!|\#GETTING_DATA))
  | (?P<ref>(?:{_SHEET})?(?:{_CELL}:{_CELL}|{_WHOLE_COLUMNS}|{_WHOLE_ROWS}|{_CELL})(?![A-Za-z0-9_.(]))
  | (?P<name>(?:{_SHEET})?[A-Za-z_\\À-￿][A-Za-z0-9_.À-￿]*)
  | (?P<number>(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?)
  | (?P<op><>|<=|>=|[=<>+\-*/^&%])
  | (?P<open>\()
  | (?P<close>\))
  | (?P<comma>,)
  | (?P<semicolon>;)
  | (?P<lbrace>\{{)
  | (?P<rbrace>\}})
    """,
    re.VERBOSE,
)


@dataclass(frozen=True, slots=True)
class Token:
    kind: str
    text: str
    at: int


def tokenize(source: str) -> list[Token]:
    """Every token of a formula, without the leading equals sign."""
    out: list[Token] = []
    position = 0
    while position < len(source):
        match = _TOKEN.match(source, position)
        if match is None:
            raise FormulaError(f"cannot read the formula at {source[position:position + 12]!r}")
        position = match.end()
        kind = match.lastgroup or ""
        if kind == "ws":
            continue
        out.append(Token(kind, match.group(0), match.start()))
    out.append(Token("eof", "", len(source)))
    return out


def split_sheet(text: str) -> tuple[str, str]:
    """``'My Sheet'!A1`` split into the sheet's name and the rest."""
    if "!" not in text:
        return "", text
    head, _, rest = text.rpartition("!")
    if head.startswith("'") and head.endswith("'"):
        head = head[1:-1].replace("''", "'")
    return head, rest


# --- the parser ---------------------------------------------------------------------

#: Loosest first; every one is left-associative.
_LEVELS: Final = (
    ("=", "<>", "<", ">", "<=", ">="),
    ("&",),
    ("+", "-"),
    ("*", "/"),
    ("^",),
)


class Parser:
    def __init__(self, tokens: list[Token], source: str) -> None:
        self.tokens = tokens
        self.source = source
        self.at = 0

    @property
    def token(self) -> Token:
        return self.tokens[self.at]

    def advance(self) -> Token:
        token = self.tokens[self.at]
        if token.kind != "eof":
            self.at += 1
        return token

    def expect(self, kind: str) -> Token:
        if self.token.kind != kind:
            raise FormulaError(
                f"expected {kind} but found {self.token.text!r} in {self.source!r}"
            )
        return self.advance()

    def parse(self) -> Node:
        node = self.expression()
        if self.token.kind != "eof":
            raise FormulaError(f"unexpected {self.token.text!r} in {self.source!r}")
        return node

    def expression(self, level: int = 0) -> Node:
        if level >= len(_LEVELS):
            return self.unary()
        left = self.expression(level + 1)
        while self.token.kind == "op" and self.token.text in _LEVELS[level]:
            op = self.advance().text
            right = self.expression(level + 1)
            left = Binary(op=op, left=left, right=right)
        return left

    def unary(self) -> Node:
        """Powers, which chain to the left and take a signed operand.

        Both halves are measured: ``2^3^2`` is 64 rather than 512, and
        ``-2^2`` is 4, because a minus sign binds tighter than the power
        rather than looser as it does in mathematics.
        """
        node = self.tight()
        while self.token.kind == "op" and self.token.text == "^":
            self.advance()
            node = Binary(op="^", left=node, right=self.tight())
        return node

    def tight(self) -> Node:
        """A signed operand with its percent signs, and no power."""
        if self.token.kind == "op" and self.token.text in ("-", "+"):
            op = self.advance().text
            operand = self.tight()
            return operand if op == "+" else Unary(op="-", operand=operand)
        node = self.primary()
        while self.token.kind == "op" and self.token.text == "%":
            self.advance()
            node = Unary(op="%", operand=node)
        return node

    def primary(self) -> Node:
        token = self.token
        if token.kind == "number":
            self.advance()
            return Literal(value=float(token.text))
        if token.kind == "text":
            self.advance()
            return Literal(value=token.text[1:-1].replace('""', '"'))
        if token.kind == "error":
            self.advance()
            text = "#" + token.text.rpartition("!#")[2] if "!#" in token.text else token.text
            return Literal(value=ERRORS.get(text, ExcelError(text)))
        if token.kind == "ref":
            self.advance()
            sheet, body = split_sheet(token.text)
            return Reference(text=body.replace("$", ""), sheet=sheet)
        if token.kind == "name":
            return self.name_or_call()
        if token.kind == "open":
            self.advance()
            inner = self.expression()
            self.expect("close")
            return inner
        if token.kind == "lbrace":
            return self.array()
        raise FormulaError(f"unexpected {token.text!r} in {self.source!r}")

    def at_kind(self, *kinds: str) -> bool:
        return self.tokens[self.at].kind in kinds

    def name_or_call(self) -> Node:
        token = self.advance()
        sheet, body = split_sheet(token.text)
        if self.at_kind("open"):
            self.advance()
            args: list[Node] = []
            if self.at_kind("close"):
                self.advance()
                return Call(name=body, args=args)
            while True:
                if self.at_kind("comma", "close"):
                    # An argument left out, as in OFFSET(A1,1,,2).
                    args.append(Literal(value=None))
                else:
                    args.append(self.expression())
                if self.at_kind("comma"):
                    self.advance()
                    continue
                break
            self.expect("close")
            return Call(name=body, args=args)
        upper = body.upper()
        if upper == "TRUE":
            return Literal(value=True)
        if upper == "FALSE":
            return Literal(value=False)
        return NameNode(name=body, sheet=sheet)

    def array(self) -> Node:
        self.expect("lbrace")
        rows: list[list[Node]] = [[]]
        while True:
            rows[-1].append(self.expression())
            if self.token.kind == "comma":
                self.advance()
                continue
            if self.token.kind == "semicolon":
                self.advance()
                rows.append([])
                continue
            break
        self.expect("rbrace")
        return ArrayLiteral(rows=rows)


def parse(formula: str) -> Node:
    """Parse a formula, with or without its leading equals sign."""
    body = formula.strip()
    if body.startswith("="):
        body = body[1:]
    if not body:
        raise FormulaError("an empty formula")
    return Parser(tokenize(body), body).parse()


# --- what a formula depends on --------------------------------------------------------

#: Functions whose answer can change without any cell changing, so a
#: formula holding one is recalculated every time.
VOLATILE: Final = frozenset(
    {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"}
)


def references(node: Node | None) -> list[Reference]:
    """Every reference a formula names, for working out what feeds it."""
    found: list[Reference] = []
    _walk(node, found)
    return found


def _walk(node: Node | None, found: list[Reference]) -> None:
    if node is None:
        return
    if isinstance(node, Reference):
        found.append(node)
    elif isinstance(node, Call):
        for argument in node.args:
            _walk(argument, found)
    elif isinstance(node, Binary):
        _walk(node.left, found)
        _walk(node.right, found)
    elif isinstance(node, Unary):
        _walk(node.operand, found)
    elif isinstance(node, ArrayLiteral):
        for row in node.rows:
            for item in row:
                _walk(item, found)


def names(node: Node | None) -> list[NameNode]:
    """Every defined name a formula reaches for."""
    found: list[NameNode] = []
    _walk_names(node, found)
    return found


def _walk_names(node: Node | None, found: list[NameNode]) -> None:
    if node is None:
        return
    if isinstance(node, NameNode):
        found.append(node)
    elif isinstance(node, Call):
        for argument in node.args:
            _walk_names(argument, found)
    elif isinstance(node, Binary):
        _walk_names(node.left, found)
        _walk_names(node.right, found)
    elif isinstance(node, Unary):
        _walk_names(node.operand, found)
    elif isinstance(node, ArrayLiteral):
        for row in node.rows:
            for item in row:
                _walk_names(item, found)


# --- moving a formula ------------------------------------------------------------------

_CORNER: Final = re.compile(r"^(\$?)([A-Za-z]{1,3})?(\$?)([0-9]{1,7})?$")


def shift_text(formula: str, down: int, across: int) -> str:
    """The same formula written for a cell ``down`` and ``across`` away.

    Writing a formula to a block, or copying one, moves every reference
    that is not held by a dollar sign.  The rewrite is done by splicing
    the original text so that spacing, case and everything else the
    author wrote survives.
    """
    if not down and not across:
        return formula
    body = formula[1:] if formula.startswith("=") else formula
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(body):
        if token.kind != "ref":
            continue
        sheet, reference = split_sheet(token.text)
        moved = shift_reference(reference, down, across)
        prefix = token.text[: len(token.text) - len(reference)] if sheet or "!" in token.text else ""
        pieces.append((token.at, token.at + len(token.text), prefix + moved))
    out = body
    for start, stop, replacement in reversed(pieces):
        out = out[:start] + replacement + out[stop:]
    return ("=" if formula.startswith("=") else "") + out


def shift_reference(text: str, down: int, across: int) -> str:
    """One reference moved, with anything behind a dollar sign left alone."""
    parts = text.split(":")
    moved = [_shift_corner(part, down, across) for part in parts]
    if any(part == "#REF!" for part in moved):
        return "#REF!"
    return ":".join(moved)


def _shift_corner(part: str, down: int, across: int) -> str:
    from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_letter, column_number

    match = _CORNER.match(part)
    if match is None:
        return part
    column_fixed, letters, row_fixed, digits = match.groups()
    out = ""
    if letters:
        if column_fixed:
            out += f"${letters}"
        else:
            number = column_number(letters) + across
            if not 1 <= number <= MAX_COLUMNS:
                return "#REF!"
            out += column_letter(number)
    if digits:
        if row_fixed:
            out += f"${digits}"
        else:
            number = int(digits) + down
            if not 1 <= number <= MAX_ROWS:
                return "#REF!"
            out += str(number)
    return out or part


def is_volatile(node: Node | None) -> bool:
    """Whether the formula has to be recalculated whatever else changed."""
    if node is None:
        return False
    if isinstance(node, Call):
        if node.name.upper() in VOLATILE:
            return True
        return any(is_volatile(argument) for argument in node.args)
    if isinstance(node, Binary):
        return is_volatile(node.left) or is_volatile(node.right)
    if isinstance(node, Unary):
        return is_volatile(node.operand)
    if isinstance(node, ArrayLiteral):
        return any(is_volatile(item) for row in node.rows for item in row)
    return False
