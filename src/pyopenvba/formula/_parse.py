"""Reading an Excel formula.

The grammar is not VBA's: ``&`` is the only concatenation, ``^`` binds
tighter than unary minus in one direction and looser in the other,
``%`` follows its operand, a reference is a value, and TRUE is a
constant rather than a keyword. References join tighter than anything:
``:`` spans two of them, a space intersects them, and a comma inside
brackets unites them; a function's bracket follows its name with no
space between. A table's name followed by square brackets is a
structured reference, Table1[Qty], read by :func:`read_structured`.

What is parsed here is the text after the leading ``=``.  A formula
that this cannot read raises :class:`FormulaError`, which is a compile
error rather than a cell error: Excel refuses to accept such a formula
at all rather than showing an error value.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import ROUND_DOWN, Decimal, localcontext
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
class Invoke(Node):
    """A call of what a call gives: a LAMBDA called where it is written, ``LAMBDA(x,x*2)(5)``."""

    target: Node | None = None
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
    #: Written in brackets, which keeps Excel from setting a last sum that cancels to zero.
    grouped: bool = field(default=False, compare=False)


@dataclass(slots=True)
class ArrayLiteral(Node):
    rows: list[list[Node]] = field(default_factory=lambda: [])


#: The special items of a structured reference, as Excel spells them.
ALL: Final = "#All"
DATA: Final = "#Data"
HEADERS: Final = "#Headers"
TOTALS: Final = "#Totals"
THIS_ROW: Final = "#This Row"
#: The order Excel writes items in, and the sets of them it takes (tests/fixtures/structured_references/).
ITEM_ORDER: Final = (ALL, HEADERS, DATA, TOTALS, THIS_ROW)
_ITEM_SETS: Final = frozenset({frozenset(), frozenset({ALL}), frozenset({DATA}), frozenset({HEADERS}),
                               frozenset({TOTALS}), frozenset({THIS_ROW}), frozenset({HEADERS, DATA}),
                               frozenset({DATA, TOTALS})})


@dataclass(slots=True)
class Structured(Node):
    """``Table1[Qty]``, ``Table1[[#Headers],[Qty]:[Price]]``, ``Table1[@Qty]``: part of a table, by name.

    ``items`` are the special items named, spelled and ordered as Excel
    has them, none meaning the data rows, and ``@`` is ``#This Row``.
    ``first`` and ``last`` are the columns, the same one twice for one
    column and None for all of them; ``span`` is set when they were
    written as a range, [[Qty]:[Qty]] too. ``spaced`` and
    ``comma_spaced`` are the spaces Excel keeps inside the brackets and
    after a comma. ``table`` is empty where the formula leaves it out,
    inside the table.
    """

    table: str = ""
    items: tuple[str, ...] = ()
    first: str | None = None
    last: str | None = None
    span: bool = False
    spaced: bool = False
    comma_spaced: bool = False
    sheet: str = ""
    #: Where the reference starts in the formula's text after the equals sign.
    at: int = field(default=0, compare=False)


# --- tokens -------------------------------------------------------------------------

#: A sheet name in front of a reference, quoted or not.
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_WHOLE_COLUMNS = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
_WHOLE_ROWS = r"\$?[0-9]{1,7}:\$?[0-9]{1,7}"
#: A column name inside brackets, where an apostrophe escapes the character after it.
_ESCAPED = r"(?:'[\s\S]|[^\[\]'])"
_BRACKETED = rf"\[{_ESCAPED}*\]"
#: A structured reference: a table's name, or none inside the table, and what stands in its brackets.
STRUCTURED: Final = (rf"(?:{_SHEET})?(?:[A-Za-z_\\À-￿][A-Za-z0-9_.À-￿]*)?"
                     rf"\[(?:{_ESCAPED}*|\s*@?\s*{_BRACKETED}(?:\s*[,:]\s*{_BRACKETED})*\s*)\]")

_TOKEN: Final = re.compile(
    rf"""
    (?P<ws>\s+)
  | (?P<text>"(?:[^"]|"")*")
  | (?P<structured>{STRUCTURED})
  | (?P<error>(?:{_SHEET})?(?i:\#N/A|\#NULL!|\#DIV/0!|\#VALUE!|\#REF!|\#NAME\?|\#NUM!|\#SPILL!|\#CALC!|\#GETTING_DATA))
  | (?P<ref>(?:{_SHEET})?(?:{_CELL}:{_CELL}|{_WHOLE_COLUMNS}|{_WHOLE_ROWS}|{_CELL})(?![A-Za-z0-9_.(]))
  | (?P<name>(?:{_SHEET})?(?:[A-Za-z_\\À-￿][A-Za-z0-9_.?\\À-￿]*|'(?:[^']|'')+'(?!!)))
  | (?P<number>(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?)
  | (?P<op><>|<=|>=|[=<>+\-*/^&%:@#])
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


def tokenize(source: str, *, spaces: bool = False) -> list[Token]:
    """Every token of a formula, without the leading equals sign; ``spaces`` keeps the runs of white space too."""
    out: list[Token] = []
    position = 0
    while position < len(source):
        match = _TOKEN.match(source, position)
        if match is None:
            raise FormulaError(f"cannot read the formula at {source[position:position + 12]!r}")
        position = match.end()
        kind = match.lastgroup or ""
        if kind == "ws" and not spaces:
            continue
        if kind == "ref" and not _on_the_sheet(match.group(0)):
            # XFE1 and A1048577 are past the sheet's edge: names, as Excel reads them
            # (tests/fixtures/formula_notation.json).
            kind = "name"
        out.append(Token(kind, match.group(0), match.start()))
    out.append(Token("eof", "", len(source)))
    return _bind_cells(out)


#: The functions that bind names, as a formula or a file spells them.
_BINDERS: Final = {"LET": "LET", "_XLFN.LET": "LET", "LAMBDA": "LAMBDA", "_XLFN.LAMBDA": "LAMBDA"}
#: A cell a LET or a LAMBDA can bind as a name: one, with no $ and no sheet.
_CELL_NAME: Final = re.compile(r"[A-Za-z]{1,3}[0-9]{1,7}")


@dataclass
class _Binding:
    """A bracket open while a formula is read, and what a LET or a LAMBDA opening it binds."""

    binder: str = ""
    bound: set[str] = field(default_factory=lambda: set())
    argument: int = 0
    starting: bool = True


def _bind_cells(tokens: list[Token]) -> list[Token]:
    """``tokens`` with each cell a LET or a LAMBDA binds, and each standing for it in its scope, made a name:
    Excel takes LET(x1,5,x1) to be 5, x1 a name there, and keeps it as written (tests/fixtures/bound_names.json)."""
    if not any(token.kind == "name" and token.text.upper() in _BINDERS for token in tokens):
        return tokens
    out = list(tokens)
    solid = [index for index, token in enumerate(tokens) if token.kind != "ws"]
    open_: list[_Binding] = []
    for place, index in enumerate(solid):
        token = tokens[index]
        following = tokens[solid[place + 1]] if place + 1 < len(solid) else token
        if token.kind == "ref" and _CELL_NAME.fullmatch(token.text) and _on_the_sheet(token.text):
            key = token.text.upper()
            inner = open_[-1] if open_ else None
            if inner is not None and inner.binder and inner.starting and following.kind == "comma" \
                    and (inner.binder == "LAMBDA" or inner.argument % 2 == 0):
                inner.bound.add(key)
                out[index] = Token("name", token.text, token.at)
            elif any(key in bracket.bound for bracket in open_):
                out[index] = Token("name", token.text, token.at)
        if open_ and token.kind != "eof":
            open_[-1].starting = False
        if token.kind in ("open", "lbrace"):
            previous = tokens[solid[place - 1]] if place else None
            called = previous is not None and previous.kind == "name" and token.kind == "open" \
                and token.at == previous.at + len(previous.text)
            open_.append(_Binding(binder=_BINDERS.get(previous.text.upper(), "") if called and previous else ""))
        elif token.kind in ("close", "rbrace") and open_:
            open_.pop()
        elif token.kind == "comma" and open_:
            open_[-1].argument += 1
            open_[-1].starting = True
    return out


_ON_THE_SHEET = re.compile(r"\$?([A-Za-z]{1,3})?\$?([0-9]{1,7})?")


def _on_the_sheet(text: str) -> bool:
    """Whether each corner of a reference is on the sheet: its column XFD at most, its row 1048576 at most."""
    from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_number

    for corner in split_sheet(text)[1].split(":"):
        found = _ON_THE_SHEET.fullmatch(corner)
        if found is None:
            return False
        letters, digits = found.groups()
        if letters and column_number(letters) > MAX_COLUMNS or digits and not 1 <= int(digits) <= MAX_ROWS:
            return False
    return True


def literal(text: str) -> float:
    """A number written in a formula, as Excel keeps it: to fifteen significant digits, the rest cut off.

    Excel drops the digits past the fifteenth rather than rounding them, so
    =123456789012345678 holds 123456789012345000.
    """
    number = Decimal(text)
    if not number:
        return 0.0
    with localcontext() as context:
        context.prec = 1000
        return float(number.quantize(Decimal(1).scaleb(number.adjusted() - 14), rounding=ROUND_DOWN))


def split_sheet(text: str) -> tuple[str, str]:
    """``'My Sheet'!A1`` split into the sheet's name and the rest."""
    if "!" not in text:
        return "", text
    head, _, rest = text.rpartition("!")
    if head.startswith("'") and head.endswith("'"):
        head = head[1:-1].replace("''", "'")
    return head, rest


def read_structured(text: str, *, file: bool = False, at: int = 0) -> Structured:
    """A structured reference as a formula spells it, or as a file does with ``file``.

    On screen an @ in a column's name is escaped, '@home, since it would
    otherwise mean this row; a file spells this row [#This Row] and has no
    need to (tests/fixtures/structured_references/). A spelling Excel does
    not take raises :class:`FormulaError`.
    """
    opening = text.index("[")
    sheet, table = split_sheet(text[:opening])
    body = text[opening + 1:-1]
    inner = body.strip()
    node = Structured(table=table, sheet=sheet, at=at)
    if not inner:
        return node
    if inner.startswith("@") and not file:
        rest = inner[1:].strip()
        node.items = (THIS_ROW,)
        if rest.startswith("["):
            node.first, node.last, node.span = _columns(rest, file)
        elif rest:
            node.first = node.last = _unescaped(rest, file)
        node.spaced = body != inner
        return node
    if inner.startswith("["):
        items: list[str] = []
        parts, node.comma_spaced = _parts(inner)
        for part in parts:
            if part.startswith("[#"):
                items.append(_item(part[1:-1]))
            elif node.first is None:
                node.first, node.last, node.span = _columns(part, file)
            else:
                raise FormulaError(f"two sets of columns in {text!r}")
        node.items = _checked_items(items, text)
        node.spaced = body != inner
        return node
    if inner.startswith("#"):
        node.items = _checked_items([_item(inner)], text)
        return node
    # One column, its name everything in the brackets, spaces too: a space in front also keeps [ [ ab] ].
    node.first = node.last = _unescaped(body, file)
    node.spaced = body.startswith(" ")
    return node


def _item(text: str) -> str:
    wanted = " ".join(text.split()).lower()
    found = next((item for item in ITEM_ORDER if item.lower() == wanted), None)
    if found is None:
        raise FormulaError(f"{text!r} is not a special item of a table")
    return found


def _checked_items(items: list[str], text: str) -> tuple[str, ...]:
    if len(set(items)) != len(items) or frozenset(items) not in _ITEM_SETS:
        raise FormulaError(f"{text!r} names special items Excel does not take together")
    return tuple(item for item in ITEM_ORDER if item in items)


def _parts(inner: str) -> tuple[list[str], bool]:
    """``[#All], [Qty]:[Price]`` split at the commas between the brackets, and whether a comma had a space by it."""
    parts: list[str] = []
    spaced = False
    depth = 0
    start = 0
    index = 0
    while index < len(inner):
        char = inner[index]
        if char == "'" and depth:
            index += 2
            continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
        elif char == "," and not depth:
            parts.append(inner[start:index].strip())
            spaced = spaced or inner[index - 1:index].isspace() or inner[index + 1:index + 2].isspace()
            start = index + 1
        index += 1
    parts.append(inner[start:].strip())
    if any(not part for part in parts):
        raise FormulaError(f"an empty part in {inner!r}")
    return parts, spaced


def _columns(part: str, file: bool) -> tuple[str, str, bool]:
    """``[Qty]`` or ``[Qty]:[Price]``: the first and last column, and whether they were written as a range."""
    pieces = re.fullmatch(rf"({_BRACKETED})(?:\s*:\s*({_BRACKETED}))?", part)
    if pieces is None:
        raise FormulaError(f"cannot read the columns {part!r}")
    first = _unescaped(pieces.group(1)[1:-1], file)
    second = pieces.group(2)
    return first, first if second is None else _unescaped(second[1:-1], file), second is not None


def _unescaped(name: str, file: bool) -> str:
    """A column's name with its escapes undone: '# is #, '' is '."""
    out: list[str] = []
    index = 0
    while index < len(name):
        char = name[index]
        if char == "'" and index + 1 < len(name):
            out.append(name[index + 1])
            index += 2
            continue
        if char in "[]#'" or (char == "@" and not file):
            raise FormulaError(f"{char!r} unescaped in the column {name!r}")
        out.append(char)
        index += 1
    if not out:
        raise FormulaError("a column with no name")
    return "".join(out)


# --- the parser ---------------------------------------------------------------------

#: The operators that join references: a range between two, their intersection, and their union.
REFERENCE_OPS: Final = frozenset({":", " ", ","})

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
        #: Where each node's text starts in the source, by the node's id.
        self.starts: dict[int, int] = {}

    def placed(self, node: Node, start: int) -> Node:
        """``node``, noted as starting at ``start`` in the source."""
        self.starts[id(node)] = start
        return node

    def start(self, node: Node) -> int:
        return self.starts.get(id(node), -1)

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
            left = self.placed(Binary(op=op, left=left, right=right), self.start(left))
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
            node = self.placed(Binary(op="^", left=node, right=self.tight()), self.start(node))
        return node

    def tight(self) -> Node:
        """A signed operand with its percent signs, and no power; an @ binds as a sign does, over a range or an
        intersection whole: @A1:A3 A2:A3."""
        if self.token.kind == "op" and self.token.text in ("-", "+", "@"):
            start = self.token.at
            op = self.advance().text
            # A + is kept: it reads cells as values, so COUNTIF(+A1:A3,1) is refused (tests/fixtures/formula/).
            return self.placed(Unary(op=op, operand=self.tight()), start)
        node = self.operand()
        while self.token.kind == "op" and self.token.text == "%":
            self.advance()
            node = self.placed(Unary(op="%", operand=node), self.start(node))
        return node

    def operand(self) -> Node:
        """A primary with the reference operators that bind tightest: a range (``:``) and an intersection (a space).
        A ``#`` after a cell is what spilled from it, A1#."""
        node = self.spilled(self.primary())
        while True:
            if self.token.kind == "op" and self.token.text == ":":
                self.advance()
                joined = Binary(op=":", left=self.referable(node), right=self.referable(self.primary()))
                node = self.placed(joined, self.start(node))
            elif self.token.kind in ("ref", "structured", "name", "open") and self.spaced():
                joined = Binary(op=" ", left=self.referable(node), right=self.referable(self.primary()))
                node = self.placed(joined, self.start(node))
            else:
                return node

    def spilled(self, node: Node) -> Node:
        """``node`` and a ``#`` after it, what spilled from a cell: A1#."""
        while self.token.kind == "op" and self.token.text == "#":
            self.advance()
            node = self.placed(Unary(op="#", operand=node), self.start(node))
        return node

    def spaced(self) -> bool:
        """Whether white space stands between the last token read and the next."""
        if not self.at:
            return False
        before = self.tokens[self.at - 1]
        return self.token.at > before.at + len(before.text)

    def referable(self, node: Node) -> Node:
        """A reference operator's operand, which has to be a reference, a name or a call that might give one."""
        if isinstance(node, (Reference, Structured, NameNode, Call)) \
                or (isinstance(node, Binary) and node.op in REFERENCE_OPS) or isinstance(node, Unary) and node.op == "#":
            return node
        raise FormulaError(f"a reference operator needs references in {self.source!r}")

    def primary(self) -> Node:
        token = self.token
        if token.kind == "open":
            # The node inside brackets starts inside them, where Formula2 puts its @: (@A1:A3).
            return self.bracketed()
        return self.placed(self.unplaced(), token.at)

    def unplaced(self) -> Node:
        token = self.token
        if token.kind == "number":
            self.advance()
            return Literal(value=literal(token.text))
        if token.kind == "text":
            self.advance()
            return Literal(value=token.text[1:-1].replace('""', '"'))
        if token.kind == "error":
            self.advance()
            text = ("#" + token.text.rpartition("!#")[2] if "!#" in token.text else token.text).upper()
            return Literal(value=ERRORS.get(text, ExcelError(text)))
        if token.kind == "ref":
            self.advance()
            sheet, body = split_sheet(token.text)
            return Reference(text=body.replace("$", ""), sheet=sheet)
        if token.kind == "structured":
            self.advance()
            return read_structured(token.text, at=token.at)
        if token.kind == "name":
            node = self.name_or_call()
            # Brackets straight after a call call what it gives: LAMBDA(x,x*2)(5) is 10 (tests/fixtures/formula/).
            while isinstance(node, (Call, Invoke)) and self.at_kind("open") and not self.spaced():
                self.advance()
                node = self.placed(Invoke(target=node, args=self.arguments()), token.at)
            return node
        if token.kind == "lbrace":
            return self.array()
        raise FormulaError(f"unexpected {token.text!r} in {self.source!r}")

    def bracketed(self) -> Node:
        """What stands in brackets, the node inside them keeping its own start."""
        self.advance()
        inner = self.expression()
        while self.token.kind == "comma":
            # Inside brackets a comma joins references into a union.
            self.advance()
            joined = Binary(op=",", left=self.referable(inner), right=self.referable(self.expression()))
            inner = self.placed(joined, self.start(inner))
        self.expect("close")
        if isinstance(inner, Binary):
            inner.grouped = True
        return inner

    def at_kind(self, *kinds: str) -> bool:
        return self.tokens[self.at].kind in kinds

    def name_or_call(self) -> Node:
        token = self.advance()
        sheet, body = split_sheet(token.text)
        # A call's bracket follows its name straight away; after a space it opens an intersection.
        if self.at_kind("open") and not self.spaced():
            self.advance()
            return Call(name=body, args=self.arguments())
        upper = body.upper()
        if upper == "TRUE":
            return Literal(value=True)
        if upper == "FALSE":
            return Literal(value=False)
        return NameNode(name=body, sheet=sheet)

    def arguments(self) -> list[Node]:
        """A call's arguments, its opening bracket read, up to and past its closing one."""
        args: list[Node] = []
        if self.at_kind("close"):
            self.advance()
            return args
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
        return args

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
    return parse_placed(formula)[0]


def parse_placed(formula: str) -> tuple[Node, dict[int, int]]:
    """A formula parsed, with where each node's text starts after the = by the node's id."""
    body = formula.strip()
    if body.startswith("="):
        body = body[1:]
    if not body:
        raise FormulaError("an empty formula")
    parser = Parser(tokenize(body), body)
    return parser.parse(), parser.starts


# --- what a formula depends on --------------------------------------------------------

#: Functions whose answer can change without any cell changing, so a
#: formula holding one is recalculated every time. SUBTOTAL changes as
#: rows are hidden, filtered or shown, which Excel recalculates it for.
VOLATILE: Final = frozenset(
    {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO", "SUBTOTAL"}
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
    elif isinstance(node, Invoke):
        for argument in [node.target, *node.args]:
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


def structured_references(node: Node | None) -> list[Structured]:
    """Every structured reference a formula names."""
    found: list[Structured] = []
    _walk_structured(node, found)
    return found


def _walk_structured(node: Node | None, found: list[Structured]) -> None:
    if isinstance(node, Structured):
        found.append(node)
    elif isinstance(node, Call):
        for argument in node.args:
            _walk_structured(argument, found)
    elif isinstance(node, Invoke):
        for argument in [node.target, *node.args]:
            _walk_structured(argument, found)
    elif isinstance(node, Binary):
        _walk_structured(node.left, found)
        _walk_structured(node.right, found)
    elif isinstance(node, Unary):
        _walk_structured(node.operand, found)


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
    elif isinstance(node, Invoke):
        for argument in [node.target, *node.args]:
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


def transpose_text(formula: str, source: tuple[int, int], target: tuple[int, int]) -> str:
    """A formula pasted transposed, from the cell at ``source`` to the one at ``target``.

    Measured in live Excel: a reference relative in both its row and its
    column keeps its distance from the formula with the rows and columns
    swapped, so =C5 in A1 pasted transposed into D1 is =H3. A reference
    with a dollar sign anywhere stays as it is.
    """
    from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_letter, column_number

    body = formula[1:] if formula.startswith("=") else formula
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(body):
        if token.kind != "ref":
            continue
        _, reference = split_sheet(token.text)
        corners: list[str] = []
        for part in reference.split(":"):
            found = _CORNER.match(part)
            column_fixed, letters, row_fixed, digits = found.groups() if found else ("$", None, "$", None)
            if column_fixed or row_fixed or not letters or not digits:
                corners.append(part)
                continue
            row = target[0] + column_number(letters) - source[1]
            column = target[1] + int(digits) - source[0]
            if not (1 <= row <= MAX_ROWS and 1 <= column <= MAX_COLUMNS):
                corners = ["#REF!"]
                break
            corners.append(f"{column_letter(column)}{row}")
        prefix = token.text[: len(token.text) - len(reference)]
        pieces.append((token.at, token.at + len(token.text), prefix + ":".join(corners)))
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
    if isinstance(node, Invoke):
        return any(is_volatile(argument) for argument in [node.target, *node.args])
    if isinstance(node, Binary):
        return is_volatile(node.left) or is_volatile(node.right)
    if isinstance(node, Unary):
        return is_volatile(node.operand)
    if isinstance(node, ArrayLiteral):
        return any(is_volatile(item) for row in node.rows for item in row)
    return False
