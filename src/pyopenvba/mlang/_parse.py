"""Reading the M language.

M is a functional language: a query is one expression, usually a ``let``
whose steps each name the one before.  There are no statements, every
step is a binding, and a name may be written ``#"like this"`` when it
has spaces in it.

The grammar here is the one a query uses.  Sections, ``#shared``, type
ascription beyond naming a type, and the metadata operator are read and
carried rather than acted on, so a formula that has one still parses.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Final

from pyopenvba.exceptions import PyOpenVBAError


class MSyntaxError(PyOpenVBAError):
    """Raised for M that does not parse."""


# --- the tree ---------------------------------------------------------------------


@dataclass(slots=True)
class Node:
    pass


@dataclass(slots=True)
class Literal(Node):
    value: object = None


@dataclass(slots=True)
class Name(Node):
    name: str = ""


@dataclass(slots=True)
class Let(Node):
    bindings: list[tuple[str, Node]] = field(default_factory=lambda: [])
    body: Node | None = None


@dataclass(slots=True)
class If(Node):
    condition: Node | None = None
    then: Node | None = None
    otherwise: Node | None = None


@dataclass(slots=True)
class RecordLiteral(Node):
    fields: list[tuple[str, Node]] = field(default_factory=lambda: [])


@dataclass(slots=True)
class ListLiteral(Node):
    items: list[Node] = field(default_factory=lambda: [])


@dataclass(slots=True)
class FunctionLiteral(Node):
    parameters: list[tuple[str, bool, str]] = field(default_factory=lambda: [])
    body: Node | None = None
    returns: str = ""


@dataclass(slots=True)
class Each(Node):
    body: Node | None = None


@dataclass(slots=True)
class Call(Node):
    target: Node | None = None
    args: list[Node] = field(default_factory=lambda: [])


@dataclass(slots=True)
class FieldAccess(Node):
    target: Node | None = None
    name: str = ""
    optional: bool = False


@dataclass(slots=True)
class ItemAccess(Node):
    target: Node | None = None
    index: Node | None = None
    optional: bool = False


@dataclass(slots=True)
class Projection(Node):
    """``table[[A], [B]]``: the named fields only."""

    target: Node | None = None
    names: list[str] = field(default_factory=lambda: [])
    optional: bool = False


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
class TryExpr(Node):
    body: Node | None = None
    otherwise: Node | None = None


@dataclass(slots=True)
class ErrorExpr(Node):
    body: Node | None = None


@dataclass(slots=True)
class TypeExpr(Node):
    name: str = ""
    detail: Node | None = None


@dataclass(slots=True)
class NotImplemented_(Node):
    """``...``, which M raises on if it is ever reached."""


# --- tokens -------------------------------------------------------------------------

_KEYWORDS: Final = frozenset(
    """
    and as each else error false if in is let meta not null or otherwise section shared then
    true try type
    """.split()
)

_TOKEN: Final = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<comment>//[^\r\n]*|/\*.*?\*/)
  | (?P<text>"(?:[^"]|"")*")
  | (?P<quoted>\#"(?:[^"]|"")*")
  | (?P<hashword>\#[a-z]+)
  | (?P<number>0[xX][0-9A-Fa-f]+|(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)(?:[eE][-+]?[0-9]+)?)
  | (?P<name>[A-Za-z_][A-Za-z0-9_.]*)
  | (?P<op><>|<=|>=|=>|\?\?|\.\.\.|\.\.|[=<>+\-*/&?@])
  | (?P<open>\()
  | (?P<close>\))
  | (?P<lbrace>\{)
  | (?P<rbrace>\})
  | (?P<lbracket>\[)
  | (?P<rbracket>\])
  | (?P<comma>,)
  | (?P<semicolon>;)
    """,
    re.VERBOSE | re.DOTALL,
)

_ESCAPES: Final = {"lf": "\n", "cr": "\r", "tab": "\t", "#": "#", "(": "("}


@dataclass(frozen=True, slots=True)
class Token:
    kind: str
    text: str
    at: int


def tokenize(source: str) -> list[Token]:
    out: list[Token] = []
    position = 0
    while position < len(source):
        match = _TOKEN.match(source, position)
        if match is None:
            raise MSyntaxError(f"cannot read the formula at {source[position:position + 16]!r}")
        position = match.end()
        kind = match.lastgroup or ""
        if kind in ("ws", "comment"):
            continue
        out.append(Token(kind, match.group(0), match.start()))
    out.append(Token("eof", "", len(source)))
    return out


def unquote(text: str) -> str:
    """``#"a b"`` as the name it stands for."""
    body = text[2:-1] if text.startswith('#"') else text[1:-1]
    return body.replace('""', '"')


def text_value(raw: str) -> str:
    """A text literal's characters, with M's escapes read.

    ``#(lf)`` is a newline, ``#(tab)`` a tab, ``#(#)`` a hash, and a
    four-digit code is the character it names.
    """
    body = raw[1:-1].replace('""', '"')
    out: list[str] = []
    index = 0
    while index < len(body):
        if body.startswith("#(", index):
            close = body.find(")", index)
            if close > 0:
                for piece in body[index + 2 : close].split(","):
                    if piece in _ESCAPES:
                        out.append(_ESCAPES[piece])
                    elif re.fullmatch(r"[0-9A-Fa-f]{4,8}", piece):
                        out.append(chr(int(piece, 16)))
                    else:
                        out.append(f"#({piece})")
                index = close + 1
                continue
        out.append(body[index])
        index += 1
    return "".join(out)


# --- the parser ----------------------------------------------------------------------

#: Loosest first.
_LEVELS: Final = (
    ("or",),
    ("and",),
    ("=", "<>", "<", ">", "<=", ">="),
    ("&", "+", "-"),
    ("*", "/"),
)


class Parser:
    def __init__(self, tokens: list[Token], source: str) -> None:
        self.tokens = tokens
        self.source = source
        self.at = 0

    # -- helpers

    @property
    def token(self) -> Token:
        return self.tokens[self.at]

    def kind(self) -> str:
        return self.tokens[self.at].kind

    def text(self) -> str:
        return self.tokens[self.at].text

    def advance(self) -> Token:
        token = self.tokens[self.at]
        if token.kind != "eof":
            self.at += 1
        return token

    def at_word(self, *words: str) -> bool:
        token = self.tokens[self.at]
        return token.kind == "name" and token.text in words

    def accept_word(self, *words: str) -> bool:
        if self.at_word(*words):
            self.advance()
            return True
        return False

    def at_kind(self, *kinds: str) -> bool:
        return self.tokens[self.at].kind in kinds

    def accept_kind(self, kind: str) -> bool:
        if self.at_kind(kind):
            self.advance()
            return True
        return False

    def expect(self, kind: str) -> Token:
        if not self.at_kind(kind):
            raise MSyntaxError(f"expected {kind} but found {self.text()!r} in {self.source[:60]!r}")
        return self.advance()

    def expect_word(self, word: str) -> None:
        if not self.accept_word(word):
            raise MSyntaxError(f"expected {word!r} but found {self.text()!r}")

    # -- expressions

    def parse(self) -> Node:
        node = self.expression()
        if not self.at_kind("eof"):
            raise MSyntaxError(f"unexpected {self.text()!r} after the expression")
        return node

    def expression(self) -> Node:
        if self.at_word("let"):
            return self.let()
        if self.at_word("if"):
            return self.conditional()
        if self.at_word("each"):
            self.advance()
            return Each(body=self.expression())
        if self.at_word("try"):
            return self.try_expression()
        if self.at_word("error"):
            self.advance()
            return ErrorExpr(body=self.expression())
        if self.at_word("type"):
            return self.type_expression()
        if self.at_kind("open") and self._is_function_head():
            return self.function()
        return self.operators()

    def let(self) -> Node:
        self.expect_word("let")
        bindings: list[tuple[str, Node]] = []
        while True:
            name = self.identifier()
            if not self.at_kind("op") or self.text() != "=":
                raise MSyntaxError(f"a let step needs an '=' after {name!r}")
            self.advance()
            bindings.append((name, self.expression()))
            if self.accept_kind("comma"):
                continue
            break
        self.expect_word("in")
        return Let(bindings=bindings, body=self.expression())

    def conditional(self) -> Node:
        self.expect_word("if")
        condition = self.expression()
        self.expect_word("then")
        then = self.expression()
        self.expect_word("else")
        return If(condition=condition, then=then, otherwise=self.expression())

    def try_expression(self) -> Node:
        self.expect_word("try")
        body = self.expression()
        if self.accept_word("otherwise"):
            return TryExpr(body=body, otherwise=self.expression())
        return TryExpr(body=body)

    def type_expression(self) -> Node:
        self.expect_word("type")
        if self.at_kind("lbracket", "lbrace"):
            return TypeExpr(name="record", detail=self.primary())
        name = self.identifier() if self.at_kind("name", "quoted") else ""
        if name == "table" and self.at_kind("lbracket"):
            return TypeExpr(name="table", detail=self.primary())
        if name == "nullable" and self.at_kind("name"):
            return TypeExpr(name=f"nullable {self.identifier()}")
        return TypeExpr(name=name or "any")

    def _is_function_head(self) -> bool:
        """Whether the bracket opens a parameter list rather than a group."""
        depth = 0
        index = self.at
        while index < len(self.tokens):
            token = self.tokens[index]
            if token.kind == "open":
                depth += 1
            elif token.kind == "close":
                depth -= 1
                if depth == 0:
                    after = self.tokens[index + 1] if index + 1 < len(self.tokens) else None
                    return after is not None and after.kind == "op" and after.text == "=>"
            index += 1
        return False

    def function(self) -> Node:
        self.expect("open")
        parameters: list[tuple[str, bool, str]] = []
        while not self.at_kind("close"):
            optional = self.accept_word("optional")
            name = self.identifier()
            declared = ""
            if self.accept_word("as"):
                declared = self._type_name()
            parameters.append((name, optional, declared))
            if not self.accept_kind("comma"):
                break
        self.expect("close")
        returns = ""
        if not self.at_kind("op") or self.text() != "=>":
            if self.accept_word("as"):
                returns = self._type_name()
        if not (self.at_kind("op") and self.text() == "=>"):
            raise MSyntaxError("a function needs '=>' after its parameters")
        self.advance()
        return FunctionLiteral(parameters=parameters, body=self.expression(), returns=returns)

    def _type_name(self) -> str:
        if self.accept_word("type"):
            pass
        if self.accept_word("nullable"):
            return f"nullable {self.identifier()}" if self.at_kind("name") else "nullable any"
        if self.at_kind("name", "quoted"):
            return self.identifier()
        return "any"

    def operators(self, level: int = 0) -> Node:
        if level >= len(_LEVELS):
            return self.unary()
        left = self.operators(level + 1)
        while True:
            token = self.tokens[self.at]
            op = ""
            if token.kind == "op" and token.text in _LEVELS[level]:
                op = token.text
            elif token.kind == "name" and token.text in _LEVELS[level]:
                op = token.text
            if not op:
                break
            self.advance()
            left = Binary(op=op, left=left, right=self.operators(level + 1))
        if level == 0:
            left = self._tail(left)
        return left

    def _tail(self, left: Node) -> Node:
        """The operators that sit at the very end: as, is, meta and ??."""
        while True:
            if self.accept_word("as"):
                left = Binary(op="as", left=left, right=TypeExpr(name=self._type_name()))
                continue
            if self.accept_word("is"):
                left = Binary(op="is", left=left, right=TypeExpr(name=self._type_name()))
                continue
            if self.accept_word("meta"):
                left = Binary(op="meta", left=left, right=self.operators(1))
                continue
            if self.at_kind("op") and self.text() == "??":
                self.advance()
                left = Binary(op="??", left=left, right=self.operators(1))
                continue
            return left

    def unary(self) -> Node:
        if self.at_kind("op") and self.text() in ("-", "+"):
            op = self.advance().text
            operand = self.unary()
            return operand if op == "+" else Unary(op="-", operand=operand)
        if self.at_word("not"):
            self.advance()
            return Unary(op="not", operand=self.unary())
        return self.postfix(self.primary())

    def postfix(self, node: Node) -> Node:
        while True:
            if self.at_kind("open"):
                self.advance()
                args: list[Node] = []
                while not self.at_kind("close"):
                    args.append(self.expression())
                    if not self.accept_kind("comma"):
                        break
                self.expect("close")
                node = Call(target=node, args=args)
                continue
            if self.at_kind("lbrace"):
                self.advance()
                index = self.expression()
                self.expect("rbrace")
                optional = self._optional_marker()
                node = ItemAccess(target=node, index=index, optional=optional)
                continue
            if self.at_kind("lbracket"):
                node = self._bracket(node)
                continue
            return node

    def _optional_marker(self) -> bool:
        if self.at_kind("op") and self.text() == "?":
            self.advance()
            return True
        return False

    def _bracket(self, node: Node | None) -> Node:
        self.expect("lbracket")
        if self.at_kind("lbracket"):
            names: list[str] = []
            while self.at_kind("lbracket"):
                self.advance()
                names.append(self.identifier())
                self.expect("rbracket")
                if not self.accept_kind("comma"):
                    break
            self.expect("rbracket")
            return Projection(target=node, names=names, optional=self._optional_marker())
        name = self.identifier()
        self.expect("rbracket")
        return FieldAccess(target=node, name=name, optional=self._optional_marker())

    def identifier(self) -> str:
        token = self.tokens[self.at]
        if token.kind == "quoted":
            self.advance()
            return unquote(token.text)
        if token.kind == "name":
            self.advance()
            return token.text
        raise MSyntaxError(f"expected a name but found {token.text!r}")

    def primary(self) -> Node:
        token = self.tokens[self.at]
        kind = token.kind
        if kind == "number":
            self.advance()
            text = token.text
            value = float(int(text, 16)) if text[:2].lower() == "0x" else float(text)
            return Literal(value=int(value) if value.is_integer() else value)
        if kind == "text":
            self.advance()
            return Literal(value=text_value(token.text))
        if kind == "quoted":
            self.advance()
            return Name(name=unquote(token.text))
        if kind == "hashword":
            return self.hash_word()
        if kind == "lbrace":
            self.advance()
            items: list[Node] = []
            while not self.at_kind("rbrace"):
                item = self.expression()
                if self.at_kind("op") and self.text() == "..":
                    # {1..3} is the list of the numbers between them.
                    self.advance()
                    item = Call(target=Name(name="#range"), args=[item, self.expression()])
                items.append(item)
                if not self.accept_kind("comma"):
                    break
            self.expect("rbrace")
            return ListLiteral(items=items)
        if kind == "lbracket":
            if self._is_record_literal():
                return self.record_literal()
            # A bracket that is not a record is a field of the row: M
            # reads [Amount] as _[Amount], which is what makes
            # each [Amount] > 100 work.
            return self._bracket(None)
        if kind == "open":
            self.advance()
            inner = self.expression()
            self.expect("close")
            return inner
        if kind == "op" and token.text == "...":
            self.advance()
            return NotImplemented_()
        if kind == "op" and token.text == "@":
            self.advance()
            return Name(name=self.identifier())
        if kind == "name":
            if token.text == "true":
                self.advance()
                return Literal(value=True)
            if token.text == "false":
                self.advance()
                return Literal(value=False)
            if token.text == "null":
                self.advance()
                return Literal(value=None)
            if token.text in _KEYWORDS and token.text not in ("not", "type", "each"):
                raise MSyntaxError(f"unexpected keyword {token.text!r}")
            self.advance()
            return Name(name=token.text)
        raise MSyntaxError(f"unexpected {token.text!r} in {self.source[:60]!r}")

    def _is_record_literal(self) -> bool:
        """Whether the bracket opens a record rather than a field name.

        ``[a = 1]`` is a record and ``[a]`` is the field a of the row,
        and what separates them is the equals sign after the name.
        """
        after = self.tokens[self.at + 1] if self.at + 1 < len(self.tokens) else None
        if after is None or after.kind == "rbracket":
            return True
        if after.kind == "lbracket":
            return False
        following = self.tokens[self.at + 2] if self.at + 2 < len(self.tokens) else None
        return following is not None and following.kind == "op" and following.text == "="

    def record_literal(self) -> Node:
        self.expect("lbracket")
        fields: list[tuple[str, Node]] = []
        while not self.at_kind("rbracket"):
            name = self.identifier()
            if not (self.at_kind("op") and self.text() == "="):
                raise MSyntaxError(f"a record field needs an '=' after {name!r}")
            self.advance()
            fields.append((name, self.expression()))
            if not self.accept_kind("comma"):
                break
        self.expect("rbracket")
        return RecordLiteral(fields=fields)

    def hash_word(self) -> Node:
        token = self.advance()
        word = token.text[1:]
        if word in ("date", "datetime", "datetimezone", "time", "duration", "table", "binary"):
            name = Name(name=f"#{word}")
            return self.postfix(name) if self.at_kind("open") else name
        if word == "infinity":
            return Literal(value=float("inf"))
        if word == "nan":
            return Literal(value=float("nan"))
        if word in ("sections", "shared"):
            return Name(name=f"#{word}")
        raise MSyntaxError(f"unexpected #{word}")


def parse(source: str) -> Node:
    """Parse one M expression, which is what a query's formula is."""
    body = source.strip()
    if not body:
        raise MSyntaxError("an empty formula")
    return Parser(tokenize(body), body).parse()
