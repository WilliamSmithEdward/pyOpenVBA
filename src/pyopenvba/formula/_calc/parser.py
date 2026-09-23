"""Building a formula's tree, with Excel's precedence.

Excel's operators bind in an order of their own, tightest first:

1. the reference operators, ``:`` then a space then ``,``
2. negation, so ``-2^2`` is ``(-2)^2``, which is 4
3. ``%``
4. ``^``, which groups from the left: ``2^3^2`` is 64
5. ``*`` and ``/``
6. ``+`` and ``-``
7. ``&``
8. the comparisons

Two of those need care. A space is the intersection operator only between
two operands; anywhere else it is whitespace. And a comma is the union
operator only where it cannot separate arguments: inside parentheses of its
own, or at the top of a defined name's formula.
"""

from __future__ import annotations

from pyopenvba.formula._calc.lexer import FormulaSyntaxError, Kind, Token, tokenize
from pyopenvba.formula._calc.nodes import (
    BINDING,
    PERCENT_BINDING,
    PREFIX_BINDING,
    SPILL_BINDING,
    AreaReference,
    ArrayItem,
    ArrayLiteral,
    Binary,
    Call,
    CellReference,
    ErrorLiteral,
    Invoke,
    Logical,
    Missing,
    Node,
    Number,
    Paren,
    Postfix,
    Text,
    Unary,
)
from pyopenvba.formula._calc.cells import CellError

#: What ``@`` applies to: a range, so it binds below ``:``, and nothing
#: looser, so ``@A1:A4 B2:B9`` intersects after it.
_IMPLICIT_BINDING = BINDING[":"] - 1


def parse(formula: str) -> Node:
    """The tree of ``formula``, as a file stores it or as it is typed.

    A leading ``=`` is accepted and dropped. Raises
    :class:`FormulaSyntaxError` for text Excel would refuse.
    """
    text = formula[1:] if formula.startswith("=") else formula
    return _Parser(text, tokenize(text)).run()


class _Parser:
    def __init__(self, formula: str, tokens: list[Token]) -> None:
        self.formula = formula
        self.tokens = tokens
        self.index = 0

    def fail(self, message: str, token: Token | None = None) -> FormulaSyntaxError:
        where = (token or self.tokens[self.index]).start
        return FormulaSyntaxError(message, self.formula, where)

    def peek(self) -> Token:
        return self.tokens[self.index]

    def advance(self) -> Token:
        token = self.tokens[self.index]
        if token.kind is not Kind.END:
            self.index += 1
        return token

    def skip_space(self) -> None:
        while self.tokens[self.index].kind is Kind.SPACE:
            self.index += 1

    def run(self) -> Node:
        self.skip_space()
        if self.peek().kind is Kind.END:
            raise self.fail("a formula needs something to calculate")
        node = self.expression(0, union=True)
        self.skip_space()
        token = self.peek()
        if token.kind is not Kind.END:
            raise self.fail(f"unexpected {token.text!r}", token)
        return node

    def expression(self, floor: int, *, union: bool) -> Node:
        left = self.operand(union=union)
        while True:
            found = self.infix(union=union)
            if found is None:
                return left
            op, binding, resume = found
            if binding <= floor:
                return left
            self.index = resume
            if op in ("%", "#"):
                left = Postfix(op, left)
                continue
            right = self.expression(binding, union=union)
            left = _joined(op, left, right)

    def infix(self, *, union: bool) -> tuple[str, int, int] | None:
        """The infix or postfix operator next, how tightly it binds, and
        where the tokens resume after it, without consuming anything."""
        index = self.index
        while self.tokens[index].kind is Kind.SPACE:
            index += 1
        spaced = index > self.index
        token = self.tokens[index]
        if token.kind is Kind.OPERATOR:
            op = token.text
            if op == "%":
                return op, PERCENT_BINDING, index + 1
            if op == "#":
                return op, SPILL_BINDING, index + 1
            if op in BINDING:
                return op, BINDING[op], index + 1
            if op == "@" and spaced:
                return " ", BINDING[" "], index
            return None
        if token.kind is Kind.COMMA and union:
            return ",", BINDING[","], index + 1
        if spaced and token.kind in (Kind.OPERAND, Kind.FUNCTION, Kind.OPEN):
            # Two operands with only whitespace between: the intersection.
            return " ", BINDING[" "], index
        return None

    def operand(self, *, union: bool) -> Node:
        self.skip_space()
        token = self.advance()
        if token.kind is Kind.OPERAND:
            assert token.node is not None
            return token.node
        if token.kind is Kind.FUNCTION:
            return self.invoked(Call(token.text[:-1], self.arguments()))
        if token.kind is Kind.OPEN:
            inner = self.expression(0, union=True)
            self.skip_space()
            closing = self.advance()
            if closing.kind is not Kind.CLOSE:
                raise self.fail("expected ')'", closing)
            return self.invoked(Paren(inner))
        if token.kind is Kind.OPEN_ARRAY:
            return self.array()
        if token.kind is Kind.OPERATOR and token.text in ("-", "+"):
            return Unary(token.text, self.expression(PREFIX_BINDING, union=union))
        if token.kind is Kind.OPERATOR and token.text == "@":
            return Unary("@", self.expression(_IMPLICIT_BINDING, union=union))
        if token.kind is Kind.END:
            raise self.fail("the formula ends where a value should be", token)
        raise self.fail(f"unexpected {token.text!r}", token)

    def invoked(self, target: Node) -> Node:
        """``target`` called as many times as a ``(`` follows it straight
        away, as ``LAMBDA(x,x*2)(4)`` calls the function LAMBDA gives."""
        while self.peek().kind is Kind.OPEN:
            self.advance()
            target = Invoke(target, self.arguments())
        return target

    def arguments(self) -> tuple[Node, ...]:
        """A call's arguments, just after its ``(``, through its ``)``."""
        self.skip_space()
        if self.peek().kind is Kind.CLOSE:
            self.advance()
            return ()
        args: list[Node] = []
        while True:
            self.skip_space()
            if self.peek().kind in (Kind.COMMA, Kind.CLOSE):
                args.append(Missing())
            else:
                args.append(self.expression(0, union=False))
            self.skip_space()
            token = self.advance()
            if token.kind is Kind.CLOSE:
                return tuple(args)
            if token.kind is not Kind.COMMA:
                raise self.fail("expected ',' or ')'", token)

    def array(self) -> ArrayLiteral:
        """An array constant, just after its ``{``, through its ``}``."""
        rows: list[list[ArrayItem]] = [[]]
        while True:
            self.skip_space()
            rows[-1].append(self.array_item())
            self.skip_space()
            token = self.advance()
            if token.kind is Kind.COMMA:
                continue
            if token.kind is Kind.SEMICOLON:
                rows.append([])
                continue
            if token.kind is Kind.CLOSE_ARRAY:
                break
            raise self.fail("expected ',', ';' or '}'", token)
        if any(len(row) != len(rows[0]) for row in rows):
            raise self.fail("every row of an array constant needs as many items as the first")
        return ArrayLiteral(tuple(tuple(row) for row in rows))

    def array_item(self) -> ArrayItem:
        token = self.advance()
        sign = 1.0
        if token.kind is Kind.OPERATOR and token.text in ("-", "+"):
            sign = -1.0 if token.text == "-" else 1.0
            token = self.advance()
            if token.kind is not Kind.OPERAND or not isinstance(token.node, Number):
                raise self.fail("only a number can have a sign in an array constant", token)
        node = token.node if token.kind is Kind.OPERAND else None
        if isinstance(node, Number):
            return sign * node.value
        if isinstance(node, Text):
            return node.value
        if isinstance(node, Logical):
            return node.value
        if isinstance(node, ErrorLiteral):
            return CellError(node.code)
        raise self.fail("an array constant holds numbers, text, TRUE, FALSE and errors only", token)


def _joined(op: str, left: Node, right: Node) -> Node:
    """``left op right``, with a range between two cells on one sheet made
    the area it is: ``Data!A1:Data!B2`` and ``Data!A1:B2`` both mean the
    block on ``Data``."""
    if (
        op == ":"
        and isinstance(left, CellReference)
        and isinstance(right, CellReference)
        and (right.prefix is None or right.prefix == left.prefix)
    ):
        return AreaReference(left.ref, right.ref, left.prefix)
    return Binary(op, left, right)


__all__ = ["FormulaSyntaxError", "parse"]
