"""A Jet SQL executor over the storage engine.

``execute(db, sql, parameters)`` runs SELECT, INSERT, UPDATE and DELETE
statements against :class:`~pyopenvba.access.database.AccessDatabase`
tables, in pure Python.  SELECT covers a column list or ``*``, one table
or INNER / LEFT / RIGHT JOINs, WHERE, GROUP BY with Count, Sum, Avg, Min
and Max, First, Last, StDev, StDevP, Var and VarP, HAVING, ORDER BY (by
name or by position), DISTINCT and TOP (a count or a percentage); expressions cover the comparison, logical, arithmetic
and concatenation operators, LIKE with the engine's wildcards, IN,
BETWEEN, IS NULL, ``[parameters]`` and the functions a Jet expression can
name: text (Len, UCase, LCase, Trim, Left, Right, Mid, InStr, Replace,
Space, String, StrComp, StrReverse, Asc, Chr), maths (Abs, Int, Fix,
Round, Sgn, Sqr, Exp, Log), conversion (the C-family, Val, Str, Hex,
Oct), dates (Now, Date, Time, DateAdd, DateDiff, DatePart, DateSerial,
TimeSerial, Weekday, WeekdayName, MonthName, DateValue, TimeValue, Year
through Second), the yes-or-no tests (IsNull, IsNumeric, IsDate) and
IIf, Nz, Switch, Choose, Format and Partition.  Three-valued logic
follows the engine: a comparison against Null is Null, and a WHERE that
comes out Null drops the row.  A computed truth value comes back as -1
or 0, as the engine hands one back; only a Boolean column read on its
own keeps its type.

The statement text is split with the same clause splitter the saved-query
writer uses (:mod:`_queries`).  DML goes through the table writers, so a
DELETE with no WHERE takes the engine's truncation path and everything
else the row path, page for page as the engine would.
"""

from __future__ import annotations

import datetime as _dt
import struct
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from decimal import ROUND_HALF_EVEN, Decimal, localcontext
from typing import TYPE_CHECKING, Any

from pyopenvba.access._ddl import execute_ddl, is_ddl
from pyopenvba.access._format import format_value, partition
from pyopenvba.access._queries import (
    parse_from,
    select_list,
    split_alias,
    split_clauses,
    split_top_level,
    split_top_level_words,
)
from pyopenvba.access._schema import ColumnSpec
from pyopenvba.access._tdef import (
    FIXED_LENGTH_TYPES,
    OFFSET_NEXT_AUTONUMBER,
    OFFSET_ROW_COUNT,
    TYPE_COMPLEX,
    TYPE_BINARY,
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_MONEY,
    TYPE_NUMERIC,
    TYPE_TEXT,
    ColumnDef,
)
from pyopenvba.access_read import AccessError

if TYPE_CHECKING:
    from pyopenvba.access.database import AccessDatabase, RowId, Table

Row = dict[str, object]

# --- tokens ----------------------------------------------------------------------

_TOKEN = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<date>\#[^#]*\#)
  | (?P<string>'(?:[^']|'')*'|"(?:[^"]|"")*")
  | (?P<number>(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][-+]?\d+)?)
  | (?P<bracket>\[[^\]]+\])
  | (?P<name>[A-Za-z_][A-Za-z0-9_]*)
  | (?P<op><>|<=|>=|=|<|>|\+|-|\*|/|\\|\^|&|\(|\)|,|\.)
    """,
    re.VERBOSE,
)


@dataclass(frozen=True)
class Token:
    kind: str
    text: str
    start: int = 0
    end: int = 0


def tokenize(text: str) -> list[Token]:
    out: list[Token] = []
    pos = 0
    while pos < len(text):
        match = _TOKEN.match(text, pos)
        if not match:
            raise AccessError(f"cannot read SQL at {text[pos:pos + 20]!r}")
        pos = match.end()
        kind = match.lastgroup or ""
        if kind == "ws":
            continue
        out.append(Token(kind, match.group(0), match.start(), match.end()))
    return out


# --- expressions -----------------------------------------------------------------

Env = Callable[[str, str | None], object]


class Expr:
    def eval(self, env: Env) -> object:  # pragma: no cover - overridden
        raise NotImplementedError

    def columns(self) -> list[tuple[str | None, str]]:
        return []


@dataclass(frozen=True)
class Literal(Expr):
    value: object

    def eval(self, env: Env) -> object:
        return self.value


@dataclass(frozen=True)
class ColumnRef(Expr):
    qualifier: str | None
    name: str

    def eval(self, env: Env) -> object:
        return env(self.name, self.qualifier)

    def columns(self) -> list[tuple[str | None, str]]:
        return [(self.qualifier, self.name)]


@dataclass(frozen=True)
class Unary(Expr):
    op: str
    operand: Expr

    def eval(self, env: Env) -> object:
        value = self.operand.eval(env)
        if self.op == "NOT":
            return None if value is None else not _truthy(value)
        if value is None:
            return None
        return -_number(value)

    def columns(self) -> list[tuple[str | None, str]]:
        return self.operand.columns()


@dataclass(frozen=True)
class Binary(Expr):
    op: str
    left: Expr
    right: Expr

    def eval(self, env: Env) -> object:
        op = self.op
        if op == "AND":
            a = self.left.eval(env)
            if a is not None and not _truthy(a):
                return False
            b = self.right.eval(env)
            if b is not None and not _truthy(b):
                return False
            return None if a is None or b is None else True
        if op == "OR":
            a = self.left.eval(env)
            if a is not None and _truthy(a):
                return True
            b = self.right.eval(env)
            if b is not None and _truthy(b):
                return True
            return None if a is None or b is None else False
        a = self.left.eval(env)
        b = self.right.eval(env)
        if op == "IS":
            return a is None if b is None else a == b
        if op == "IS NOT":
            return a is not None if b is None else a != b
        if op == "&":
            # Jet's concatenation reads Null as an empty string; only Null
            # on both sides gives Null (measured against DAO).
            if a is None and b is None:
                return None
            return ("" if a is None else _text(a)) + ("" if b is None else _text(b))
        if a is None or b is None:
            return None
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return _compare(op, a, b)
        if op == "LIKE":
            return like_match(_text(a), _text(b))
        if op == "+":
            if isinstance(a, str) and isinstance(b, str):
                return a + b
            if isinstance(a, str) or isinstance(b, str):
                # Text on one side and a number on the other is an
                # addition, not a join: `'5' + 5` is 10.  Text that will
                # not read as a number makes the whole thing Null, which
                # is what the engine answers -- it does not refuse.
                try:
                    return _arith(_number(a), _number(b), lambda x, y: x + y)
                except AccessError:
                    return None
            return _arith(a, b, lambda x, y: x + y)
        if op == "-":
            return _arith(a, b, lambda x, y: x - y)
        if op == "*":
            return _arith(a, b, lambda x, y: x * y)
        if op == "/":
            # Dividing by zero is Null here, not an error: a query that
            # meets a zero in one row goes on and answers Null for it,
            # which is what the engine does (measured against DAO).
            if _number(b) == 0:
                return None
            return float(_number(a)) / float(_number(b))
        if op in ("\\", "MOD"):
            # Both sides round to whole numbers first, and the division
            # truncates toward zero, which is what VBA does.
            x, y = _whole(a), _whole(b)
            if y == 0:
                return None
            quotient = abs(x) // abs(y) * (1 if (x < 0) == (y < 0) else -1)
            return quotient if op == "\\" else x - y * quotient
        if op == "^":
            return float(_number(a)) ** float(_number(b))
        raise AccessError(f"unknown operator {op}")

    def columns(self) -> list[tuple[str | None, str]]:
        return self.left.columns() + self.right.columns()


@dataclass(frozen=True)
class InList(Expr):
    operand: Expr
    options: tuple[Expr, ...]
    negate: bool

    def eval(self, env: Env) -> object:
        value = self.operand.eval(env)
        if value is None:
            return None
        hit = any(_compare("=", value, option) for o in self.options if (option := o.eval(env)) is not None)
        return not hit if self.negate else hit

    def columns(self) -> list[tuple[str | None, str]]:
        return self.operand.columns() + [c for o in self.options for c in o.columns()]


@dataclass(frozen=True)
class Between(Expr):
    operand: Expr
    low: Expr
    high: Expr
    negate: bool

    def eval(self, env: Env) -> object:
        value, low, high = self.operand.eval(env), self.low.eval(env), self.high.eval(env)
        if value is None or low is None or high is None:
            return None
        hit = bool(_compare(">=", value, low)) and bool(_compare("<=", value, high))
        return not hit if self.negate else hit

    def columns(self) -> list[tuple[str | None, str]]:
        return self.operand.columns() + self.low.columns() + self.high.columns()


@dataclass(frozen=True)
class Call(Expr):
    name: str
    args: tuple[Expr, ...]
    star: bool = False

    @property
    def is_aggregate(self) -> bool:
        return self.name.upper() in AGGREGATES

    def eval(self, env: Env) -> object:
        upper = self.name.upper()
        if upper in DOMAIN_FUNCTIONS:
            return self._domain(env, upper)
        if self.is_aggregate:
            raise AccessError(f"{self.name} needs a GROUP BY or a whole-table aggregate context")
        values = [a.eval(env) for a in self.args]
        return _call(self.name, values)

    def _domain(self, env: Env, upper: str) -> object:
        """A domain function is a query over another table, so it runs as
        one: `DLookup("Total", "Orders", "Id = 1")` is
        `SELECT Total FROM Orders WHERE Id = 1`, and the aggregates wrap
        the expression in the aggregate of the same name.

        Its arguments are strings holding SQL, which is what lets a
        criteria mention a column of the table the expression names.
        """
        if not 2 <= len(self.args) <= 3:
            raise AccessError(f"{self.name} takes an expression, a domain and an optional criteria")
        values = [a.eval(env) for a in self.args]
        expression, domain = _text(values[0]), _text(values[1])
        criteria = _text(values[2]) if len(values) > 2 and values[2] is not None else ""
        if not expression or not domain:
            return None
        aggregate = DOMAIN_FUNCTIONS[upper]
        select = expression if aggregate is None else f"{aggregate}({expression})"
        sql = f"SELECT {select} FROM {_bracketed(domain)}"
        if criteria.strip():
            sql += f" WHERE {criteria}"
        # `Env` is declared as the callable that resolves a column; the
        # environment a running query passes also knows how to run SQL,
        # which is what a domain function needs.
        runner = getattr(env, "run", None)
        if runner is None:
            raise AccessError(f"{self.name} can only be used inside a query")
        rows = runner(sql)
        if not rows:
            # A count over nothing is nothing counted; the rest are Null.
            return 0 if upper == "DCOUNT" else None
        first = rows[0]
        return next(iter(first.values())) if first else None

    def columns(self) -> list[tuple[str | None, str]]:
        return [c for a in self.args for c in a.columns()]


AGGREGATES = {"COUNT", "SUM", "AVG", "MIN", "MAX", "FIRST", "LAST", "STDEV", "STDEVP", "VAR", "VARP"}
#: The domain functions, and the aggregate each wraps its expression in.
#: `DLookup` wraps it in nothing and takes the first row it finds.
DOMAIN_FUNCTIONS: dict[str, str | None] = {
    "DLOOKUP": None,
    "DCOUNT": "Count",
    "DSUM": "Sum",
    "DAVG": "Avg",
    "DMIN": "Min",
    "DMAX": "Max",
    "DFIRST": "First",
    "DLAST": "Last",
    "DSTDEV": "StDev",
    "DSTDEVP": "StDevP",
    "DVAR": "Var",
    "DVARP": "VarP",
}


def _bracketed(name: str) -> str:
    """A domain is a table or query name, and Access takes it either
    bracketed or bare."""
    trimmed = name.strip()
    if trimmed.startswith(("[", "(")) or " " not in trimmed:
        return trimmed
    return f"[{trimmed}]"


@dataclass(frozen=True)
class SubQuery(Expr):
    """A SELECT inside an expression.  ``kind`` says how its rows are read:
    ``scalar`` takes the first column of the first row, ``exists`` asks
    whether it returned any, ``in`` looks for a value among them, and
    ``quantified`` compares against every row (ALL) or against any one of
    them (ANY, which SOME is another spelling of)."""

    sql: str
    kind: str = "scalar"
    operand: Expr | None = None
    negate: bool = False
    comparison: str = "="
    every: bool = False

    def eval(self, env: Env) -> object:
        if not isinstance(env, Environment):
            raise AccessError("a subquery needs a query to run inside")
        rows = env.run(self.sql)
        if self.kind == "exists":
            return bool(rows) != self.negate
        values = [next(iter(r.values()), None) for r in rows]
        if self.kind == "quantified":
            if self.operand is None:
                raise AccessError("a quantified comparison needs a left side")
            value = self.operand.eval(env)
            if value is None:
                return None
            present = [v for v in values if v is not None]
            if not present:
                # ALL over nothing holds; ANY over nothing does not.
                return self.every
            tests = (_compare(self.comparison, value, v) for v in present)
            return all(tests) if self.every else any(tests)
        if self.kind == "in":
            if self.operand is None:
                raise AccessError("IN needs something to look for")
            value = self.operand.eval(env)
            if value is None:
                return None
            hit = any(_compare("=", value, v) for v in values if v is not None)
            return not hit if self.negate else hit
        if len(values) > 1:
            raise AccessError("a subquery used as a value returned more than one row")
        return values[0] if values else None

    def columns(self) -> list[tuple[str | None, str]]:
        return self.operand.columns() if self.operand is not None else []


def _number_literal(text: str) -> int | float:
    """What a number in the SQL text stands for.  One that names a whole
    number -- `1E3` and `1.5E2` included -- is an integer; the rest are
    read as floating point, which is how the executor carries them."""
    if "." not in text and "E" not in text.upper():
        return int(text)
    exact = Decimal(text)
    if exact == exact.to_integral_value():
        return int(exact)
    return float(text)


class Parser:
    """Precedence-climbing parser for Jet SQL expressions."""

    def __init__(self, tokens: list[Token], text: str = "") -> None:
        self.tokens = tokens
        self.text = text
        self.pos = 0

    @classmethod
    def parse(cls, text: str) -> Expr:
        parser = cls(tokenize(text), text)
        expr = parser.expression()
        if parser.pos != len(parser.tokens):
            raise AccessError(f"unexpected {parser.tokens[parser.pos].text!r} in {text!r}")
        return expr

    def peek(self, *words: str) -> Token | None:
        if self.pos < len(self.tokens):
            token = self.tokens[self.pos]
            if not words or token.text.upper() in words:
                return token
        return None

    def take(self, *words: str) -> Token:
        token = self.peek(*words)
        if token is None:
            expected = " or ".join(words) if words else "more"
            got = self.tokens[self.pos].text if self.pos < len(self.tokens) else "end"
            raise AccessError(f"expected {expected}, got {got!r}")
        self.pos += 1
        return token

    def expression(self) -> Expr:
        return self.disjunction()

    def disjunction(self) -> Expr:
        left = self.conjunction()
        while self.peek("OR"):
            self.take()
            left = Binary("OR", left, self.conjunction())
        return left

    def conjunction(self) -> Expr:
        left = self.negation()
        while self.peek("AND"):
            self.take()
            left = Binary("AND", left, self.negation())
        return left

    def negation(self) -> Expr:
        if self.peek("NOT"):
            self.take()
            if self.peek("EXISTS"):
                self.take()
                return SubQuery(self._subquery_text(), "exists", negate=True)
            return Unary("NOT", self.negation())
        if self.peek("EXISTS"):
            self.take()
            return SubQuery(self._subquery_text(), "exists")
        return self.comparison()

    def _subquery_text(self) -> str:
        """The text of a parenthesized SELECT starting at the next token."""
        if not self.peek("("):
            raise AccessError("expected ( before a subquery")
        opening = self.take()
        depth = 1
        while self.pos < len(self.tokens):
            token = self.tokens[self.pos]
            self.pos += 1
            if token.text == "(":
                depth += 1
            elif token.text == ")":
                depth -= 1
                if depth == 0:
                    return self.text[opening.end : token.start].strip()
        raise AccessError("a subquery is missing its closing bracket")

    def _at_subquery(self) -> bool:
        return (
            self.peek("(") is not None
            and self.pos + 1 < len(self.tokens)
            and self.tokens[self.pos + 1].text.upper() in ("SELECT", "PARAMETERS", "TRANSFORM")
        )

    def comparison(self) -> Expr:
        left = self.concatenation()
        negate = False
        if self.peek("NOT"):
            self.take()
            negate = True
        if self.peek("IS"):
            self.take()
            if self.peek("NOT"):
                self.take()
                negate = not negate
            self.take("NULL")
            return Binary("IS NOT" if negate else "IS", left, Literal(None))
        if self.peek("LIKE"):
            self.take()
            expr: Expr = Binary("LIKE", left, self.concatenation())
            return Unary("NOT", expr) if negate else expr
        if self.peek("IN"):
            self.take()
            if self._at_subquery():
                return SubQuery(self._subquery_text(), "in", operand=left, negate=negate)
            self.take("(")
            options: list[Expr] = [self.expression()]
            while self.peek(","):
                self.take()
                options.append(self.expression())
            self.take(")")
            return InList(left, tuple(options), negate)
        if self.peek("BETWEEN"):
            self.take()
            low = self.concatenation()
            self.take("AND")
            return Between(left, low, self.concatenation(), negate)
        if negate:
            raise AccessError("NOT must be followed by LIKE, IN, BETWEEN or IS")
        token = self.peek("=", "<>", "<", ">", "<=", ">=")
        if token is not None:
            self.take()
            quantifier = self.peek("ALL", "ANY", "SOME")
            if quantifier is not None:
                self.take()
                if not self._at_subquery():
                    raise AccessError(f"{quantifier.text.upper()} needs a subquery")
                return SubQuery(self._subquery_text(), "quantified", operand=left, comparison=token.text,
                                every=quantifier.text.upper() == "ALL")
            return Binary(token.text, left, self.concatenation())
        return left

    def concatenation(self) -> Expr:
        left = self.additive()
        while self.peek("&"):
            self.take()
            left = Binary("&", left, self.additive())
        return left

    def additive(self) -> Expr:
        left = self.modulo()
        while (token := self.peek("+", "-")) is not None:
            self.take()
            left = Binary(token.text, left, self.modulo())
        return left

    def modulo(self) -> Expr:
        left = self.int_division()
        while self.peek("MOD"):
            self.take()
            left = Binary("MOD", left, self.int_division())
        return left

    def int_division(self) -> Expr:
        left = self.multiplicative()
        while self.peek("\\"):
            self.take()
            left = Binary("\\", left, self.multiplicative())
        return left

    def multiplicative(self) -> Expr:
        left = self.unary()
        while (token := self.peek("*", "/")) is not None:
            self.take()
            left = Binary(token.text, left, self.unary())
        return left

    def unary(self) -> Expr:
        if self.peek("-"):
            self.take()
            return Unary("-", self.unary())
        if self.peek("+"):
            self.take()
            return self.unary()
        return self.power()

    def power(self) -> Expr:
        left = self.atom()
        while self.peek("^"):
            self.take()
            left = Binary("^", left, self.unary())
        return left

    def atom(self) -> Expr:
        token = self.take()
        if token.kind == "number":
            return Literal(_number_literal(token.text))
        if token.kind == "string":
            quote = token.text[0]
            return Literal(token.text[1:-1].replace(quote + quote, quote))
        if token.kind == "date":
            return Literal(parse_date_literal(token.text[1:-1]))
        if token.text == "(":
            if self.pos < len(self.tokens) and self.tokens[self.pos].text.upper() in ("SELECT", "PARAMETERS", "TRANSFORM"):
                self.pos -= 1
                return SubQuery(self._subquery_text())
            inner = self.expression()
            self.take(")")
            return inner
        if token.kind in ("name", "bracket"):
            name = token.text[1:-1] if token.kind == "bracket" else token.text
            upper = name.upper()
            if upper == "NULL":
                return Literal(None)
            if upper in ("TRUE", "YES", "ON"):
                return Literal(True)
            if upper in ("FALSE", "NO", "OFF"):
                return Literal(False)
            if self.peek("(") and token.kind == "name":
                self.take()
                if self.peek("*"):
                    self.take()
                    self.take(")")
                    return Call(name, (), star=True)
                args: list[Expr] = []
                if not self.peek(")"):
                    args.append(self.expression())
                    while self.peek(","):
                        self.take()
                        args.append(self.expression())
                self.take(")")
                return Call(name, tuple(args))
            if self.peek("."):
                self.take()
                column = self.take()
                if column.kind not in ("name", "bracket"):
                    raise AccessError(f"expected a column after {name}.")
                return ColumnRef(name, column.text[1:-1] if column.kind == "bracket" else column.text)
            return ColumnRef(None, name)
        raise AccessError(f"unexpected {token.text!r}")


# --- values ----------------------------------------------------------------------


def _computed(value: object) -> object:
    """A value on its way out of an expression.  Jet has no Boolean of its
    own: a comparison, a logical operator or a function that answers yes
    or no gives back -1 or 0, where a Boolean column keeps its type
    (measured -- every one of `N > 10`, `Not (...)`, `IsNull(T)` and the
    literal `True` came out of the engine as -1 or 0)."""
    return -1 if value is True else (0 if value is False else value)


def _truthy(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        return value != 0
    if isinstance(value, str):
        return value != ""
    return value is not None


def _whole(value: object) -> int:
    """A value as the whole number VBA makes of it, halves to even."""
    number = _number(value)
    if isinstance(number, int):
        return number
    return int(Decimal(str(number)).quantize(Decimal(1), rounding=ROUND_HALF_EVEN))


def _round_half_even(value: float, places: int) -> float:
    """Round the number as it reads, not as the machine holds it.

    The engine rounds half to even -- 2.345 to two places is 2.34 and
    2.355 is 2.36 -- and it does so on the decimal that was written, not
    on the binary double nearest it.  The two disagree often: the double
    for 2.345 sits just above the decimal and the double for 2.675 just
    below, so rounding the double gives 2.35 and 2.67 where the engine
    answers 2.34 and 2.68.
    """
    if value != value or value in (float("inf"), float("-inf")):
        return value
    quantum = Decimal(1).scaleb(-places)
    rounded = Decimal(repr(value)).quantize(quantum, rounding=ROUND_HALF_EVEN)
    return float(rounded)


#: What a text comparison argument means: 0 compares by character code,
#: anything else -- and the default in a query -- ignores case.
COMPARE_BINARY = 0


def _replace(
    text: str, find: str, put: str, start: int, count: int, blind: bool
) -> str:
    """`Replace` as a query means it.

    `start` does not just say where to look: the answer begins there, so
    `Replace("abcabc", "b", "X", 3)` is `caXc` and not `abcaXc`.  `count`
    caps how many go, and the search ignores case unless told otherwise.
    """
    if start < 1:
        raise AccessError("Replace starts at 1 or later")
    text = text[start - 1 :]
    if not find or count == 0:
        return text
    if not blind:
        return text.replace(find, put) if count < 0 else text.replace(find, put, count)
    out: list[str] = []
    hay, pin = text.lower(), find.lower()
    at = done = 0
    while count < 0 or done < count:
        found = hay.find(pin, at)
        if found < 0:
            break
        out.append(text[at:found])
        out.append(put)
        at = found + len(find)
        done += 1
    out.append(text[at:])
    return "".join(out)


def _case_blind(args: Sequence[object], at: int) -> bool:
    """Whether a text function should ignore case, given where its
    comparison argument sits."""
    if len(args) <= at or args[at] is None:
        return True
    return int(_number(args[at])) != COMPARE_BINARY


def _number(value: object) -> int | float | Decimal:
    if isinstance(value, bool):
        return -1 if value else 0
    if isinstance(value, (int, float, Decimal)):
        return value
    if isinstance(value, str):
        text = value.strip()
        try:
            return int(text)
        except ValueError:
            pass
        try:
            # Reaches an exponent as well as a decimal point: the engine
            # reads `1e3` as a thousand.
            return float(text)
        except ValueError as exc:
            raise AccessError(f"{value!r} is not a number") from exc
    if isinstance(value, _dt.datetime):
        import struct

        from pyopenvba.access._rows import encode_datetime

        return struct.unpack("<d", encode_datetime(value))[0]
    raise AccessError(f"{value!r} is not a number")


def _arith(a: object, b: object, fn: Callable[[Any, Any], Any]) -> object:
    x, y = _number(a), _number(b)
    if isinstance(x, Decimal) or isinstance(y, Decimal):
        return fn(Decimal(str(x)), Decimal(str(y)))
    return fn(x, y)


def _text(value: object) -> str:
    if isinstance(value, bool):
        # A query writes a Boolean as the number it is: `True & ''` is
        # `-1`, not `True`.
        return "-1" if value else "0"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, _dt.datetime):
        return value.strftime("%#m/%#d/%Y %H:%M:%S") if value.time() != _dt.time() else value.strftime("%#m/%#d/%Y")
    return str(value)


def _compare(op: str, a: object, b: object) -> bool:
    x: Any
    y: Any
    if isinstance(a, str) and isinstance(b, str):
        x, y = a.lower(), b.lower()
    elif isinstance(a, str) or isinstance(b, str):
        x, y = _text(a).lower(), _text(b).lower()
    elif isinstance(a, _dt.datetime) or isinstance(b, _dt.datetime):
        x, y = _number(a), _number(b)
    else:
        x, y = _number(a), _number(b)
    if op == "=":
        return x == y
    if op == "<>":
        return x != y
    if op == "<":
        return bool(x < y)
    if op == ">":
        return bool(x > y)
    if op == "<=":
        return bool(x <= y)
    return bool(x >= y)


def like_match(value: str, pattern: str) -> bool:
    """The engine's LIKE: ``*`` any run, ``?`` one character, ``#`` one
    digit, ``[abc]`` / ``[!a-z]`` character lists, case-blind."""
    regex = ""
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            regex += ".*"
        elif ch == "?":
            regex += "."
        elif ch == "#":
            regex += r"\d"
        elif ch == "[":
            end = pattern.find("]", i + 1)
            if end < 0:
                raise AccessError(f"unterminated character list in {pattern!r}")
            body = pattern[i + 1 : end]
            if body.startswith("!"):
                body = "^" + body[1:]
            regex += "[" + body.replace("\\", "\\\\") + "]"
            i = end
        else:
            regex += re.escape(ch)
        i += 1
    return re.fullmatch(regex, value, re.IGNORECASE | re.DOTALL) is not None


def parse_date_literal(text: str) -> _dt.datetime:
    text = text.strip()
    for form in ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%H:%M:%S"):
        try:
            value = _dt.datetime.strptime(text, form)
            if form == "%H:%M:%S":
                value = value.replace(year=1899, month=12, day=30)
            return value
        except ValueError:
            continue
    raise AccessError(f"cannot read date #{text}#")


WEEKDAY_NAMES = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)


def _as_date(value: object, name: str) -> _dt.datetime:
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime.combine(value, _dt.time())
    if isinstance(value, str):
        return parse_date_literal(value)
    if isinstance(value, (int, float, Decimal)):
        return _from_serial(float(value))
    raise AccessError(f"{name} needs a date, not {value!r}")


def _from_serial(serial: float) -> _dt.datetime:
    """A stored date: whole days from 1899-12-30, the fraction the time.
    A negative serial counts the time forward from midnight all the same,
    which is how the engine reads one."""
    whole = int(serial) if serial >= 0 else -int(-serial)
    day = _dt.datetime(1899, 12, 30) + _dt.timedelta(days=whole)
    return day + _dt.timedelta(seconds=round(abs(serial - whole) * 86400))


def _add_months(when: _dt.datetime, months: int) -> _dt.datetime:
    """Jet keeps the day of the month, clamped to the target month's last
    day (measured: adding a month to 31 January gives 29 February)."""
    import calendar

    total = when.month - 1 + months
    year, month = when.year + total // 12, total % 12 + 1
    day = min(when.day, calendar.monthrange(year, month)[1])
    return when.replace(year=year, month=month, day=day)


def _date_add(interval: str, count: int, when: _dt.datetime) -> _dt.datetime:
    key = interval.lower()
    if key == "yyyy":
        return _add_months(when, 12 * count)
    if key == "q":
        return _add_months(when, 3 * count)
    if key == "m":
        return _add_months(when, count)
    if key in ("d", "y", "w"):
        return when + _dt.timedelta(days=count)
    if key == "ww":
        return when + _dt.timedelta(weeks=count)
    if key == "h":
        return when + _dt.timedelta(hours=count)
    if key == "n":
        return when + _dt.timedelta(minutes=count)
    if key == "s":
        return when + _dt.timedelta(seconds=count)
    raise AccessError(f"DateAdd has no interval {interval!r}")


def _date_diff(interval: str, first: _dt.datetime, second: _dt.datetime) -> int:
    """Jet counts boundaries crossed, not whole units: a day from 23:00 to
    01:00 is one day, and December to January is one month."""
    key = interval.lower()
    if key == "yyyy":
        return second.year - first.year
    if key == "q":
        return (second.year * 4 + (second.month - 1) // 3) - (first.year * 4 + (first.month - 1) // 3)
    if key == "m":
        return (second.year * 12 + second.month) - (first.year * 12 + first.month)
    if key in ("d", "y"):
        return (second.date() - first.date()).days
    if key == "w":
        return (second.date() - first.date()).days // 7
    if key == "ww":
        return (_week_start(second) - _week_start(first)).days // 7
    if key in ("h", "n", "s"):
        # Boundaries again: the two dates are counted in whole hours,
        # minutes or seconds from one fixed point and subtracted.
        unit = {"h": 3600, "n": 60, "s": 1}[key]
        return _units(second, unit) - _units(first, unit)
    raise AccessError(f"DateDiff has no interval {interval!r}")


def _units(when: _dt.datetime, seconds: int) -> int:
    """How many whole units of ``seconds`` have passed at ``when``, from
    the engine's own zero (1899-12-30), rounding down."""
    return int((when - _dt.datetime(1899, 12, 30)).total_seconds() // seconds)


def _week_start(when: _dt.datetime) -> _dt.date:
    """The Sunday of this date's week, which is where Jet's weeks start."""
    return when.date() - _dt.timedelta(days=(when.date().weekday() + 1) % 7)


def _date_part(interval: str, when: _dt.datetime) -> int:
    key = interval.lower()
    if key == "yyyy":
        return when.year
    if key == "q":
        return (when.month - 1) // 3 + 1
    if key == "m":
        return when.month
    if key == "y":
        return when.timetuple().tm_yday
    if key == "d":
        return when.day
    if key == "w":
        return (when.weekday() + 1) % 7 + 1
    if key == "ww":
        return (when.date() - _week_start(_dt.datetime(when.year, 1, 1))).days // 7 + 1
    if key == "h":
        return when.hour
    if key == "n":
        return when.minute
    if key == "s":
        return when.second
    raise AccessError(f"DatePart has no interval {interval!r}")


def _call(name: str, args: list[object]) -> object:
    upper = name.upper()
    if upper == "IIF":
        if len(args) != 3:
            raise AccessError("IIf takes three arguments")
        return args[1] if _truthy(args[0]) and args[0] is not None else args[2]
    if upper == "NZ":
        if not args:
            raise AccessError("Nz takes one or two arguments")
        return args[0] if args[0] is not None else (args[1] if len(args) > 1 else 0)
    if upper == "NOW":
        return _dt.datetime.now().replace(microsecond=0)
    if upper == "DATE":
        return _dt.datetime.combine(_dt.date.today(), _dt.time())
    if upper == "TIME":
        now = _dt.datetime.now()
        return _dt.datetime(1899, 12, 30, now.hour, now.minute, now.second)
    if upper == "ISNULL":
        return args[0] is None
    if upper == "ISNUMERIC":
        return _is_numeric(args[0])
    if upper == "ISDATE":
        return _is_date(args[0])
    if upper == "FORMAT":
        if not 1 <= len(args) <= 2:
            raise AccessError("Format takes a value and a pattern")
        return format_value(args[0], _text(args[1]) if len(args) > 1 else "")
    if upper == "SWITCH":
        # Pairs of condition and answer; the first true one wins and Null
        # comes back when none does.
        for i in range(0, len(args) - 1, 2):
            if args[i] is not None and _truthy(args[i]):
                return args[i + 1]
        return None
    if upper == "CHOOSE":
        if args[0] is None:
            return None
        index = int(_number(args[0]))
        return args[index] if 1 <= index < len(args) else None
    if any(a is None for a in args):
        return None
    if upper == "LEN":
        return len(_text(args[0]))
    if upper == "UCASE":
        return _text(args[0]).upper()
    if upper == "LCASE":
        return _text(args[0]).lower()
    if upper == "TRIM":
        return _text(args[0]).strip(" ")
    if upper == "LTRIM":
        return _text(args[0]).lstrip(" ")
    if upper == "RTRIM":
        return _text(args[0]).rstrip(" ")
    if upper == "LEFT":
        return _text(args[0])[: int(_number(args[1]))]
    if upper == "RIGHT":
        count = int(_number(args[1]))
        return _text(args[0])[-count:] if count else ""
    if upper == "MID":
        text = _text(args[0])
        start = int(_number(args[1])) - 1
        return text[start : start + int(_number(args[2]))] if len(args) > 2 else text[start:]
    if upper == "INSTR":
        wide = len(args) > 2
        haystack, needle = (
            (_text(args[1]), _text(args[2])) if wide else (_text(args[0]), _text(args[1]))
        )
        start = int(_number(args[0])) if wide else 1
        blind = _case_blind(args, 3)
        hay = haystack.lower() if blind else haystack
        pin = needle.lower() if blind else needle
        return hay.find(pin, start - 1) + 1
    if upper == "ABS":
        return abs(_number(args[0]))  # pyright: ignore[reportArgumentType]
    if upper == "INT":
        import math

        return math.floor(float(_number(args[0])))
    if upper == "ROUND":
        places = int(_number(args[1])) if len(args) > 1 else 0
        return _round_half_even(float(_number(args[0])), places)
    if upper in ("YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND"):
        when = args[0]
        if not isinstance(when, _dt.datetime):
            raise AccessError(f"{name} needs a date")
        return getattr(when, upper.lower())
    if upper == "CSTR":
        return _text(args[0])
    if upper in ("CLNG", "CINT"):
        return round(float(_number(args[0])))
    if upper == "CDBL":
        return float(_number(args[0]))
    if upper == "CSNG":
        import struct as _struct

        return _struct.unpack("<f", _struct.pack("<f", float(_number(args[0]))))[0]
    if upper == "CBOOL":
        return _truthy(args[0])
    if upper == "CBYTE":
        value = round(float(_number(args[0])))
        if not 0 <= value <= 255:
            raise AccessError("CByte takes 0 to 255")
        return value
    if upper == "CCUR":
        return Decimal(str(float(_number(args[0])))).quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
    if upper == "CDATE":
        return _as_date(args[0], "CDate")
    if upper == "DATEVALUE":
        return _dt.datetime.combine(_as_date(args[0], "DateValue").date(), _dt.time())
    if upper == "TIMEVALUE":
        when = _as_date(args[0], "TimeValue")
        return _dt.datetime(1899, 12, 30, when.hour, when.minute, when.second)
    if upper == "DATESERIAL":
        year, month, day = (int(_number(a)) for a in args[:3])
        return _add_months(_dt.datetime(year, 1, 1), month - 1) + _dt.timedelta(days=day - 1)
    if upper == "TIMESERIAL":
        hour, minute, second = (int(_number(a)) for a in args[:3])
        return _dt.datetime(1899, 12, 30) + _dt.timedelta(hours=hour, minutes=minute, seconds=second)
    if upper == "DATEADD":
        return _date_add(_text(args[0]), int(_number(args[1])), _as_date(args[2], "DateAdd"))
    if upper == "DATEDIFF":
        return _date_diff(_text(args[0]), _as_date(args[1], "DateDiff"), _as_date(args[2], "DateDiff"))
    if upper == "DATEPART":
        return _date_part(_text(args[0]), _as_date(args[1], "DatePart"))
    if upper == "WEEKDAY":
        first = int(_number(args[1])) if len(args) > 1 else 1
        return ((_as_date(args[0], "Weekday").weekday() + 1) % 7) - (first - 1) % 7 + 1
    if upper == "WEEKDAYNAME":
        index = int(_number(args[0]))
        if not 1 <= index <= 7:
            raise AccessError("WeekdayName takes 1 to 7")
        return WEEKDAY_NAMES[index - 1]
    if upper == "MONTHNAME":
        index = int(_number(args[0]))
        if not 1 <= index <= 12:
            raise AccessError("MonthName takes 1 to 12")
        return MONTH_NAMES[index - 1]
    if upper == "REPLACE":
        return _replace(
            _text(args[0]),
            _text(args[1]),
            _text(args[2]),
            int(_number(args[3])) if len(args) > 3 and args[3] is not None else 1,
            int(_number(args[4])) if len(args) > 4 and args[4] is not None else -1,
            _case_blind(args, 5),
        )
    if upper == "SPACE":
        return " " * int(_number(args[0]))
    if upper == "STRING":
        fill = _text(args[1])
        return (fill[:1] if fill else " ") * int(_number(args[0]))
    if upper == "STRCOMP":
        a, b = _text(args[0]), _text(args[1])
        if _case_blind(args, 2):
            a, b = a.lower(), b.lower()
        return (a > b) - (a < b)
    if upper == "STRREVERSE":
        return _text(args[0])[::-1]
    if upper == "ASC":
        text = _text(args[0])
        if not text:
            raise AccessError("Asc needs a character")
        return ord(text[0])
    if upper == "CHR":
        return chr(int(_number(args[0])))
    if upper == "SGN":
        value = float(_number(args[0]))
        return (value > 0) - (value < 0)
    if upper == "SQR":
        import math

        value = float(_number(args[0]))
        if value < 0:
            raise AccessError("Sqr needs a number that is not negative")
        return math.sqrt(value)
    if upper == "EXP":
        import math

        return math.exp(float(_number(args[0])))
    if upper == "LOG":
        import math

        value = float(_number(args[0]))
        if value <= 0:
            raise AccessError("Log needs a number above zero")
        return math.log(value)
    if upper == "FIX":
        value = float(_number(args[0]))
        return int(value) if value >= 0 else -int(-value)
    if upper == "VAL":
        return _val(_text(args[0]))
    if upper == "STR":
        # A leading space stands in for the sign, and a value under one
        # loses its leading zero, which is what VBA writes.
        text = _text(_number(args[0]))
        for zero, without in (("-0.", "-."), ("0.", ".")):
            if text.startswith(zero):
                text = without + text[len(zero) :]
                break
        return text if text.startswith("-") else " " + text
    if upper == "HEX":
        value = int(_number(args[0]))
        return format(value & 0xFFFFFFFF if value < 0 else value, "X")
    if upper == "OCT":
        value = int(_number(args[0]))
        return format(value & 0xFFFFFFFF if value < 0 else value, "o")
    if upper == "PARTITION":
        if len(args) != 4:
            raise AccessError("Partition takes a number, a start, a stop and an interval")
        return partition(float(_number(args[0])), *(int(_number(a)) for a in args[1:4]))
    raise AccessError(f"function {name} is not available")


def _is_numeric(value: object) -> bool:
    if isinstance(value, bool) or value is None:
        return False
    if isinstance(value, (int, float, Decimal)):
        return True
    if not isinstance(value, str):
        return False
    try:
        float(value.strip())
    except ValueError:
        return False
    return True


def _is_date(value: object) -> bool:
    if isinstance(value, (_dt.datetime, _dt.date)):
        return True
    if not isinstance(value, str):
        return False
    try:
        parse_date_literal(value)
    except AccessError:
        return False
    return True


def _val(text: str) -> int | float:
    """As much of the front of the text as reads as a number, ignoring
    spaces, else 0."""
    body = text.replace(" ", "")
    best: int | float = 0
    for end in range(len(body), 0, -1):
        try:
            best = float(body[:end])
        except ValueError:
            continue
        return int(best) if best.is_integer() and "." not in body[:end] else best
    return best


# --- statements ------------------------------------------------------------------


@dataclass
class Source:
    """One thing a FROM clause names -- a table, a saved query or a
    bracketed SELECT -- and how it is joined to the ones before it."""

    alias: str
    columns: list[str]
    rows: list[Row]
    table: Table | None = None
    join: str | None = None  # None, INNER, LEFT, RIGHT, CROSS
    condition: Expr | None = None

    def holds(self, name: str) -> bool:
        return any(c.lower() == name.lower() for c in self.columns)


def _table_source(table: Table, alias: str) -> Source:
    rows: list[Row] = []
    for row_id, values in table.rows_with_ids():
        out: Row = {f"{alias}.{name}".lower(): value for name, value in values.items()}
        out["__rowid__." + alias.lower()] = row_id
        rows.append(out)
    return Source(alias=alias, columns=[c.name for c in table.columns], rows=rows, table=table)


def _rows_source(alias: str, columns: list[str], rows: list[Row]) -> Source:
    return Source(
        alias=alias,
        columns=columns,
        rows=[{f"{alias}.{name}".lower(): value for name, value in row.items()} for row in rows],
    )


class Environment:
    """What a column name means inside one row of one query, and how a
    subquery in that query runs.  A subquery falls back to the query it
    sits in, which is what makes a correlated one work."""

    def __init__(
        self,
        row: Row,
        sources: Sequence[Source],
        parameters: Mapping[str, object],
        runner: Callable[[str, Environment], list[Row]] | None = None,
        outer: Environment | None = None,
    ) -> None:
        self.row = row
        self.sources = sources
        self.parameters = parameters
        self.runner = runner or (outer.runner if outer is not None else None)
        self.outer = outer
        self.aliases = [s.alias.lower() for s in sources]

    def run(self, sql: str) -> list[Row]:
        if self.runner is None:
            raise AccessError("a subquery needs a query to run inside")
        return self.runner(sql, self)

    def __call__(self, name: str, qualifier: str | None) -> object:
        key = name.lower()
        if qualifier is not None:
            q = qualifier.lower()
            if q in self.aliases:
                full = f"{q}.{key}"
                if full in self.row:
                    return self.row[full]
                if self.sources[self.aliases.index(q)].holds(key):
                    return None
                raise AccessError(f"no column {qualifier}.{name}")
            if self.outer is not None:
                return self.outer(name, qualifier)
            raise AccessError(f"no table or alias {qualifier!r} in the query")
        hits = [f"{a}.{key}" for a in self.aliases if f"{a}.{key}" in self.row]
        if len(hits) > 1:
            raise AccessError(f"column {name!r} is ambiguous; qualify it")
        if hits:
            return self.row[hits[0]]
        for source in self.sources:
            if source.holds(key):
                return None  # a missing side of an outer join
        for pname, pvalue in self.parameters.items():
            if pname.lower().strip("[]") == key:
                return pvalue
        if self.outer is not None:
            return self.outer(name, None)
        raise AccessError(f"no column or parameter {name!r}")


def _environment(
    row: Row,
    sources: Sequence[Source],
    parameters: Mapping[str, object],
    outer: Environment | None = None,
    runner: Callable[[str, Environment], list[Row]] | None = None,
) -> Environment:
    return Environment(row, sources, parameters, runner=runner, outer=outer)


def _join(sources: list[Source], parameters: Mapping[str, object], outer: Environment | None = None) -> list[Row]:
    """The rows of the FROM clause, in the order the engine produces them.

    The order matters: it is what a SELECT with no ORDER BY answers in,
    and what an UPDATE over the join writes in.  Measured against the
    engine (docs/access_engine.md): an inner join scans the side with
    fewer rows and probes the other through a temporary index, which
    hands back rows sharing a key in the reverse of their stored order;
    on a tie the second-listed side is scanned.  Joins fold left to
    right, the result so far counting as the smaller of its two inputs
    and, on a tie with a table, being the side scanned.  An outer join
    scans its preserved side.  A cross join loops the larger side
    outside (the second-listed on a tie) and the smaller inside, both in
    stored order."""
    rows: list[Row] = list(sources[0].rows)
    estimate = len(rows)
    folded = False
    for k in range(1, len(sources)):
        source = sources[k]
        right_rows = source.rows
        scope = sources[: k + 1]

        def matches(a: Row, b: Row) -> bool:
            if source.condition is None:
                return True
            verdict = source.condition.eval(_environment({**a, **b}, scope, parameters, outer))
            return verdict is not None and _truthy(verdict)

        joined: list[Row] = []
        if source.join == "CROSS":
            outer_side, inner_side = (right_rows, rows) if len(right_rows) >= len(rows) else (rows, right_rows)
            joined = [{**a, **b} for a in outer_side for b in inner_side]
        elif source.join == "RIGHT" or (
            source.join != "LEFT" and (len(right_rows) < estimate or (len(right_rows) == estimate and not folded))
        ):
            # Scan the right side, probe what came before it.
            for right in right_rows:
                hits = [{**left, **right} for left in reversed(rows) if matches(left, right)]
                joined.extend(hits)
                if not hits and source.join == "RIGHT":
                    joined.append(dict(right))
        else:
            # Scan what came before, probe the right side.
            for left in rows:
                hits = [{**left, **right} for right in reversed(right_rows) if matches(left, right)]
                joined.extend(hits)
                if not hits and source.join == "LEFT":
                    joined.append(dict(left))
        rows = joined
        estimate = len(rows) if source.join == "CROSS" else min(estimate, len(right_rows))
        folded = True
    return rows


#: The engine's Decimal is 96 bits wide, so it carries as many digits as
#: fit in that -- 29 where the value allows, not Python's default 28.
DECIMAL_DIGITS = 29
DECIMAL_LIMIT = 2**96 - 1


def _divide_decimal(total: Decimal, count: int) -> Decimal:
    """Divide as the engine divides a Decimal: to as many digits as its
    96-bit decimal holds.

    An exact quotient keeps the places it needs and no more, so
    `12346.92 / 3` is `4115.64`; an inexact one runs to the width of the
    number, so `12346.9289 / 3` is `4115.6429666666666666666666667`.
    """
    with localcontext() as context:
        context.prec = DECIMAL_DIGITS
        quotient = total / count
    while _unscaled(quotient) > DECIMAL_LIMIT:
        _sign, _digits, exponent = quotient.as_tuple()
        quotient = quotient.quantize(Decimal(1).scaleb(int(exponent) + 1))
    return quotient


def _unscaled(value: Decimal) -> int:
    """The digits of the number without its decimal point, which is what
    has to fit the engine's 96 bits."""
    _sign, digits, _exponent = value.as_tuple()
    return int("".join(str(d) for d in digits) or "0")


def _money_column(expr: Expr, sources: Sequence[Source]) -> bool:
    """Whether the expression is a plain read of a Currency column.

    Currency and Decimal both arrive as `Decimal`, and the engine averages
    them differently, so the value alone cannot say which it is -- only
    the column it came from can.
    """
    if not isinstance(expr, ColumnRef):
        return False
    for source in sources:
        if expr.qualifier and source.alias.lower() != expr.qualifier.lower():
            continue
        if source.table is None or not source.holds(expr.name):
            continue
        try:
            return source.table.definition.column(expr.name).type_code == TYPE_MONEY
        except AccessError:
            return False
    return False


def _aggregate(name: str, values: list[object], money: bool = False) -> object:
    """Every aggregate answers with a number where its values are truth
    values, the Boolean column included: measured on a column of True,
    False, True, Max was 0, Min and First -1 and Sum -2."""
    upper = name.upper()
    present = [_computed(v) for v in values if v is not None]
    if upper == "COUNT":
        return len(present)
    if not present:
        return None
    if upper == "SUM":
        total: object = present[0]
        for v in present[1:]:
            total = _arith(total, v, lambda x, y: x + y)
        return total
    if upper == "AVG":
        total = present[0]
        for v in present[1:]:
            total = _arith(total, v, lambda x, y: x + y)
        if isinstance(total, Decimal):
            average = _divide_decimal(total, len(present))
            # Averaging a Currency column rounds to the four places
            # Currency holds; averaging a Decimal one keeps every digit
            # the division gives, which is 28 of them.  Both arrive here
            # as `Decimal`, so which it is comes from the column.
            return average.quantize(Decimal("0.0001")) if money else average
        return float(_number(total)) / len(present)
    if upper == "MIN":
        return min(present, key=_sort_key)  # pyright: ignore[reportArgumentType]
    if upper == "MAX":
        return max(present, key=_sort_key)  # pyright: ignore[reportArgumentType]
    if upper == "FIRST":
        return present[0]
    if upper == "LAST":
        return present[-1]
    if upper in ("STDEV", "STDEVP", "VAR", "VARP"):
        numbers = [float(_number(v)) for v in present]
        divisor = len(numbers) - 1 if upper in ("STDEV", "VAR") else len(numbers)
        if divisor <= 0:
            return None
        mean = sum(numbers) / len(numbers)
        variance = sum((n - mean) ** 2 for n in numbers) / divisor
        return variance if upper.startswith("VAR") else variance**0.5
    raise AccessError(f"aggregate {name} is not available")


def _sort_key(value: object) -> tuple[int, object]:
    if value is None:
        return (0, 0)
    if isinstance(value, str):
        return (1, value.lower())
    if isinstance(value, bool):
        return (1, -1 if value else 0)
    if isinstance(value, _dt.datetime):
        return (1, _number(value))
    if isinstance(value, Decimal):
        return (1, float(value))
    if isinstance(value, bytes):
        return (1, value)
    return (1, value)


def _evaluate_with_aggregates(
    expr: Expr,
    group: list[Row],
    sources: Sequence[Source],
    parameters: Mapping[str, object],
    env_for: Callable[[Row], Environment] | None = None,
) -> object:
    """Evaluate an expression over a group: aggregate calls collapse the
    group, everything else is read off its first row."""

    def make(row: Row) -> Environment:
        return env_for(row) if env_for is not None else _environment(row, sources, parameters)

    def rewrite(node: Expr) -> Expr:
        if isinstance(node, Call) and node.is_aggregate:
            if node.star:
                return Literal(len(group))
            if len(node.args) != 1:
                raise AccessError(f"{node.name} takes one argument")
            values = [node.args[0].eval(make(r)) for r in group]
            return Literal(
                _aggregate(node.name, values, money=_money_column(node.args[0], sources))
            )
        if isinstance(node, Binary):
            return Binary(node.op, rewrite(node.left), rewrite(node.right))
        if isinstance(node, Unary):
            return Unary(node.op, rewrite(node.operand))
        if isinstance(node, Call):
            return Call(node.name, tuple(rewrite(a) for a in node.args))
        if isinstance(node, InList):
            return InList(rewrite(node.operand), tuple(rewrite(o) for o in node.options), node.negate)
        if isinstance(node, Between):
            return Between(rewrite(node.operand), rewrite(node.low), rewrite(node.high), node.negate)
        return node

    return rewrite(expr).eval(make(group[0] if group else {}))


def _has_aggregate(expr: Expr) -> bool:
    if isinstance(expr, Call):
        return expr.is_aggregate or any(_has_aggregate(a) for a in expr.args)
    if isinstance(expr, Binary):
        return _has_aggregate(expr.left) or _has_aggregate(expr.right)
    if isinstance(expr, Unary):
        return _has_aggregate(expr.operand)
    if isinstance(expr, InList):
        return _has_aggregate(expr.operand) or any(_has_aggregate(o) for o in expr.options)
    if isinstance(expr, Between):
        return any(_has_aggregate(e) for e in (expr.operand, expr.low, expr.high))
    return False


def _sources(db: AccessDatabase, from_clause: str, parameters: Mapping[str, object], outer: Environment | None = None) -> list[Source]:
    tables, joins = parse_from(from_clause)
    sources: list[Source] = []
    for t in tables:
        name = (t.name1 or (f"({t.expression})" if t.expression else "")).strip()
        alias = (t.name2 or name).strip("[]")
        if name.startswith("("):
            inner = name[1 : name.rfind(")")].strip()
            rows = _run(db, inner, parameters, outer)
            sources.append(_rows_source(alias, list(rows[0]) if rows else [], rows))
            continue
        name = name.strip("[]")
        table = db.table(name) if any(n.lower() == name.lower() for n in db.table_names()) else None
        if table is not None:
            sources.append(_table_source(table, alias))
            continue
        saved = next((q for q in db.queries() if q.name.lower() == name.lower()), None)
        if saved is None:
            raise AccessError(f"no table or query named {name!r}")
        rows = _run(db, saved.sql, parameters, outer)
        sources.append(_rows_source(alias, list(rows[0]) if rows else [], rows))
    for j in joins:
        target = (j.name2 or "").strip("[]").lower()
        source = next((s for s in sources[1:] if s.alias.lower() == target), None)
        if source is None:
            raise AccessError(f"join names {j.name2!r}, which is not a table of the query")
        source.join = {1: "INNER", 2: "LEFT", 3: "RIGHT"}.get(j.flag or 1, "INNER")
        source.condition = Parser.parse(j.expression or "")
    if len(sources) > 1 and not joins:
        for s in sources[1:]:
            s.join = "CROSS"
    return sources


def execute(
    db: AccessDatabase,
    sql: str,
    parameters: Mapping[str, object] | None = None,
    *,
    created: object | None = None,
    updated: object | None = None,
    referenced_updated: object | None = None,
    owner_updated: object | None = None,
) -> list[Row] | int:
    """Run one statement.  SELECT returns its rows as dicts keyed by the
    output column names (aliases, column names, or ``Expr1000``...);
    INSERT, UPDATE and DELETE return the number of rows affected, and DDL
    returns 0 as DAO does.  ``created`` and ``updated`` are the catalog
    timestamps a DDL statement stamps, for reproducing a database the
    engine wrote."""
    parameters = dict(parameters or {})
    text = sql.strip().rstrip(";").strip()
    if is_ddl(text):
        return execute_ddl(db, text, created=created, updated=updated, referenced_updated=referenced_updated, owner_updated=owner_updated)
    if text.upper().startswith("PARAMETERS "):
        _, _, text = text.partition(";")
        text = text.strip()
    members = union_members(text)
    if members is not None:
        return _union(db, members, parameters)
    clauses = split_clauses(text)
    verb = clauses[0][0]
    if verb == "TRANSFORM":
        return _crosstab(db, clauses, parameters)
    if verb == "SELECT":
        if "INTO" in dict(clauses):
            with db.transaction():
                return _make_table(
                    db, clauses, parameters, created=created, updated=updated, owner_updated=owner_updated
                )
        return _select(db, clauses, parameters)
    # An action query is all or nothing: one that fails on its third row
    # leaves the first two unwritten, which is what the engine does.
    if verb in ("INSERT INTO", "DELETE"):
        progress = _Progress()
        try:
            with db.transaction():
                if verb == "INSERT INTO":
                    return _insert(db, clauses, parameters, progress)
                return _delete(db, clauses, parameters, progress)
        except AccessError:
            progress.keep(db)
            raise
    if verb == "UPDATE":
        with db.transaction():
            return _update(db, clauses, parameters)
    raise AccessError(f"statement {verb} is not supported")


def _run(db: AccessDatabase, sql: str, parameters: Mapping[str, object], outer: Environment | None = None) -> list[Row]:
    """Run a nested SELECT: a subquery, a derived table or a saved query."""
    text = sql.strip().rstrip(";").strip()
    if text.upper().startswith("PARAMETERS "):
        _, _, text = text.partition(";")
        text = text.strip()
    members = union_members(text)
    if members is not None:
        return _union(db, members, parameters, outer)
    clauses = split_clauses(text)
    if clauses[0][0] == "TRANSFORM":
        return _crosstab(db, clauses, parameters, outer)
    return _select(db, clauses, parameters, outer)


def _crosstab(
    db: AccessDatabase,
    clauses: list[tuple[str, str]],
    parameters: Mapping[str, object],
    outer: Environment | None = None,
) -> list[Row]:
    """Run ``TRANSFORM value SELECT headings ... PIVOT column [IN (...)]``:
    the SELECT list gives one row per group, the pivot values give the
    columns after them, and the TRANSFORM aggregate fills the cells.  A
    cell with no rows behind it is Null, and an ``IN`` list fixes the
    columns and their order (empty ones included)."""
    by_word = dict(clauses)
    pivot_clause = by_word.get("PIVOT", "").strip()
    if not pivot_clause:
        raise AccessError("a TRANSFORM query needs a PIVOT clause")
    if "HAVING" in by_word:
        raise AccessError("a crosstab query takes no HAVING clause")
    value_text, _value_alias = split_alias(clauses[0][1].strip())
    value = Parser.parse(value_text)
    if not _has_aggregate(value):
        raise AccessError("TRANSFORM needs an aggregate, such as Sum or Count")
    pivot_text, in_list = _split_pivot(pivot_clause)
    pivot = Parser.parse(pivot_text)

    inner = [("SELECT", by_word.get("SELECT", "")), ("FROM", by_word.get("FROM", ""))]
    if "WHERE" in by_word:
        inner.append(("WHERE", by_word["WHERE"]))
    sources = _sources(db, by_word.get("FROM", ""), parameters, outer)

    def runner(sql: str, env: Environment) -> list[Row]:
        return _run(db, sql, parameters, env)

    def env_for(row: Row) -> Environment:
        return _environment(row, sources, parameters, outer, runner)

    rows = _join(sources, parameters, outer)
    if "WHERE" in by_word:
        where = Parser.parse(by_word["WHERE"])
        rows = [r for r in rows if (v := where.eval(env_for(r))) is not None and _truthy(v)]

    _flags, _top, headings = select_list(by_word.get("SELECT", ""))
    heading_names: list[str] = []
    heading_exprs: list[Expr] = []
    counter = 1000
    for expression, alias in headings:
        expr = Parser.parse(expression)
        heading_exprs.append(expr)
        if alias:
            heading_names.append(alias.strip("[]"))
        elif isinstance(expr, ColumnRef):
            heading_names.append(expr.name)
        else:
            heading_names.append(f"Expr{counter}")
            counter += 1

    # Rows group by the GROUP BY clause when there is one, else by the
    # headings that are not aggregates; an aggregate heading is worked out
    # over the group, like the transformed value itself.
    if "GROUP BY" in by_word:
        key_exprs = [Parser.parse(item.strip()) for item in split_top_level(by_word["GROUP BY"], ",")]
    else:
        key_exprs = [e for e in heading_exprs if not _has_aggregate(e)]
    groups: dict[tuple[object, ...], list[Row]] = {}
    columns: list[object] = list(in_list) if in_list is not None else []
    cells: dict[tuple[tuple[object, ...], object], list[Row]] = {}
    for row in rows:
        env = env_for(row)
        key = tuple(_hashable(e.eval(env)) for e in key_exprs)
        groups.setdefault(key, []).append(row)
        column = pivot.eval(env)
        if in_list is None and not any(_same(column, c) for c in columns):
            columns.append(column)
        cells.setdefault((key, _hashable(column)), []).append(row)

    if in_list is None:
        columns.sort(key=_sort_key)
    out: list[Row] = []
    for key, group in groups.items():
        record: Row = {}
        for name, expr in zip(heading_names, heading_exprs, strict=True):
            heading = (
                _evaluate_with_aggregates(expr, group, sources, parameters, env_for)
                if _has_aggregate(expr)
                else expr.eval(env_for(group[0]))
            )
            record[name] = heading if isinstance(expr, ColumnRef) else _computed(heading)
        for column in columns:
            behind = cells.get((key, _hashable(column)), [])
            record[_column_name(column)] = (
                _evaluate_with_aggregates(value, behind, sources, parameters, env_for) if behind else None
            )
        out.append(record)
    # Without an ORDER BY the engine hands back a crosstab's rows sorted by
    # their headings (measured: True before False on a Boolean heading).
    for name in reversed(heading_names):
        out.sort(key=lambda record, name=name: _sort_key(record.get(name)))
    if "ORDER BY" in by_word:
        for item in reversed(split_top_level(by_word["ORDER BY"], ",")):
            item = item.strip()
            descending = item.upper().endswith(" DESC")
            if descending or item.upper().endswith(" ASC"):
                item = item.rsplit(" ", 1)[0].strip()
            expr = Parser.parse(item)
            out.sort(
                key=lambda record, expr=expr: _sort_key(
                    _ordinal(expr, record)
                    if _ordinal(expr, record) is not None
                    else (record[expr.name] if isinstance(expr, ColumnRef) and expr.name in record else None)
                ),
                reverse=descending,
            )
    return out


def _same(a: object, b: object) -> bool:
    return _hashable(a) == _hashable(b)


def _column_name(value: object) -> str:
    """What a pivot value is called as a column: its text, and ``<>`` for
    Null, which is how the engine shows a crosstab's null heading.  A
    truth value is written the way the engine writes one, so pivoting on
    a comparison gives columns named -1 and 0."""
    if value is None:
        return "<>"
    return _text(_computed(value))


def _split_pivot(clause: str) -> tuple[str, list[object] | None]:
    """The pivot expression and the values an ``IN`` list fixes."""
    parts = split_top_level_words(clause, "IN")
    if len(parts) < 2:
        return clause.strip(), None
    head = parts[0].strip()
    rest = "IN".join(parts[1:]).strip() if len(parts) > 2 else parts[1].strip()
    inner = rest[rest.find("(") + 1 : rest.rfind(")")]
    env = _environment({}, [], {})
    return head, [Parser.parse(item.strip()).eval(env) for item in split_top_level(inner, ",")]



def union_members(text: str) -> list[tuple[str, str]] | None:
    """A top-level union as ``(operator, member)`` pairs, the first
    operator empty, or None when the statement is not a union."""
    upper = text.upper()
    parts: list[tuple[str, str]] = []
    depth = 0
    quote: str | None = None
    start = 0
    operator = ""
    i = 0
    while i < len(text):
        ch = text[i]
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == "[":
            quote = "]"
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif (
            depth == 0
            and upper.startswith("UNION", i)
            and (i == 0 or text[i - 1].isspace())
            and (i + 5 == len(text) or not (text[i + 5].isalnum() or text[i + 5] == "_"))
        ):
            parts.append((operator, text[start:i].strip()))
            rest = text[i + 5 :]
            stripped = rest.lstrip()
            keep_all = stripped[:3].upper() == "ALL" and (len(stripped) == 3 or not stripped[3].isalnum())
            operator = "UNION ALL" if keep_all else "UNION"
            i = i + 5 + (len(rest) - len(stripped)) + (3 if keep_all else 0)
            start = i
            continue
        i += 1
    if not parts:
        return None
    parts.append((operator, text[start:].strip()))
    return parts


def _union(db: AccessDatabase, members: list[tuple[str, str]], parameters: Mapping[str, object], outer: Environment | None = None) -> list[Row]:
    """Run a union left to right, under the first member's column names.
    ``UNION`` makes what it has so far distinct; ``UNION ALL`` keeps every
    row.  A trailing ORDER BY belongs to the union, not its last member."""
    out: list[Row] = []
    names: list[str] = []
    order: str | None = None
    for index, (operator, member) in enumerate(members):
        text = member
        if index == len(members) - 1:
            head, sep, tail = _split_trailing_order(text)
            if sep:
                text, order = head, tail
        rows = _run(db, text, parameters, outer)
        if index == 0 and rows:
            names = list(rows[0])
        for row in rows:
            out.append(dict(zip(names, row.values(), strict=False)) if names else dict(row))
        if operator == "UNION":
            seen: set[tuple[object, ...]] = set()
            distinct: list[Row] = []
            for row in out:
                key = tuple(_hashable(v) for v in row.values())
                if key not in seen:
                    seen.add(key)
                    distinct.append(row)
            out = distinct
    if order is not None:
        for item in reversed(split_top_level(order, ",")):
            item = item.strip()
            descending = item.upper().endswith(" DESC")
            if descending or item.upper().endswith(" ASC"):
                item = item.rsplit(" ", 1)[0].strip()
            column = item.strip("[]")
            out.sort(key=lambda row, column=column: _sort_key(row.get(column)), reverse=descending)
    return out


def _split_trailing_order(text: str) -> tuple[str, str, str]:
    """A member and its trailing ORDER BY, which belongs to the union."""
    clauses = split_clauses(text)
    for word, body in clauses:
        if word == "ORDER BY":
            head = text[: text.upper().rindex("ORDER BY")].rstrip()
            return head, "ORDER BY", body
    return text, "", ""


def _select(
    db: AccessDatabase,
    clauses: list[tuple[str, str]],
    parameters: Mapping[str, object],
    outer: Environment | None = None,
) -> list[Row]:
    by_word = dict(clauses)
    if "FROM" not in by_word:
        raise AccessError("SELECT needs a FROM clause")
    flags, top, items = select_list(clauses[0][1])
    sources = _sources(db, by_word["FROM"], parameters, outer)

    def runner(sql: str, env: Environment) -> list[Row]:
        return _run(db, sql, parameters, env)

    def env_for(row: Row) -> Environment:
        return _environment(row, sources, parameters, outer, runner)

    rows = _join(sources, parameters, outer)
    if "WHERE" in by_word:
        where = Parser.parse(by_word["WHERE"])
        rows = [r for r in rows if (v := where.eval(env_for(r))) is not None and _truthy(v)]
    outputs = _outputs(flags, items, sources)
    grouped = "GROUP BY" in by_word or any(o[1] is not None and _has_aggregate(o[1]) for o in outputs)
    groups: list[list[Row]]
    if grouped:
        keys = [Parser.parse(g.strip()) for g in split_top_level(by_word.get("GROUP BY", ""), ",")] if "GROUP BY" in by_word else []
        buckets: dict[tuple[object, ...], tuple[list[Row], tuple[object, ...]]] = {}
        for r in rows:
            env = env_for(r)
            values = tuple(k.eval(env) for k in keys)
            key = tuple(_hashable(v) for v in values)
            buckets.setdefault(key, ([], values))[0].append(r)
        # The engine groups by sorting, so the groups come out in key
        # order -- Null first, then ascending, text case-blind -- whatever
        # order the rows came in; a group shows the first value it met
        # (measured).
        ordered = sorted(buckets.values(), key=lambda pair: tuple(_sort_key(v) for v in pair[1]))
        groups = [pair[0] for pair in ordered] if buckets or keys else [rows]
        if "HAVING" in by_word:
            having = Parser.parse(by_word["HAVING"])
            groups = [g for g in groups if (v := _evaluate_with_aggregates(having, g, sources, parameters, env_for)) is not None and _truthy(v)]
    else:
        groups = [[r] for r in rows]
    result: list[tuple[Row, Row]] = []
    for group in groups:
        out: Row = {}
        for name, expr, star in outputs:
            if star is not None:
                source = next(s for s in sources if s.alias.lower() == star.lower())
                first = group[0] if group else {}
                for column in source.columns:
                    out[column] = first.get(f"{star.lower()}.{column.lower()}")
            else:
                value = (
                    _evaluate_with_aggregates(expr, group, sources, parameters, env_for)  # pyright: ignore[reportArgumentType]
                    if grouped
                    else expr.eval(env_for(group[0]))  # pyright: ignore[reportOptionalMemberAccess]
                )
                # A column read straight out keeps its type; anything
                # computed comes back the way Jet writes a truth value.
                out[name] = value if isinstance(expr, ColumnRef) else _computed(value)
        result.append((out, group[0] if group else {}))
    if flags & 0x02:
        # DISTINCT is done by sorting too: the rows that survive come out
        # in the order of their values, first-seen row for each (measured),
        # and an ORDER BY then sorts those.
        seen: set[tuple[object, ...]] = set()
        unique: list[tuple[Row, Row]] = []
        for pair in result:
            key = tuple(_hashable(v) for v in pair[0].values())
            if key not in seen:
                seen.add(key)
                unique.append(pair)
        unique.sort(key=lambda pair: tuple(_sort_key(v) for v in pair[0].values()))
        result = unique
    if "ORDER BY" in by_word:
        orderings: list[tuple[Expr, bool]] = []
        for item in split_top_level(by_word["ORDER BY"], ","):
            item = item.strip()
            descending = False
            if item.upper().endswith(" DESC"):
                item, descending = item[:-5].rstrip(), True
            elif item.upper().endswith(" ASC"):
                item = item[:-4].rstrip()
            orderings.append((Parser.parse(item), descending))
        for expr, descending in reversed(orderings):
            result.sort(key=lambda pair, expr=expr: _order_key(pair, expr, grouped, sources, parameters, env_for), reverse=descending)
    output = [pair[0] for pair in result]
    if top is not None:
        count, _, percent = top.partition(" ")
        limit = -(-len(output) * int(count) // 100) if percent.strip().upper() == "PERCENT" else int(count)
        output = output[:limit]
    return output


#: An output column: its name, its expression, and -- for a `t.*` --
#: the alias whose every column it stands for instead.
Output = tuple[str, "Expr | None", "str | None"]


def _outputs(flags: int, items: list[tuple[str, str | None]], sources: Sequence[Source]) -> list[Output]:
    """The select list as output columns, named as the engine names them:
    the alias, the column, or ``Expr`` and the position."""
    outputs: list[Output] = []
    _reject_duplicate_aliases(items)
    if flags & 0x01 or not items:
        for s in sources:
            outputs.append((s.alias, None, s.alias))
    for position, (expression, alias) in enumerate(items):
        if expression.endswith("*"):
            qualifier = expression[:-1].rstrip(".")
            outputs.append((qualifier, None, qualifier or sources[0].alias))
            continue
        expr = Parser.parse(expression)
        if alias:
            name = alias.strip("[]")
        elif isinstance(expr, ColumnRef):
            name = expr.name
        else:
            name = _expr_name(position)
        outputs.append((name, expr, None))
    _qualify_repeats(outputs, items)
    _number_repeats(outputs, items)
    return outputs


#: What the engine calls an output column with no name of its own.
EXPR_NAME_BASE = 1000


def _expr_name(position: int) -> str:
    """`Expr` and 1000 plus where the column sits in the select list.

    It is the position, not a count of the ones before it that needed a
    name: `SELECT Id, Id + 1` answers with `Id` and `Expr1001`.
    """
    return f"Expr{EXPR_NAME_BASE + position}"


def _reject_duplicate_aliases(items: list[tuple[str, str | None]]) -> None:
    """Two output columns cannot be given the same name, and the engine
    says so rather than dropping one."""
    seen: set[str] = set()
    for _expression, alias in items:
        if alias is None:
            continue
        key = alias.strip("[]").lower()
        if key in seen:
            raise AccessError(f"duplicate output alias {alias.strip('[]')!r}")
        seen.add(key)


def _number_repeats(
    outputs: list[tuple[str, Expr | None, str | None]], items: list[tuple[str, str | None]]
) -> None:
    """A name the select list holds more than once is kept by its last
    column only; the ones before it are named for their position.

    `SELECT Id, Id` answers with `Expr1000` and `Id`, and
    `SELECT Total, Total, Total` with `Expr1000`, `Expr1001` and `Total`
    (measured against the engine).  Without this the columns collapse into
    one and the row comes back a column short.
    """
    counts: dict[str, int] = {}
    for name, _expr, star in outputs:
        if star is None:
            counts[name.lower()] = counts.get(name.lower(), 0) + 1
    if all(n < 2 for n in counts.values()):
        return
    last: dict[str, int] = {}
    for i, (name, _expr, star) in enumerate(outputs):
        if star is None:
            last[name.lower()] = i
    written = 0
    for i, (name, expr, star) in enumerate(outputs):
        if star is not None:
            continue
        position = written
        written += 1
        if counts.get(name.lower(), 0) > 1 and last[name.lower()] != i:
            outputs[i] = (_expr_name(position), expr, star)


def _qualify_repeats(outputs: list[tuple[str, Expr | None, str | None]], items: list[tuple[str, str | None]]) -> None:
    """Two sources can hold the same column name.  The engine then names
    every one of them for its table -- ``a.Id`` and ``b.Id`` -- and only
    then, which is what this puts back into the output list."""
    counts: dict[str, int] = {}
    for name, _expr, star in outputs:
        if star is None:
            counts[name.lower()] = counts.get(name.lower(), 0) + 1
    written = 0
    for i, (name, expr, star) in enumerate(outputs):
        if star is not None:
            continue
        alias = items[written][1] if written < len(items) else None
        written += 1
        if counts.get(name.lower(), 0) > 1 and alias is None and isinstance(expr, ColumnRef) and expr.qualifier:
            outputs[i] = (f"{expr.qualifier}.{expr.name}", expr, star)


def _order_key(
    pair: tuple[Row, Row],
    expr: Expr,
    grouped: bool,
    sources: Sequence[Source],
    parameters: Mapping[str, object],
    env_for: Callable[[Row], Environment],
) -> tuple[int, object]:
    out, base = pair
    ordinal = _ordinal(expr, out)
    if ordinal is not None:
        return _sort_key(ordinal)
    if isinstance(expr, ColumnRef) and expr.qualifier is None and expr.name in out:
        return _sort_key(out[expr.name])
    if grouped:
        return _sort_key(_evaluate_with_aggregates(expr, [base], sources, parameters, env_for))
    return _sort_key(expr.eval(env_for(base)))


def _ordinal(expr: Expr, row: Row) -> object | None:
    """``ORDER BY 2`` names the second output column, not the number 2."""
    if not isinstance(expr, Literal) or isinstance(expr.value, bool) or not isinstance(expr.value, int):
        return None
    names = list(row)
    if not 1 <= expr.value <= len(names):
        raise AccessError(f"ORDER BY {expr.value} names no column")
    return row[names[expr.value - 1]]


def _hashable(value: object) -> object:
    if isinstance(value, str):
        return value.lower()
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, Decimal):
        return float(value)
    return value


# --- SELECT ... INTO --------------------------------------------------------------

#: The engine's Decimal, as a make-table query creates one from an
#: expression: 28 digits, the scale the expression works out to.
DECIMAL_PRECISION = 28

#: Functions whose answer is text; the column is a Text(255).
_TEXT_FUNCTIONS = frozenset(
    "CSTR STR HEX OCT CHR SPACE STRING UCASE LCASE TRIM LTRIM RTRIM LEFT RIGHT MID "
    "FORMAT REPLACE STRREVERSE WEEKDAYNAME MONTHNAME PARTITION".split()
)
#: Functions whose answer is a Long.
_LONG_FUNCTIONS = frozenset("LEN INSTR DATEDIFF DATEPART CLNG".split())
#: Functions whose answer is an Integer -- the yes/no ones included, since
#: the engine has no Boolean outside a column.
_INTEGER_FUNCTIONS = frozenset(
    "ASC STRCOMP SGN CINT CBOOL CBYTE YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY "
    "ISNULL ISNUMERIC ISDATE".split()
)
#: Functions whose answer is a Double, whatever they were given.
_DOUBLE_FUNCTIONS = frozenset("ABS ROUND SQR EXP LOG VAL CDBL CSNG".split())
#: Functions whose answer is a date.
_DATE_FUNCTIONS = frozenset(
    "CDATE DATEVALUE TIMEVALUE DATESERIAL TIMESERIAL DATEADD NOW DATE TIME".split()
)


def _computed_column(name: str, kind: str, size: int | tuple[int, int] | None = None) -> ColumnSpec:
    """The column an expression becomes.  The engine keeps it among the
    variable-length columns rather than in the fixed block a Long or a
    Double would normally take -- every type but Integer, which stays
    fixed (measured on every expression type) -- so the row layout of a
    made table differs from the same columns declared by CREATE TABLE."""
    return ColumnSpec(name, kind, size, variable=kind != "integer")


def _text_column(name: str) -> ColumnSpec:
    """The column an expression that answers text becomes: 255 characters,
    without Unicode compression (measured)."""
    return ColumnSpec(name, "text", 255, compressed=False, variable=True)


def _decimal_column(name: str, scale: int) -> ColumnSpec:
    return _computed_column(name, "decimal", (DECIMAL_PRECISION, scale))


def _kind(spec: ColumnSpec) -> str:
    return spec.type.lower()


def _scale(spec: ColumnSpec) -> int:
    return spec.size[1] if _kind(spec) == "decimal" and isinstance(spec.size, tuple) else 0


def _promote(name: str, left: ColumnSpec, right: ColumnSpec, op: str) -> ColumnSpec:
    """The column `left op right` makes, for the arithmetic operators.

    Measured: the whole-number types come together as a Long; a Single
    with anything is a Double; Currency with a Long or a Double stays
    Currency; a Decimal keeps its digits, adding scales under `*` and
    taking the wider under `+` and `-`; a Date plus a number is a Date and
    a Date less a Date is a Double."""
    kinds = {_kind(left), _kind(right)}
    if kinds & {"text", "memo"}:
        return _text_column(name)
    if "datetime" in kinds:
        return _computed_column(name, "double") if kinds == {"datetime"} else _computed_column(name, "datetime")
    if "decimal" in kinds:
        scales = (_scale(left), _scale(right))
        return _decimal_column(name, sum(scales) if op == "*" else max(scales))
    if "currency" in kinds:
        return _computed_column(name, "currency")
    if kinds & {"double", "single"}:
        return _computed_column(name, "double")
    return _computed_column(name, "long")


def _widen(name: str, specs: Sequence[ColumnSpec]) -> ColumnSpec:
    """One column for a set of alternatives -- the branches of IIf, the
    answers of Switch -- with Null ignored, text winning, and the numbers
    promoted as `+` promotes them (measured: IIf(b, L, D) is a Double,
    IIf(b, 1, 'x') is Text)."""
    live = [s for s in specs if _kind(s) != "binary"]
    if not live:
        return _computed_column(name, "binary", 510)
    if any(_kind(s) in ("text", "memo") for s in live):
        return _text_column(name)
    spec = _computed_column(name, live[0].type, live[0].size)
    for other in live[1:]:
        spec = _promote(name, spec, other, "+")
    return spec


def _column_copy(name: str, table: Table, column: str) -> ColumnSpec:
    """A column read straight out of a table keeps its definition: type,
    size, AutoNumber and Unicode compression alike, though none of its
    properties (measured: Required, DefaultValue and the validation
    fields are all dropped)."""
    definition = table.definition.column(column)
    size: int | tuple[int, int] | None = None
    if definition.type_code == TYPE_TEXT:
        size = definition.length // 2
    elif definition.type_code == TYPE_BINARY:
        size = definition.length
    elif definition.type_code == TYPE_NUMERIC:
        size = (definition.precision, definition.scale)
    return ColumnSpec(
        name,
        definition.type_name,
        size,
        autonumber=definition.auto_number,
        compressed=definition.compressed_unicode,
        variable=not definition.is_fixed and definition.type_code in FIXED_LENGTH_TYPES,
    )


def _value_column(name: str, value: object) -> ColumnSpec:
    """A column for a value whose expression cannot be typed -- a column of
    a derived table -- taken from the value itself."""
    if isinstance(value, bool):
        return _computed_column(name, "integer")
    if isinstance(value, int):
        return _computed_column(name, "long")
    if isinstance(value, float):
        return _computed_column(name, "double")
    if isinstance(value, Decimal):
        return _computed_column(name, "currency")
    if isinstance(value, _dt.datetime):
        return _computed_column(name, "datetime")
    if isinstance(value, (bytes, bytearray)):
        return _computed_column(name, "ole")
    return _text_column(name)


def _static_type(
    name: str, expr: Expr, sources: Sequence[Source], db: AccessDatabase, parameters: Mapping[str, object]
) -> ColumnSpec:
    """The column a make-table query creates for an expression.  The
    engine decides this from the expression alone -- the same column comes
    out of a query over no rows -- so this walks the expression rather
    than looking at values (measured over 120 expressions, in
    ``docs/access_engine.md``)."""
    if isinstance(expr, ColumnRef):
        source = _column_source(expr, sources)
        if source.table is not None:
            return _column_copy(name, source.table, expr.name)
        # A column of a saved query or a bracketed SELECT: type it from
        # the first value it holds, the expression being out of reach.
        key = f"{source.alias}.{expr.name}".lower()
        value = next((r[key] for r in source.rows if r.get(key) is not None), None)
        return _value_column(name, value)
    if isinstance(expr, Literal):
        value = expr.value
        if value is None:
            # A Null has no type; the engine gives it a Binary as wide as a
            # Text(255) (measured).
            return _computed_column(name, "binary", 510)
        if isinstance(value, bool):
            return _computed_column(name, "integer")
        if isinstance(value, int):
            return _computed_column(name, "long") if -(1 << 31) <= value < 1 << 31 else _decimal_column(name, 0)
        if isinstance(value, float):
            # A number with a fraction is a Decimal, scaled to the digits
            # it was written with: `5.5` makes a Decimal(28, 1).
            exponent = Decimal(repr(value)).normalize().as_tuple().exponent
            return _decimal_column(name, -exponent if isinstance(exponent, int) and exponent < 0 else 0)
        if isinstance(value, _dt.datetime):
            return _computed_column(name, "datetime")
        return _text_column(name)
    if isinstance(expr, Unary):
        if expr.op == "NOT":
            return _computed_column(name, "integer")
        return _promote(name, _static_type(name, expr.operand, sources, db, parameters), _computed_column(name, "long"), "-")
    if isinstance(expr, Binary):
        op = expr.op
        if op in ("=", "<>", "<", ">", "<=", ">=", "LIKE", "IS", "IS NOT", "AND", "OR"):
            return _computed_column(name, "integer")
        if op == "&":
            return _text_column(name)
        if op in ("/", "^"):
            return _computed_column(name, "double")
        if op in ("\\", "MOD"):
            return _computed_column(name, "long")
        left = _static_type(name, expr.left, sources, db, parameters)
        right = _static_type(name, expr.right, sources, db, parameters)
        return _promote(name, left, right, op)
    if isinstance(expr, (InList, Between)):
        return _computed_column(name, "integer")
    if isinstance(expr, SubQuery):
        if expr.kind != "scalar":
            return _computed_column(name, "integer")
        return _subquery_type(name, expr.sql, sources, db, parameters)
    if isinstance(expr, Call):
        return _call_type(name, expr, sources, db, parameters)
    raise AccessError(f"cannot work out the column type of {expr!r}")


def _column_source(ref: ColumnRef, sources: Sequence[Source]) -> Source:
    if ref.qualifier is not None:
        source = next((s for s in sources if s.alias.lower() == ref.qualifier.lower()), None)
        if source is None:
            raise AccessError(f"{ref.qualifier!r} is not a table of the query")
        return source
    holders = [s for s in sources if s.holds(ref.name)]
    if not holders:
        raise AccessError(f"no column named {ref.name!r} in any table of the query")
    return holders[0]


def _subquery_type(
    name: str, sql: str, sources: Sequence[Source], db: AccessDatabase, parameters: Mapping[str, object]
) -> ColumnSpec:
    """A scalar subquery is typed as its first output column would be."""
    text = sql.strip().rstrip(";").strip()
    clauses = split_clauses(text)
    if clauses[0][0] != "SELECT" or "FROM" not in dict(clauses):
        return _computed_column(name, "double")
    flags, _top, items = select_list(clauses[0][1])
    inner = _sources(db, dict(clauses)["FROM"], parameters)
    outputs = _outputs(flags, items, list(inner) + list(sources))
    if not outputs:
        return _computed_column(name, "double")
    first_name, first_expr, star = outputs[0]
    if star is not None or first_expr is None:
        source = next(s for s in inner if s.alias.lower() == (star or first_name).lower())
        if source.table is None or not source.columns:
            return _computed_column(name, "double")
        spec = _column_copy(name, source.table, source.columns[0])
    else:
        spec = _static_type(name, first_expr, list(inner) + list(sources), db, parameters)
    # The subquery's value is computed as far as the outer query knows:
    # text comes back at the full width, and nothing stays fixed-length.
    if _kind(spec) in ("text", "memo"):
        return _text_column(name)
    return _computed_column(name, spec.type, spec.size)


def _call_type(
    name: str, call: Call, sources: Sequence[Source], db: AccessDatabase, parameters: Mapping[str, object]
) -> ColumnSpec:
    upper = call.name.upper()

    def arg(index: int) -> ColumnSpec:
        return _static_type(name, call.args[index], sources, db, parameters)

    if upper in AGGREGATES:
        if upper == "COUNT":
            return _computed_column(name, "long")
        if upper in ("STDEV", "STDEVP", "VAR", "VARP"):
            return _computed_column(name, "double")
        operand = arg(0)
        kind = _kind(operand)
        if upper in ("SUM", "AVG"):
            # A sum or average of whole numbers is a Double; Currency and
            # Decimal keep their type (measured).
            if kind not in ("currency", "decimal"):
                return _computed_column(name, "double")
        # Min, Max, First and Last answer in the column's own type, except
        # that text comes back at the full 255 characters -- and like any
        # computed column, none of them stays in the fixed block.
        if kind in ("text", "memo"):
            return _text_column(name)
        return _computed_column(name, operand.type, operand.size)
    if upper in DOMAIN_FUNCTIONS:
        aggregate = DOMAIN_FUNCTIONS[upper]
        if aggregate is None or aggregate in ("First", "Last"):
            return _text_column(name)
        return _computed_column(name, "long") if aggregate == "Count" else _computed_column(name, "double")
    if upper in _TEXT_FUNCTIONS:
        return _text_column(name)
    if upper in _LONG_FUNCTIONS:
        return _computed_column(name, "long")
    if upper in _INTEGER_FUNCTIONS:
        return _computed_column(name, "integer")
    if upper in _DOUBLE_FUNCTIONS:
        return _computed_column(name, "double")
    if upper in _DATE_FUNCTIONS:
        return _computed_column(name, "datetime")
    if upper == "CCUR":
        return _computed_column(name, "currency")
    if upper in ("INT", "FIX"):
        # Int keeps a whole-number type (as a Long) and leaves a Double a
        # Double (measured).
        return _promote(name, arg(0), _computed_column(name, "long"), "+")
    if upper == "IIF":
        return _widen(name, [arg(1), arg(2)]) if len(call.args) == 3 else _computed_column(name, "double")
    if upper == "NZ":
        return _widen(name, [arg(i) for i in range(len(call.args))])
    if upper == "SWITCH":
        return _widen(name, [arg(i) for i in range(1, len(call.args), 2)])
    if upper == "CHOOSE":
        return _widen(name, [arg(i) for i in range(1, len(call.args))])
    return _computed_column(name, "double")


def _made_headers(specs: Sequence[ColumnSpec]) -> list[ColumnSpec]:
    """The two things a made table's column headers carry that CREATE
    TABLE's do not (measured, byte for byte against the engine):

    * bytes 9-10 count the columns from one rather than from zero;
    * every column after a Decimal -- copied or computed, fixed or not,
      Memo and GUID included, Text alone excepted -- carries that
      Decimal's precision and scale where its collation would be, until
      the next Decimal resets them.  A stale buffer in the engine, by the
      look of it, and reproduced as such."""
    out: list[ColumnSpec] = []
    leak: tuple[int, int] | None = None
    for position, spec in enumerate(specs):
        collation = spec.collation
        kind = _kind(spec)
        if kind == "decimal":
            leak = spec.size if isinstance(spec.size, tuple) else (18, 0)
        elif leak is not None and kind != "text":
            collation = leak
        out.append(replace(spec, ordinal=position + 1, collation=collation))
    return out


def _make_table(
    db: AccessDatabase,
    clauses: list[tuple[str, str]],
    parameters: Mapping[str, object],
    *,
    created: object | None = None,
    updated: object | None = None,
    owner_updated: object | None = None,
) -> int:
    """Run ``SELECT ... INTO NewTable FROM ...``: create the table the
    query describes and fill it with the query's rows, answering how many.

    The columns come from the select list -- a column read straight out
    keeps its definition, an expression gets the type the engine gives
    it -- and none of the source's indexes or column properties come
    along (measured).  An existing name is an error.  The three stamps
    are the new table's catalog timestamps, as for any DDL."""
    by_word = dict(clauses)
    target = by_word["INTO"].strip().strip("[]")
    if any(n.lower() == target.lower() for n in db.table_names()):
        raise AccessError(f"an object named {target!r} already exists")
    if "FROM" not in by_word:
        raise AccessError("SELECT needs a FROM clause")
    flags, _top, items = select_list(clauses[0][1])
    sources = _sources(db, by_word["FROM"], parameters)
    specs: list[ColumnSpec] = []
    for name, expr, star in _outputs(flags, items, sources):
        if star is not None:
            source = next(s for s in sources if s.alias.lower() == star.lower())
            for column in source.columns:
                if source.table is not None:
                    specs.append(_column_copy(column, source.table, column))
                else:
                    key = f"{source.alias}.{column}".lower()
                    value = next((r[key] for r in source.rows if r.get(key) is not None), None)
                    specs.append(_value_column(column, value))
            continue
        assert expr is not None
        specs.append(_static_type(name, expr, sources, db, parameters))
    rows = _select(db, [c for c in clauses if c[0] != "INTO"], parameters)
    table = db.create_table(
        target, _made_headers(specs), created=created, updated=updated, owner_updated=owner_updated
    )
    for row in rows:
        table.insert_row(_assignments(table, list(row), list(row.values())))
    if rows:
        _settle_autonumber(db, table, list(rows[0]), [list(row.values()) for row in rows])
    return len(rows)

def _settle_autonumber(db: AccessDatabase, table: Table, columns: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    """Leave the AutoNumber counter where a statement that wrote its own
    numbers leaves it: at the largest of them.  Row by row the counter
    follows each value written, so 30 then 10 would leave 10; the engine
    ends the statement at 30 (measured on INSERT ... SELECT and on a
    made table), while a later statement writing 3 does lower it to 3."""
    auto = [i for i, name in enumerate(columns) if table.definition.column(name).auto_number]
    if not auto or len(rows) < 2:
        return
    given = [v for r in rows if isinstance(v := r[auto[0]], int) and not isinstance(v, bool)]
    if not given:
        return
    top: int = max(given) & 0xFFFFFFFF
    definition = table.definition
    if definition.next_autonumber != top:
        definition.next_autonumber = top
        db.patch_definition(definition, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", top))


@dataclass
class _Progress:
    """How far an INSERT or DELETE got before it was refused, for what the
    engine leaves in the table's definition header afterwards.

    The rollback puts the rows back but not the header's counters
    (measured): the AutoNumber counter keeps every number the statement's
    rows reserved, and the row count keeps the rows written or deleted
    before the failing one."""

    table: str | None = None
    #: The counter as the last attempted row left it, when the table has
    #: an AutoNumber.
    counter: int | None = None
    #: Rows written (positive) or deleted (negative) before the failure.
    rows: int = 0

    def reserve(self, table: Table) -> None:
        """Take the next AutoNumber for a row about to be written, as the
        engine does before it looks at the row's values: a row that fails
        still moves the counter by one, and one that names its own number
        moves it by one when it fails and to that number when it lands."""
        self.table = table.name
        definition = table.definition
        if any(c.auto_number and c.type_code != TYPE_COMPLEX for c in definition.columns):
            self.counter = (definition.next_autonumber + 1) & 0xFFFFFFFF

    def keep(self, db: AccessDatabase) -> None:
        """Write the counters back after the rollback, from a definition
        read fresh off the restored page."""
        if self.table is None or (self.counter is None and not self.rows):
            return
        fresh = db.table(self.table).definition
        if self.counter is not None:
            fresh.next_autonumber = self.counter
            db.patch_definition(fresh, OFFSET_NEXT_AUTONUMBER, struct.pack("<I", self.counter))
        if self.rows:
            fresh.row_count = (fresh.row_count + self.rows) & 0xFFFFFFFF
            db.patch_definition(fresh, OFFSET_ROW_COUNT, struct.pack("<I", fresh.row_count))


def _insert(
    db: AccessDatabase,
    clauses: list[tuple[str, str]],
    parameters: Mapping[str, object],
    progress: _Progress | None = None,
) -> int:
    by_word = dict(clauses)
    head, values_clause = _split_values(clauses[0][1])
    target, _, column_list = head.partition("(")
    table = db.table(target.strip().strip("[]"))
    # Without a column list the statement has to name every column the
    # table has, the AutoNumber included; the engine counts them and
    # refuses a list of any other length.
    columns = [c.strip().strip("[]") for c in split_top_level(column_list.rsplit(")", 1)[0], ",")] if column_list else [c.name for c in table.columns]
    if values_clause is not None:
        inner = values_clause.strip()
        if inner.startswith("("):
            inner = inner[1:].rsplit(")", 1)[0]
        env = _environment({}, [], parameters, None, lambda sql, e: _run(db, sql, parameters, e))
        values = [Parser.parse(v.strip()).eval(env) for v in split_top_level(inner, ",")]
        if progress is not None:
            progress.reserve(table)
        table.insert_row(_assignments(table, columns, values))
        return 1
    if "SELECT" in by_word:
        select_clauses = [(w, b) for w, b in clauses if w != "INSERT INTO"]
        rows = _select(db, select_clauses, parameters)
        for r in rows:
            if progress is not None:
                progress.reserve(table)
            table.insert_row(_assignments(table, columns, list(r.values())))
            if progress is not None:
                progress.rows += 1
        _settle_autonumber(db, table, columns, [list(r.values()) for r in rows])
        return len(rows)
    raise AccessError("INSERT needs VALUES or SELECT")


#: What each whole-number column can hold.  A Byte is the exception:
#: the query processor stores the low byte of the number rather than
#: refusing it, so 300 lands as 44 and -1 as 255 (measured).
_INTEGER_RANGE = {TYPE_INT: (-32768, 32767), TYPE_LONG: (-(1 << 31), (1 << 31) - 1)}


def _coerce(column: ColumnDef, value: object) -> object:
    """Convert an expression's value to what the column stores, as the
    engine does when a query writes a number into a text column or a
    Currency into a Double.

    This is the query processor's conversion, which is looser than the
    one a recordset does: it truncates over-long text and wraps a Byte
    where assigning the same value to a field would be refused."""
    if value is None:
        return None
    code = column.type_code
    if code in (TYPE_TEXT, TYPE_MEMO):
        text = value if isinstance(value, str) else _text(value)
        limit = column.length // 2
        # Text is cut to fit; a Memo has no size to cut it to.
        return text[:limit] if code == TYPE_TEXT and len(text) > limit else text
    if code in (TYPE_DOUBLE, TYPE_FLOAT):
        return float(_number(value))
    if code in (TYPE_LONG, TYPE_INT, TYPE_BYTE):
        number = value if isinstance(value, int) and not isinstance(value, bool) else round(_number(value))  # pyright: ignore[reportArgumentType]
        if code == TYPE_BYTE:
            return number & 0xFF
        low, high = _INTEGER_RANGE[code]
        if not low <= number <= high:
            kind = "an Integer" if code == TYPE_INT else "a Long"
            raise AccessError(f"column {column.name!r}: {number!r} is not {kind}")
        return number
    if code == TYPE_MONEY:
        return value if isinstance(value, (Decimal, float)) else Decimal(_number(value))  # pyright: ignore[reportArgumentType]
    if code == TYPE_BOOLEAN:
        return value if isinstance(value, bool) else _number(value) != 0
    if code == TYPE_DATETIME:
        if isinstance(value, str):
            return parse_date_literal(value)
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            # A number in a Date column is the stored serial itself.
            return _from_serial(float(value))
    return value


def _assignments(table: Table, columns: Sequence[str], values: Sequence[object]) -> dict[str, object]:
    if len(values) != len(columns):
        raise AccessError("the statement lists a different number of columns and values")
    out: dict[str, object] = {}
    for name, value in zip(columns, values, strict=True):
        column = table.definition.column(name)
        if column.auto_number and value is None:
            raise AccessError(
                f"column {column.name!r} is an AutoNumber, which takes no "
                f"Null; leave it out of the statement to have one assigned"
            )
        out[column.name] = _coerce(column, value)
    return out


def _split_values(body: str) -> tuple[str, str | None]:
    """Split ``t (cols) VALUES (...)`` at its top-level VALUES keyword."""
    depth = 0
    quote: str | None = None
    upper = body.upper()
    for i, ch in enumerate(body):
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == "[":
            quote = "]"
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0 and upper.startswith("VALUES", i) and (i == 0 or not body[i - 1].isalnum()) and not body[i + 6 : i + 7].isalnum():
            return body[:i], body[i + 6 :]
    return body, None


def _has_subquery(expr: Expr) -> bool:
    if isinstance(expr, SubQuery):
        return True
    if isinstance(expr, Call):
        return any(_has_subquery(a) for a in expr.args)
    if isinstance(expr, Binary):
        return _has_subquery(expr.left) or _has_subquery(expr.right)
    if isinstance(expr, Unary):
        return _has_subquery(expr.operand)
    if isinstance(expr, InList):
        return _has_subquery(expr.operand) or any(_has_subquery(o) for o in expr.options)
    if isinstance(expr, Between):
        return any(_has_subquery(e) for e in (expr.operand, expr.low, expr.high))
    return False


def _write_target(sources: Sequence[Source], qualifier: str | None, column: str | None) -> Source:
    """The table an UPDATE or DELETE writes to: the one an alias names,
    or the one that holds the column when nothing qualifies it."""
    tables = [s for s in sources if s.table is not None]
    if qualifier is not None:
        source = next((s for s in tables if s.alias.lower() == qualifier.lower()), None)
        if source is None:
            raise AccessError(f"{qualifier!r} is not a table of the statement")
        return source
    if column is None:
        if len(sources) != 1 or not tables:
            raise AccessError(
                "DELETE over a join has to name the table to take rows from, "
                "as DELETE T.* FROM T INNER JOIN ..."
            )
        return tables[0]
    holders = [s for s in tables if s.holds(column)]
    if not holders:
        raise AccessError(f"no column named {column!r} in any table of the statement")
    if len(holders) > 1:
        raise AccessError(f"column {column!r} is ambiguous: it is in " + " and ".join(s.alias for s in holders))
    return holders[0]


def _split_qualified(name: str) -> tuple[str | None, str]:
    """``T.A``, ``[T].[A]`` or ``[T.A]`` into its alias and column."""
    name = name.strip()
    if name.startswith("[") and name.endswith("]") and "].[" not in name:
        return None, name[1:-1]
    qualifier, dot, column = name.rpartition(".")
    if not dot:
        return None, column.strip("[]")
    return qualifier.strip("[]"), column.strip("[]")


def _row_id(row: Row, alias: str) -> RowId | None:
    """The home slot of the row an alias contributed, when the join gave
    it one -- the unmatched side of an outer join has none."""
    from pyopenvba.access.database import RowId

    value = row.get("__rowid__." + alias.lower())
    return value if isinstance(value, RowId) else None


def _update(db: AccessDatabase, clauses: list[tuple[str, str]], parameters: Mapping[str, object]) -> int:
    """Write the SET clause into the tables the FROM part joins, one join
    row at a time in the order :func:`_join` produces them.  Every join
    row counts as affected.  Within a row every expression is read before
    the row's writes, so ``SET A = B, B = A`` swaps; a later join row that
    reaches a table row an earlier one wrote reads the new value, so the
    last join row's write is what stays (measured -- which row that is
    depends on the join order).  A LEFT JOIN writes its unmatched rows
    too, with Null for the side that is missing."""
    by_word = dict(clauses)
    sources = _sources(db, clauses[0][1], parameters)
    assignments: list[tuple[Source, str, Expr]] = []
    for item in split_top_level(by_word.get("SET", ""), ","):
        column, _, expression = item.partition("=")
        qualifier, name = _split_qualified(column)
        source = _write_target(sources, qualifier, name)
        assert source.table is not None
        target = source.table.definition.column(name)
        if target.auto_number:
            raise AccessError(
                f"column {target.name!r} is an AutoNumber, which no query updates"
            )
        expr = Parser.parse(expression.strip())
        if _has_subquery(expr):
            # A subquery in a WHERE clause is fine; one in a SET clause
            # makes the whole query not updateable.
            raise AccessError(
                "a SET clause cannot take its value from a subquery: the "
                "engine treats such an UPDATE as not updateable"
            )
        assignments.append((source, target.name, expr))
    where = Parser.parse(by_word["WHERE"]) if "WHERE" in by_word else None

    def runner(sql: str, env: Environment) -> list[Row]:
        return _run(db, sql, parameters, env)

    count = 0
    # What each table row holds now, once a join row has written it: a
    # later join row that reaches the same table row reads these.
    current: dict[tuple[str, RowId], dict[str, object]] = {}
    for row in _join(sources, parameters, None):
        for source in sources:
            row_id = _row_id(row, source.alias)
            if row_id is not None and (source.alias.lower(), row_id) in current:
                row = {**row, **current[(source.alias.lower(), row_id)]}
        env = _environment(row, sources, parameters, None, runner)
        if where is not None:
            verdict = where.eval(env)
            if verdict is None or not _truthy(verdict):
                continue
        values = [(source, name, expr.eval(env)) for source, name, expr in assignments]
        count += 1
        for source in sources:
            changes = {name: value for s, name, value in values if s is source}
            if not changes or source.table is None:
                continue
            row_id = _row_id(row, source.alias)
            if row_id is None:
                continue
            stored = _assignments(source.table, list(changes), list(changes.values()))
            source.table.update_row(row_id, stored)
            current.setdefault((source.alias.lower(), row_id), {}).update(
                {f"{source.alias}.{name}".lower(): value for name, value in stored.items()}
            )
    return count


def _delete(
    db: AccessDatabase,
    clauses: list[tuple[str, str]],
    parameters: Mapping[str, object],
    progress: _Progress | None = None,
) -> int:
    """Take rows out of one table of the FROM part.  Over a join the
    statement has to say which, as ``DELETE T.* FROM T INNER JOIN ...``,
    and a row the join reaches twice is an error rather than a second
    delete: the engine says "Record is deleted" and undoes the lot, though
    its row count keeps the rows it had taken out by then (measured)."""
    by_word = dict(clauses)
    sources = _sources(db, by_word["FROM"], parameters)
    head = clauses[0][1].strip()
    if head.endswith(".*"):
        head = head[:-2]
    qualifier = head.strip().strip("[]") if head not in ("", "*") else None
    target = _write_target(sources, qualifier, None)
    table = target.table
    assert table is not None
    if "WHERE" not in by_word and len(sources) == 1:
        count = table.row_count
        table.truncate()
        return count
    where = Parser.parse(by_word["WHERE"]) if "WHERE" in by_word else None

    def runner(sql: str, env: Environment) -> list[Row]:
        return _run(db, sql, parameters, env)

    doomed: list[RowId] = []
    for row in _join(sources, parameters, None):
        if where is not None:
            verdict = where.eval(_environment(row, sources, parameters, None, runner))
            if verdict is None or not _truthy(verdict):
                continue
        row_id = _row_id(row, target.alias)
        if row_id is None:
            continue
        if row_id in doomed:
            if progress is not None:
                progress.table = table.name
                progress.rows = -len(doomed)
            raise AccessError(
                f"the join reaches a row of {target.alias!r} more than once, so it "
                f"would be deleted twice; nothing was deleted"
            )
        doomed.append(row_id)
    for row_id in doomed:
        table.delete_row(row_id)
    return len(doomed)
