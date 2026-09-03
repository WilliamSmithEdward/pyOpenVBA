"""A Jet SQL executor over the storage engine.

``execute(db, sql, parameters)`` runs SELECT, INSERT, UPDATE and DELETE
statements against :class:`~pyopenvba.access.database.AccessDatabase`
tables, in pure Python.  SELECT covers a column list or ``*``, one table
or INNER / LEFT / RIGHT JOINs, WHERE, GROUP BY with Count, Sum, Avg, Min
and Max, HAVING, ORDER BY, DISTINCT and TOP; expressions cover the
comparison, logical, arithmetic and concatenation operators, LIKE with
the engine's wildcards, IN, BETWEEN, IS NULL, ``[parameters]`` and a set
of common functions (Len, UCase, LCase, Trim, Left, Right, Mid, InStr,
Abs, Int, Round, Nz, IIf, Year, Month, Day, Date, Now).  Three-valued
logic follows the engine: a comparison against Null is Null, and a WHERE
that comes out Null drops the row.

The statement text is split with the same clause splitter the saved-query
writer uses (:mod:`_queries`).  DML goes through the table writers, so a
DELETE with no WHERE takes the engine's truncation path and everything
else the row path, page for page as the engine would.
"""

from __future__ import annotations

import datetime as _dt
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from pyopenvba.access._queries import parse_from, select_list, split_clauses, split_top_level
from pyopenvba.access._tdef import (
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_MONEY,
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
  | (?P<number>\d+\.\d*|\.\d+|\d+)
  | (?P<bracket>\[[^\]]+\])
  | (?P<name>[A-Za-z_][A-Za-z0-9_]*)
  | (?P<op><>|<=|>=|=|<|>|\+|-|\*|/|&|\(|\)|,|\.)
    """,
    re.VERBOSE,
)


@dataclass(frozen=True)
class Token:
    kind: str
    text: str


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
        out.append(Token(kind, match.group(0)))
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
            if isinstance(a, str) or isinstance(b, str):
                return _text(a) + _text(b)
            return _arith(a, b, lambda x, y: x + y)
        if op == "-":
            return _arith(a, b, lambda x, y: x - y)
        if op == "*":
            return _arith(a, b, lambda x, y: x * y)
        if op == "/":
            if _number(b) == 0:
                raise AccessError("division by zero")
            return float(_number(a)) / float(_number(b))
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
        if self.is_aggregate:
            raise AccessError(f"{self.name} needs a GROUP BY or a whole-table aggregate context")
        values = [a.eval(env) for a in self.args]
        return _call(self.name, values)

    def columns(self) -> list[tuple[str | None, str]]:
        return [c for a in self.args for c in a.columns()]


AGGREGATES = {"COUNT", "SUM", "AVG", "MIN", "MAX", "FIRST", "LAST"}


class Parser:
    """Precedence-climbing parser for Jet SQL expressions."""

    def __init__(self, tokens: list[Token]) -> None:
        self.tokens = tokens
        self.pos = 0

    @classmethod
    def parse(cls, text: str) -> Expr:
        parser = cls(tokenize(text))
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
            return Unary("NOT", self.negation())
        return self.comparison()

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
            return Binary(token.text, left, self.concatenation())
        return left

    def concatenation(self) -> Expr:
        left = self.additive()
        while self.peek("&"):
            self.take()
            left = Binary("&", left, self.additive())
        return left

    def additive(self) -> Expr:
        left = self.multiplicative()
        while (token := self.peek("+", "-")) is not None:
            self.take()
            left = Binary(token.text, left, self.multiplicative())
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
        return self.atom()

    def atom(self) -> Expr:
        token = self.take()
        if token.kind == "number":
            return Literal(float(token.text) if "." in token.text else int(token.text))
        if token.kind == "string":
            quote = token.text[0]
            return Literal(token.text[1:-1].replace(quote + quote, quote))
        if token.kind == "date":
            return Literal(_parse_date_literal(token.text[1:-1]))
        if token.text == "(":
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


def _truthy(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        return value != 0
    if isinstance(value, str):
        return value != ""
    return value is not None


def _number(value: object) -> int | float | Decimal:
    if isinstance(value, bool):
        return -1 if value else 0
    if isinstance(value, (int, float, Decimal)):
        return value
    if isinstance(value, str):
        try:
            return float(value) if "." in value else int(value)
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
        return "True" if value else "False"
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


def _parse_date_literal(text: str) -> _dt.datetime:
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
        haystack, needle = (_text(args[0]), _text(args[1])) if len(args) == 2 else (_text(args[1]), _text(args[2]))
        start = int(_number(args[0])) if len(args) > 2 else 1
        found = haystack.lower().find(needle.lower(), start - 1)
        return found + 1
    if upper == "ABS":
        return abs(_number(args[0]))  # pyright: ignore[reportArgumentType]
    if upper == "INT":
        import math

        return math.floor(float(_number(args[0])))
    if upper == "ROUND":
        places = int(_number(args[1])) if len(args) > 1 else 0
        return round(float(_number(args[0])), places)
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
    raise AccessError(f"function {name} is not available")


# --- statements ------------------------------------------------------------------


@dataclass
class Source:
    """A table in a FROM clause and how it is joined to the ones before it."""

    table: Table
    alias: str
    join: str | None = None  # None, INNER, LEFT, RIGHT
    condition: Expr | None = None


def _rows_of(table: Table, alias: str) -> Iterator[Row]:
    for row_id, values in table.rows_with_ids():
        out: Row = {}
        for name, value in values.items():
            out[f"{alias}.{name}".lower()] = value
        out["__rowid__." + alias.lower()] = row_id
        yield out


def _environment(row: Row, sources: Sequence[Source], parameters: Mapping[str, object]) -> Env:
    aliases = [s.alias.lower() for s in sources]

    def lookup(name: str, qualifier: str | None) -> object:
        key = name.lower()
        if qualifier is not None:
            q = qualifier.lower()
            if q not in aliases:
                raise AccessError(f"no table or alias {qualifier!r} in the query")
            full = f"{q}.{key}"
            if full in row:
                return row[full]
            source = sources[aliases.index(q)]
            if any(c.name.lower() == key for c in source.table.columns):
                return None
            raise AccessError(f"no column {qualifier}.{name}")
        hits = [f"{a}.{key}" for a in aliases if f"{a}.{key}" in row]
        if len(hits) > 1:
            raise AccessError(f"column {name!r} is ambiguous; qualify it")
        if hits:
            return row[hits[0]]
        for source in sources:
            if any(c.name.lower() == key for c in source.table.columns):
                return None  # a missing side of an outer join
        for pname, pvalue in parameters.items():
            if pname.lower().strip("[]") == key:
                return pvalue
        raise AccessError(f"no column or parameter {name!r}")

    return lookup


def _join(db: AccessDatabase, sources: list[Source], parameters: Mapping[str, object]) -> list[Row]:
    rows: list[Row] = list(_rows_of(sources[0].table, sources[0].alias))
    for k in range(1, len(sources)):
        source = sources[k]
        right_rows = list(_rows_of(source.table, source.alias))
        joined: list[Row] = []
        matched_right: set[int] = set()
        for left in rows:
            hit = False
            for index, right in enumerate(right_rows):
                candidate = {**left, **right}
                if source.condition is not None:
                    verdict = source.condition.eval(_environment(candidate, sources[: k + 1], parameters))
                    if verdict is None or not _truthy(verdict):
                        continue
                joined.append(candidate)
                matched_right.add(index)
                hit = True
            if not hit and source.join == "LEFT":
                joined.append(dict(left))
        if source.join == "RIGHT":
            for index, right in enumerate(right_rows):
                if index not in matched_right:
                    joined.append(dict(right))
        rows = joined
    return rows


def _aggregate(name: str, values: list[object]) -> object:
    upper = name.upper()
    present = [v for v in values if v is not None]
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
        if isinstance(total, Decimal):  # Avg over Currency stays Currency
            return (total / len(present)).quantize(Decimal("0.0001"))
        return float(_number(total)) / len(present)
    if upper == "MIN":
        return min(present, key=_sort_key)  # pyright: ignore[reportArgumentType]
    if upper == "MAX":
        return max(present, key=_sort_key)  # pyright: ignore[reportArgumentType]
    if upper == "FIRST":
        return present[0]
    if upper == "LAST":
        return present[-1]
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


def _evaluate_with_aggregates(expr: Expr, group: list[Row], sources: Sequence[Source], parameters: Mapping[str, object]) -> object:
    """Evaluate an expression over a group: aggregate calls collapse the
    group, everything else is read off its first row."""

    def rewrite(node: Expr) -> Expr:
        if isinstance(node, Call) and node.is_aggregate:
            if node.star:
                return Literal(len(group))
            if len(node.args) != 1:
                raise AccessError(f"{node.name} takes one argument")
            values = [node.args[0].eval(_environment(r, sources, parameters)) for r in group]
            return Literal(_aggregate(node.name, values))
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

    env = _environment(group[0] if group else {}, sources, parameters)
    return rewrite(expr).eval(env)


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


def _sources(db: AccessDatabase, from_clause: str) -> list[Source]:
    tables, joins = parse_from(from_clause)
    sources: list[Source] = []
    for t in tables:
        name = (t.name1 or "").strip("[]")
        alias = (t.name2 or name).strip("[]")
        sources.append(Source(db.table(name), alias))
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


def execute(db: AccessDatabase, sql: str, parameters: Mapping[str, object] | None = None) -> list[Row] | int:
    """Run one statement.  SELECT returns its rows as dicts keyed by the
    output column names (aliases, column names, or ``Expr1000``...);
    INSERT, UPDATE and DELETE return the number of rows affected."""
    parameters = dict(parameters or {})
    text = sql.strip().rstrip(";").strip()
    if text.upper().startswith("PARAMETERS "):
        _, _, text = text.partition(";")
        text = text.strip()
    clauses = split_clauses(text)
    verb = clauses[0][0]
    if verb == "SELECT":
        return _select(db, clauses, parameters)
    if verb == "INSERT INTO":
        return _insert(db, clauses, parameters)
    if verb == "UPDATE":
        return _update(db, clauses, parameters)
    if verb == "DELETE":
        return _delete(db, clauses, parameters)
    raise AccessError(f"statement {verb} is not supported")


def _select(db: AccessDatabase, clauses: list[tuple[str, str]], parameters: Mapping[str, object]) -> list[Row]:
    by_word = dict(clauses)
    if "FROM" not in by_word:
        raise AccessError("SELECT needs a FROM clause")
    flags, top, items = select_list(clauses[0][1])
    sources = _sources(db, by_word["FROM"])
    rows = _join(db, sources, parameters)
    if "WHERE" in by_word:
        where = Parser.parse(by_word["WHERE"])
        rows = [r for r in rows if (v := where.eval(_environment(r, sources, parameters))) is not None and _truthy(v)]
    # Output columns.
    outputs: list[tuple[str, Expr | None, str | None]] = []  # (name, expr, star alias)
    expr_counter = 1000
    if flags & 0x01 or not items:
        for s in sources:
            outputs.append((s.alias, None, s.alias))
    for expression, alias in items:
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
            name = f"Expr{expr_counter}"
            expr_counter += 1
        outputs.append((name, expr, None))
    grouped = "GROUP BY" in by_word or any(o[1] is not None and _has_aggregate(o[1]) for o in outputs)
    groups: list[list[Row]]
    if grouped:
        keys = [Parser.parse(g.strip()) for g in split_top_level(by_word.get("GROUP BY", ""), ",")] if "GROUP BY" in by_word else []
        buckets: dict[tuple[object, ...], list[Row]] = {}
        for r in rows:
            env = _environment(r, sources, parameters)
            key = tuple(_hashable(k.eval(env)) for k in keys)
            buckets.setdefault(key, []).append(r)
        groups = list(buckets.values()) if buckets or keys else [rows]
        if "HAVING" in by_word:
            having = Parser.parse(by_word["HAVING"])
            groups = [g for g in groups if (v := _evaluate_with_aggregates(having, g, sources, parameters)) is not None and _truthy(v)]
    else:
        groups = [[r] for r in rows]
    result: list[tuple[Row, Row]] = []
    for group in groups:
        out: Row = {}
        for name, expr, star in outputs:
            if star is not None:
                source = next(s for s in sources if s.alias.lower() == star.lower())
                first = group[0] if group else {}
                for c in source.table.columns:
                    out[c.name] = first.get(f"{star.lower()}.{c.name.lower()}")
            elif grouped:
                out[name] = _evaluate_with_aggregates(expr, group, sources, parameters)  # pyright: ignore[reportArgumentType]
            else:
                out[name] = expr.eval(_environment(group[0], sources, parameters))  # pyright: ignore[reportOptionalMemberAccess]
        result.append((out, group[0] if group else {}))
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
            result.sort(key=lambda pair, expr=expr: _order_key(pair, expr, grouped, sources, parameters), reverse=descending)
    output = [pair[0] for pair in result]
    if flags & 0x02:
        seen: set[tuple[object, ...]] = set()
        unique: list[Row] = []
        for r in output:
            key = tuple(_hashable(v) for v in r.values())
            if key not in seen:
                seen.add(key)
                unique.append(r)
        output = unique
    if top is not None:
        output = output[: int(top)]
    return output


def _order_key(pair: tuple[Row, Row], expr: Expr, grouped: bool, sources: Sequence[Source], parameters: Mapping[str, object]) -> tuple[int, object]:
    out, base = pair
    if isinstance(expr, ColumnRef) and expr.qualifier is None and expr.name in out:
        return _sort_key(out[expr.name])
    if grouped:
        return _sort_key(_evaluate_with_aggregates(expr, [base], sources, parameters))
    return _sort_key(expr.eval(_environment(base, sources, parameters)))


def _hashable(value: object) -> object:
    if isinstance(value, str):
        return value.lower()
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, Decimal):
        return float(value)
    return value


def _insert(db: AccessDatabase, clauses: list[tuple[str, str]], parameters: Mapping[str, object]) -> int:
    by_word = dict(clauses)
    head, values_clause = _split_values(clauses[0][1])
    target, _, column_list = head.partition("(")
    table = db.table(target.strip().strip("[]"))
    columns = [c.strip().strip("[]") for c in split_top_level(column_list.rsplit(")", 1)[0], ",")] if column_list else [c.name for c in table.columns if not c.auto_number]
    if values_clause is not None:
        inner = values_clause.strip()
        if inner.startswith("("):
            inner = inner[1:].rsplit(")", 1)[0]
        values = [Parser.parse(v.strip()).eval(_environment({}, [], parameters)) for v in split_top_level(inner, ",")]
        table.insert_row(_assignments(table, columns, values))
        return 1
    if "SELECT" in by_word:
        select_clauses = [(w, b) for w, b in clauses if w != "INSERT INTO"]
        rows = _select(db, select_clauses, parameters)
        for r in rows:
            table.insert_row(_assignments(table, columns, list(r.values())))
        return len(rows)
    raise AccessError("INSERT needs VALUES or SELECT")


def _coerce(column: ColumnDef, value: object) -> object:
    """Convert an expression's value to what the column stores, as the
    engine does when a query writes a number into a text column or a
    Currency into a Double."""
    if value is None:
        return None
    code = column.type_code
    if code in (TYPE_TEXT, TYPE_MEMO):
        return value if isinstance(value, str) else _text(value)
    if code in (TYPE_DOUBLE, TYPE_FLOAT):
        return float(_number(value))
    if code in (TYPE_LONG, TYPE_INT, TYPE_BYTE):
        return value if isinstance(value, int) and not isinstance(value, bool) else round(_number(value))  # pyright: ignore[reportArgumentType]
    if code == TYPE_MONEY:
        return value if isinstance(value, (Decimal, float)) else Decimal(_number(value))  # pyright: ignore[reportArgumentType]
    if code == TYPE_BOOLEAN:
        return value if isinstance(value, bool) else _number(value) != 0
    if code == TYPE_DATETIME and isinstance(value, str):
        return _parse_date_literal(value)
    return value


def _assignments(table: Table, columns: Sequence[str], values: Sequence[object]) -> dict[str, object]:
    if len(values) != len(columns):
        raise AccessError("the statement lists a different number of columns and values")
    out: dict[str, object] = {}
    for name, value in zip(columns, values, strict=True):
        column = table.definition.column(name)
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


def _update(db: AccessDatabase, clauses: list[tuple[str, str]], parameters: Mapping[str, object]) -> int:
    by_word = dict(clauses)
    sources = _sources(db, clauses[0][1])
    if len(sources) != 1:
        raise AccessError("UPDATE over a join is not supported")
    table = sources[0].table
    assignments: list[tuple[str, Expr]] = []
    for item in split_top_level(by_word.get("SET", ""), ","):
        column, _, expression = item.partition("=")
        name = column.strip().strip("[]")
        if "." in name:
            name = name.split(".", 1)[1].strip("[]")
        assignments.append((table.definition.column(name).name, Parser.parse(expression.strip())))
    where = Parser.parse(by_word["WHERE"]) if "WHERE" in by_word else None
    count = 0
    for row in list(_rows_of(table, sources[0].alias)):
        env = _environment(row, sources, parameters)
        if where is not None:
            verdict = where.eval(env)
            if verdict is None or not _truthy(verdict):
                continue
        changes = _assignments(table, [name for name, _ in assignments], [expr.eval(env) for _, expr in assignments])
        table.update_row(row["__rowid__." + sources[0].alias.lower()], changes)  # pyright: ignore[reportArgumentType]
        count += 1
    return count


def _delete(db: AccessDatabase, clauses: list[tuple[str, str]], parameters: Mapping[str, object]) -> int:
    by_word = dict(clauses)
    sources = _sources(db, by_word["FROM"])
    if len(sources) != 1:
        raise AccessError("DELETE over a join is not supported")
    table = sources[0].table
    if "WHERE" not in by_word:
        count = table.row_count
        table.truncate()
        return count
    where = Parser.parse(by_word["WHERE"])
    doomed: list[RowId] = []
    for row in _rows_of(table, sources[0].alias):
        verdict = where.eval(_environment(row, sources, parameters))
        if verdict is not None and _truthy(verdict):
            doomed.append(row["__rowid__." + sources[0].alias.lower()])  # pyright: ignore[reportArgumentType]
    for row_id in doomed:
        table.delete_row(row_id)  # pyright: ignore[reportArgumentType]
    return len(doomed)
