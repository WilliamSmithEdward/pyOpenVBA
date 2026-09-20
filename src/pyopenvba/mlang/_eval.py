"""Evaluating M.

Bindings are lazy, as M's are: a ``let`` step is worked out the first
time something asks for it, so a step may name one written after it and
a step nobody uses is never run.  The same goes for a record's fields.

``each x`` is a function of one argument called ``_``, and a bare field
name inside one reads that argument's field, which is what makes
``each [Amount] > 100`` work.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field

from pyopenvba.mlang import _parse as P
from pyopenvba.mlang._values import (
    is_list,
    Builtin,
    Duration,
    Function,
    MError,
    MType,
    Parameter,
    Record,
    Table,
    as_logical,
    as_number,
    as_text,
    number_out,
    type_name,
)

#: How deep a chain of bindings may go before it is called circular.
MAX_DEPTH = 256


class Thunk:
    """A binding that has not been worked out yet."""

    __slots__ = ("node", "scope", "value", "state")

    def __init__(self, node: P.Node, scope: Scope) -> None:
        self.node = node
        self.scope = scope
        self.value: object = None
        self.state = "pending"

    def force(self) -> object:
        if self.state == "done":
            return self.value
        if self.state == "working":
            raise MError("Expression.Error", "A cyclic reference was encountered.")
        self.state = "working"
        try:
            self.value = evaluate(self.node, self.scope)
        finally:
            if self.state == "working":
                self.state = "done"
        self.state = "done"
        return self.value


@dataclass(slots=True)
class Scope:
    """Names in view, and where to look next."""

    names: dict[str, object] = field(default_factory=lambda: {})
    parent: Scope | None = None
    #: The value ``_`` stands for inside an ``each``.
    underscore: object = None
    has_underscore: bool = False

    def child(self, names: dict[str, object] | None = None) -> Scope:
        return Scope(names=names or {}, parent=self)

    def lookup(self, name: str) -> object:
        scope: Scope | None = self
        while scope is not None:
            if name in scope.names:
                found = scope.names[name]
                return found.force() if isinstance(found, Thunk) else found
            scope = scope.parent
        # Inside an each, a bare name is a field of the row.
        row = self.current_underscore()
        if row is not None and isinstance(row, Record) and row.has(name):
            return row[name]
        from pyopenvba.mlang._library import missing

        raise missing(name)

    def current_underscore(self) -> object:
        scope: Scope | None = self
        while scope is not None:
            if scope.has_underscore:
                return scope.underscore
            scope = scope.parent
        return None


def evaluate(node: P.Node, scope: Scope) -> object:
    """What one expression comes to."""
    kind = type(node)
    if kind is P.Literal:
        return node.value  # type: ignore[attr-defined]
    if kind is P.Name:
        return _name(node, scope)  # type: ignore[arg-type]
    if kind is P.Let:
        return _let(node, scope)  # type: ignore[arg-type]
    if kind is P.If:
        return _if(node, scope)  # type: ignore[arg-type]
    if kind is P.RecordLiteral:
        return _record(node, scope)  # type: ignore[arg-type]
    if kind is P.ListLiteral:
        listed: P.ListLiteral = node  # type: ignore[assignment]
        out: list[object] = []
        for item in listed.items:
            made = evaluate(item, scope)
            if isinstance(item, P.Call) and isinstance(item.target, P.Name) and item.target.name == "#range":
                out.extend(made if is_list(made) else [made])
            else:
                out.append(made)
        return out
    if kind is P.FunctionLiteral:
        written: P.FunctionLiteral = node  # type: ignore[assignment]
        return Function(
            parameters=[
                Parameter(name=name, optional=optional, declared=declared)
                for name, optional, declared in written.parameters
            ],
            body=written.body,
            closure=scope,
        )
    if kind is P.Each:
        return Function(
            parameters=[Parameter(name="_")], body=node.body, closure=scope, name="each"  # type: ignore[attr-defined]
        )
    if kind is P.Call:
        return _call_node(node, scope)  # type: ignore[arg-type]
    if kind is P.FieldAccess:
        return _field(node, scope)  # type: ignore[arg-type]
    if kind is P.ItemAccess:
        return _item(node, scope)  # type: ignore[arg-type]
    if kind is P.Projection:
        return _projection(node, scope)  # type: ignore[arg-type]
    if kind is P.Unary:
        return _unary(node, scope)  # type: ignore[arg-type]
    if kind is P.Binary:
        return _binary(node, scope)  # type: ignore[arg-type]
    if kind is P.TryExpr:
        return _try(node, scope)  # type: ignore[arg-type]
    if kind is P.ErrorExpr:
        raise _raised(evaluate(node.body, scope) if node.body is not None else None)  # type: ignore[attr-defined]
    if kind is P.TypeExpr:
        return MType(name=node.name)  # type: ignore[attr-defined]
    if kind is P.NotImplemented_:
        raise MError("Expression.Error", "The expression is not implemented.")
    raise MError("Expression.Error", f"cannot evaluate {kind.__name__}")


def _raised(value: object) -> MError:
    if isinstance(value, Record):
        return MError(
            as_text(value.get("Reason", "Expression.Error")),
            as_text(value.get("Message", "")),
            value.get("Detail"),
        )
    return MError("Expression.Error", as_text(value) if value is not None else "")


def _name(node: P.Name, scope: Scope) -> object:
    if node.name == "_":
        found = scope.current_underscore()
        if found is None and not _has_underscore(scope):
            raise MError("Expression.Error", "The name '_' wasn't recognized.")
        return found
    return scope.lookup(node.name)


def _has_underscore(scope: Scope) -> bool:
    current: Scope | None = scope
    while current is not None:
        if current.has_underscore:
            return True
        current = current.parent
    return False


def _let(node: P.Let, scope: Scope) -> object:
    inner = scope.child()
    for name, expression in node.bindings:
        inner.names[name] = Thunk(expression, inner)
    if node.body is None:
        raise MError("Expression.Error", "a let with no in")
    return evaluate(node.body, inner)


def _if(node: P.If, scope: Scope) -> object:
    if node.condition is None:
        raise MError("Expression.Error", "an if with no condition")
    taken = node.then if as_logical(evaluate(node.condition, scope)) else node.otherwise
    if taken is None:
        raise MError("Expression.Error", "an if with no branch")
    return evaluate(taken, scope)


def _record(node: P.RecordLiteral, scope: Scope) -> object:
    inner = scope.child()
    thunks: dict[str, Thunk] = {}
    for name, expression in node.fields:
        thunks[name] = Thunk(expression, inner)
        inner.names[name] = thunks[name]
    return Record({name: thunk.force() for name, thunk in thunks.items()})


def _call_node(node: P.Call, scope: Scope) -> object:
    if node.target is None:
        raise MError("Expression.Error", "a call with nothing to call")
    target = evaluate(node.target, scope)
    args = [evaluate(argument, scope) for argument in node.args]
    return apply(target, args)


def apply(target: object, args: list[object]) -> object:
    """Call a function value with already-evaluated arguments."""
    if isinstance(target, Builtin):
        if len(args) < target.minimum or (target.maximum is not None and len(args) > target.maximum):
            raise MError(
                "Expression.Error",
                f"{target.name} takes {target.minimum} argument"
                f"{'s' if target.minimum != 1 else ''} and was given {len(args)}.",
            )
        return target.call(*args)
    if isinstance(target, Function):
        closure = target.closure
        if closure is None:  # pragma: no cover - a function always has one
            raise MError("Expression.Error", "a function with no scope")
        inner = closure.child()
        required = [one for one in target.parameters if not one.optional]
        if len(args) < len(required):
            raise MError(
                "Expression.Error",
                f"a function taking {len(required)} argument"
                f"{'s' if len(required) != 1 else ''} was given {len(args)}.",
            )
        for index, parameter in enumerate(target.parameters):
            inner.names[parameter.name] = args[index] if index < len(args) else None
        if target.parameters and target.parameters[0].name == "_":
            inner.underscore = args[0] if args else None
            inner.has_underscore = True
        if target.body is None:
            raise MError("Expression.Error", "a function with no body")
        return evaluate(target.body, inner)
    raise MError(
        "Expression.Error", f"We cannot apply a value of type {type_name(target)} as a function."
    )


def _field(node: P.FieldAccess, scope: Scope) -> object:
    target = evaluate(node.target, scope) if node.target is not None else scope.current_underscore()
    if target is None and node.optional:
        return None
    if isinstance(target, Record):
        if node.optional and not target.has(node.name):
            return None
        return target[node.name]
    if isinstance(target, Table):
        # A table's field is its column, as a list.
        return target.column(node.name)
    if target is None:
        raise MError(
            "Expression.Error", f"We cannot get field '{node.name}' from a value of type null."
        )
    raise MError(
        "Expression.Error",
        f"We cannot get field '{node.name}' from a value of type {type_name(target)}.",
    )


def _item(node: P.ItemAccess, scope: Scope) -> object:
    target = evaluate(node.target, scope) if node.target is not None else None
    index = evaluate(node.index, scope) if node.index is not None else None
    if is_list(target):
        if isinstance(index, (int, float)) and not isinstance(index, bool):
            position = int(index)
            if 0 <= position < len(target):
                return target[position]
            if node.optional:
                return None
            raise MError("Expression.Error", "There weren't enough elements in the enumeration.")
    if isinstance(target, Table):
        if isinstance(index, (int, float)) and not isinstance(index, bool):
            position = int(index)
            if 0 <= position < target.height:
                return target.record_at(position)
            if node.optional:
                return None
            raise MError("Expression.Error", "There weren't enough elements in the enumeration.")
        if isinstance(index, Record):
            for row in target.records():
                if all(row.get(name) == value for name, value in index.fields.items()):
                    return row
            if node.optional:
                return None
            raise MError("Expression.Error", "The key didn't match any rows in the table.")
    raise MError(
        "Expression.Error", f"We cannot index into a value of type {type_name(target)}."
    )


def _projection(node: P.Projection, scope: Scope) -> object:
    target = evaluate(node.target, scope) if node.target is not None else None
    if isinstance(target, Record):
        out: dict[str, object] = {}
        for name in node.names:
            if not target.has(name) and node.optional:
                out[name] = None
                continue
            out[name] = target[name]
        return Record(out)
    if isinstance(target, Table):
        indexes = [target.column_index(name) for name in node.names]
        return Table(list(node.names), [[row[at] for at in indexes] for row in target.rows])
    raise MError(
        "Expression.Error", f"We cannot select fields from a value of type {type_name(target)}."
    )


def _unary(node: P.Unary, scope: Scope) -> object:
    value = evaluate(node.operand, scope) if node.operand is not None else None
    if node.op == "not":
        return not as_logical(value)
    if value is None:
        return None
    return number_out(-as_number(value))


def _try(node: P.TryExpr, scope: Scope) -> object:
    try:
        value = evaluate(node.body, scope) if node.body is not None else None
    except MError as failure:
        if node.otherwise is not None:
            return evaluate(node.otherwise, scope)
        return Record({"HasError": True, "Error": failure.as_record()})
    if node.otherwise is not None:
        return value
    return Record({"HasError": False, "Value": value})


def _elapsed(left: _dt.date, right: _dt.date) -> Duration:
    """One date taken from another, which M answers as a duration."""
    return Duration(seconds=(_as_moment(left) - _as_moment(right)).total_seconds())


def _moved(left: _dt.date, right: Duration, *, minus: bool) -> object:
    """A date or a moment moved along by a duration.

    A date stays a date: the engine answers ``#date(2021, 3, 4)`` for
    that date plus twelve hours, keeping the day the moment lands on
    rather than promoting it to a datetime.
    """
    gap = _dt.timedelta(seconds=right.total_seconds)
    moved = _as_moment(left) - gap if minus else _as_moment(left) + gap
    return moved if isinstance(left, _dt.datetime) else moved.date()


def _as_moment(value: _dt.date) -> _dt.datetime:
    """A date read as the midnight that starts it."""
    if isinstance(value, _dt.datetime):
        return value
    return _dt.datetime(value.year, value.month, value.day)


def _binary(node: P.Binary, scope: Scope) -> object:
    op = node.op
    left = evaluate(node.left, scope) if node.left is not None else None
    if op == "and":
        if not as_logical(left):
            return False
        return as_logical(evaluate(node.right, scope) if node.right is not None else None)
    if op == "or":
        if as_logical(left):
            return True
        return as_logical(evaluate(node.right, scope) if node.right is not None else None)
    if op == "??":
        if left is not None:
            return left
        return evaluate(node.right, scope) if node.right is not None else None
    right = evaluate(node.right, scope) if node.right is not None else None
    if op == "-" and isinstance(left, _dt.date) and isinstance(right, _dt.date):
        return _elapsed(left, right)
    if op in ("+", "-") and isinstance(left, _dt.date) and isinstance(right, Duration):
        return _moved(left, right, minus=op == "-")
    if op == "+" and isinstance(left, Duration) and isinstance(right, _dt.date):
        return _moved(right, left, minus=False)
    if op in ("+", "-") and isinstance(left, Duration) and isinstance(right, Duration):
        gap = left.total_seconds + (-right.total_seconds if op == "-" else right.total_seconds)
        return Duration(seconds=gap)
    if op in ("*", "/") and isinstance(left, Duration) and isinstance(right, (int, float)):
        factor = as_number(right)
        if op == "/" and factor == 0:
            raise MError("Expression.Error", "We cannot divide a duration by zero.")
        return Duration(seconds=left.total_seconds * factor if op == "*" else left.total_seconds / factor)
    if op == "*" and isinstance(left, (int, float)) and isinstance(right, Duration):
        return Duration(seconds=right.total_seconds * as_number(left))
    if op == "as":
        return left
    if op == "is":
        wanted = right.name if isinstance(right, MType) else as_text(right)
        return is_type(left, wanted)
    if op == "meta":
        return left
    return _apply_operator(op, left, right)


def is_type(value: object, wanted: str) -> bool:
    if wanted.startswith("nullable "):
        return value is None or is_type(value, wanted[9:])
    actual = type_name(value)
    if wanted in ("any", ""):
        return True
    if wanted == actual:
        return True
    return wanted.lower() == actual


def _apply_operator(op: str, left: object, right: object) -> object:
    if op in ("=", "<>"):
        same = _equal(left, right)
        return same if op == "=" else not same
    if op in ("<", ">", "<=", ">="):
        return _ordered(op, left, right)
    if op == "&":
        return _concatenate(left, right)
    if left is None or right is None:
        return None
    first = as_number(left)
    second = as_number(right)
    if op == "+":
        return number_out(first + second)
    if op == "-":
        return number_out(first - second)
    if op == "*":
        return number_out(first * second)
    if op == "/":
        if second == 0:
            return float("inf") if first > 0 else (float("-inf") if first < 0 else float("nan"))
        return number_out(first / second)
    raise MError("Expression.Error", f"unknown operator {op!r}")


def _equal(left: object, right: object) -> bool:
    if isinstance(left, Record) and isinstance(right, Record):
        return left.fields == right.fields
    if isinstance(left, Table) and isinstance(right, Table):
        return left.columns == right.columns and left.rows == right.rows
    if isinstance(left, bool) or isinstance(right, bool):
        return left is right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return float(left) == float(right)
    return left == right


def _ordered(op: str, left: object, right: object) -> object:
    if left is None or right is None:
        return None
    if isinstance(left, str) and isinstance(right, str):
        order = (left > right) - (left < right)
    elif isinstance(left, _dt.date) and isinstance(right, _dt.date):
        order = (left > right) - (left < right)  # type: ignore[operator]
    else:
        one = as_number(left)
        two = as_number(right)
        order = (one > two) - (one < two)
    if op == "<":
        return order < 0
    if op == ">":
        return order > 0
    if op == "<=":
        return order <= 0
    return order >= 0


def _concatenate(left: object, right: object) -> object:
    """``&`` joins text, lists, records and tables, each in its own way."""
    if isinstance(left, str) or isinstance(right, str):
        if left is None or right is None:
            return None
        return as_text(left) + as_text(right)
    if is_list(left) and is_list(right):
        return [*left, *right]
    if isinstance(left, Record) and isinstance(right, Record):
        merged = dict(left.fields)
        merged.update(right.fields)
        return Record(merged)
    if isinstance(left, Table) and isinstance(right, Table):
        from pyopenvba.mlang._library import combine_tables

        return combine_tables([left, right])
    if isinstance(left, _dt.date) and isinstance(right, _dt.time):
        return _dt.datetime.combine(left, right)
    if left is None or right is None:
        return None
    raise MError(
        "Expression.Error",
        f"We cannot apply operator & to types {type_name(left)} and {type_name(right)}.",
    )


def base_scope(extra: dict[str, object] | None = None) -> Scope:
    """The library, plus whatever the host adds to it."""
    from pyopenvba.mlang._library import LIBRARY

    names: dict[str, object] = dict(LIBRARY)
    if extra:
        names.update(extra)
    return Scope(names=names)


def run(source: str, extra: dict[str, object] | None = None) -> object:
    """Evaluate one M expression with the library in view."""
    return evaluate(P.parse(source), base_scope(extra))


def call_value(target: object, args: list[object]) -> object:
    """Apply a function value, for the library's own callbacks."""
    return apply(target, args)


