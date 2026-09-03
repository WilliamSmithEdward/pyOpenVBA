"""The rules a table keeps about its own rows.

Required, DefaultValue, AllowZeroLength, ValidationRule and
ValidationText are not in the table definition.  Each is a property on
the column in the definition's property blob (a table's own rule is a
property on the table), and the engine applies them itself on the way
in, so a file written without them holds rows the engine would have
refused.  Measured against DAO on a three-column table:

* ``DefaultValue`` fills any column an INSERT does not name, through SQL
  as well as through a recordset's AddNew.  Its text is a Jet expression
  with an optional leading ``=``, and a bare name that is not a function
  is its own text: ``a & b`` defaults to ``ab``, ``1+1`` to ``2`` and
  ``=Date()`` to today.
* ``Required`` refuses a null, whether the column was left out or set to
  Null, on insert and on update, with the message the engine gives.
* ``ValidationRule`` on a column is checked on insert and on update.  A
  rule that starts with an operator is about the column itself, so ``>0``
  reads as ``[A]>0``.  The message is ``ValidationText`` when there is
  one, else the engine's own sentence.  A rule on the table is a whole
  expression over the row.  A rule refuses a row only when it comes out
  False: a null makes the comparison Null and the engine takes the row.
* ``AllowZeroLength`` is not enforced by the engine through SQL (an
  empty string went in with it off), so it is stored and read, never
  checked.

DAO's ``Field.ValidationRule`` setter leaves a trailing NUL in the
property, which is what both DAO and this read back, so the text is
stripped before it is parsed and kept whole in the engine's message.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba.access_read import AccessError

if TYPE_CHECKING:
    from collections.abc import Mapping

    from pyopenvba.access._sql import Env
    from pyopenvba.access.database import Table

# A rule that opens with one of these is about its own column, so the
# column reference goes in front of it (Access writes them both ways).
OPERATOR_START = re.compile(
    r"^\s*(<>|<=|>=|=|<|>|(?:NOT\s+)?LIKE\b|(?:NOT\s+)?IN\b|(?:NOT\s+)?BETWEEN\b|IS\b)",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Rule:
    """One ValidationRule with the ValidationText that goes with it."""

    text: str
    message: str | None = None


@dataclass
class Rules:
    """What one table's properties say about the rows it will take."""

    defaults: dict[str, str] = field(default_factory=lambda: {})
    required: set[str] = field(default_factory=lambda: set())
    columns: dict[str, Rule] = field(default_factory=lambda: {})
    table: Rule | None = None

    def __bool__(self) -> bool:
        return bool(self.defaults or self.required or self.columns or self.table)


def read(table: Table) -> Rules:
    """A table's rules, read from its property blob in one pass."""
    rules = Rules()
    blob = table.property_blob()
    own_rule = blob.decoded().get("ValidationRule")
    if isinstance(own_rule, str) and _clean(own_rule):
        rules.table = Rule(_clean(own_rule), _message(blob.decoded().get("ValidationText")))
    for column in table.definition.columns:
        properties = blob.decoded_column(column.name)
        text = properties.get("DefaultValue")
        if isinstance(text, str) and text.strip():
            rules.defaults[column.name] = text
        if properties.get("Required") is True:
            rules.required.add(column.name)
        rule = properties.get("ValidationRule")
        if isinstance(rule, str) and _clean(rule):
            rules.columns[column.name] = Rule(_clean(rule), _message(properties.get("ValidationText")))
    return rules


def _clean(rule: str) -> str:
    return rule.rstrip("\x00").strip()


def _message(text: object) -> str | None:
    return text if isinstance(text, str) and text else None


def _evaluate(text: str, resolve: Env) -> object:
    from pyopenvba.access._sql import Parser

    return Parser.parse(text).eval(resolve)


def default_value(text: str) -> object:
    """What a DefaultValue property means.  A leading ``=`` is dropped and
    a name the expression cannot resolve is its own text, which is how
    ``hello`` defaults to the string and ``a & b`` to ``ab``."""
    body = text.strip()
    if body.startswith("="):
        body = body[1:]
    if not body:
        return None
    return _evaluate(body, lambda name, _qualifier: name)


def apply_defaults(rules: Rules, values: Mapping[str, object]) -> dict[str, object]:
    """``values`` with every column it does not name filled in from that
    column's DefaultValue, as the engine fills them."""
    out = dict(values)
    named = {name.lower() for name in out}
    for name, text in rules.defaults.items():
        if name.lower() not in named:
            out[name] = default_value(text)
    return out


def check(table_name: str, rules: Rules, values: Mapping[str, object], *, columns: set[str] | None = None) -> None:
    """Refuse a row the engine would refuse: a null in a Required column
    and a value against a ValidationRule, the column's and the table's.
    ``columns`` narrows the column rules to the ones being written, which
    is what an update touches."""
    lowered = {name.lower(): value for name, value in values.items()}
    for name in sorted(rules.required):
        if columns is not None and name not in columns:
            continue
        if lowered.get(name.lower()) is None:
            raise AccessError(f"You must enter a value in the '{table_name}.{name}' field.")
    for name, rule in rules.columns.items():
        if columns is not None and name not in columns:
            continue
        if lowered.get(name.lower()) is None:
            continue
        text = rule.text
        if OPERATOR_START.match(text):
            text = f"[{name}]{text}"
        if not _passes(text, lowered):
            raise AccessError(rule.message or _refusal(rule.text, f"{table_name}.{name}"))
    if rules.table is not None and not _passes(rules.table.text, lowered):
        raise AccessError(rules.table.message or _refusal(rules.table.text, table_name))


def _refusal(rule: str, where: str) -> str:
    return (
        f"One or more values are prohibited by the validation rule '{rule}' set for '{where}'. "
        f"Enter a value that the expression for this field can accept."
    )


def _passes(rule: str, lowered: Mapping[str, object]) -> bool:
    def resolve(name: str, _qualifier: str | None) -> object:
        key = name.lower()
        if key not in lowered:
            raise AccessError(f"the validation rule {rule!r} names no column {name!r}")
        return lowered[key]

    return _evaluate(rule, resolve) is not False
