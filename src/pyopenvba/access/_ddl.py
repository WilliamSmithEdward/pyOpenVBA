"""Jet DDL over the storage engine: CREATE, DROP and ALTER.

Each statement is turned into the writers the engine was measured
against, so ``db.execute("CREATE TABLE ...")`` leaves the same bytes as
``db.create_table(...)`` with the same arguments.

The type words are Jet's, read back from tables the engine made
(``docs/access_engine.md``).  Two of them trip people up: ``INTEGER`` is
four bytes here, the two-byte type being ``SHORT`` or ``SMALLINT``, and
``CHAR`` makes a *fixed-width* Text column, which this cannot write yet
and so refuses.  ``DECIMAL`` and ``NUMERIC`` are refused because the Jet
parser itself has no such type: they only reach a database through
another provider.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.access._queries import split_top_level
from pyopenvba.access._schema import ColumnSpec, IndexSpec
from pyopenvba.access._tdef import (
    TYPE_BIGINT,
    TYPE_BINARY,
    TYPE_BOOLEAN,
    TYPE_BYTE,
    TYPE_DATETIME,
    TYPE_DOUBLE,
    TYPE_FLOAT,
    TYPE_GUID,
    TYPE_INT,
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_MONEY,
    TYPE_NAMES,
    TYPE_OLE,
    TYPE_TEXT,
)
from pyopenvba.access_read import AccessError

if TYPE_CHECKING:
    from pyopenvba.access.database import AccessDatabase

#: Jet's DDL type words, measured one CREATE TABLE at a time.
DDL_TYPES: dict[str, int] = {
    "binary": TYPE_BINARY,
    "varbinary": TYPE_BINARY,
    "bit": TYPE_BOOLEAN,
    "yesno": TYPE_BOOLEAN,
    "logical": TYPE_BOOLEAN,
    "logical1": TYPE_BOOLEAN,
    "byte": TYPE_BYTE,
    "integer1": TYPE_BYTE,
    "counter": TYPE_LONG,
    "autoincrement": TYPE_LONG,
    "currency": TYPE_MONEY,
    "money": TYPE_MONEY,
    "datetime": TYPE_DATETIME,
    "date": TYPE_DATETIME,
    "time": TYPE_DATETIME,
    "timestamp": TYPE_DATETIME,
    "double": TYPE_DOUBLE,
    "float": TYPE_DOUBLE,
    "float8": TYPE_DOUBLE,
    "ieeedouble": TYPE_DOUBLE,
    "number": TYPE_DOUBLE,
    "single": TYPE_FLOAT,
    "float4": TYPE_FLOAT,
    "ieeesingle": TYPE_FLOAT,
    "real": TYPE_FLOAT,
    "short": TYPE_INT,
    "integer2": TYPE_INT,
    "smallint": TYPE_INT,
    "long": TYPE_LONG,
    "int": TYPE_LONG,
    "integer": TYPE_LONG,
    "integer4": TYPE_LONG,
    "longbinary": TYPE_OLE,
    "general": TYPE_OLE,
    "oleobject": TYPE_OLE,
    "longtext": TYPE_MEMO,
    "longchar": TYPE_MEMO,
    "memo": TYPE_MEMO,
    "note": TYPE_MEMO,
    "text": TYPE_TEXT,
    "alphanumeric": TYPE_TEXT,
    "string": TYPE_TEXT,
    "varchar": TYPE_TEXT,
    "guid": TYPE_GUID,
    "bigint": TYPE_BIGINT,
}
#: Words the engine takes but this cannot write yet, and why.
REFUSED_TYPES = {
    "char": "CHAR makes a fixed-width Text column",
    "character": "CHARACTER makes a fixed-width Text column",
    "decimal": "the Jet parser has no DECIMAL; a Decimal column comes from another provider",
    "numeric": "the Jet parser has no NUMERIC; a Decimal column comes from another provider",
}
AUTONUMBER_WORDS = {"counter", "autoincrement"}

def is_ddl(text: str) -> bool:
    upper = text.upper()
    return any(upper.startswith(verb) for verb in ("CREATE ", "DROP ", "ALTER "))


def execute_ddl(db: AccessDatabase, sql: str, *, created: object | None = None, updated: object | None = None, referenced_updated: object | None = None) -> int:
    """Run one DDL statement and return 0, the row count DAO reports."""
    text = " ".join(sql.strip().rstrip(";").split())
    upper = text.upper()
    if upper.startswith("CREATE TABLE "):
        _create_table(db, text[len("CREATE TABLE ") :], created, updated, referenced_updated)
    elif upper.startswith("CREATE "):
        _create_index(db, text, updated)
    elif upper.startswith("DROP TABLE "):
        db.drop_table(_unquote(text[len("DROP TABLE ") :]))
    elif upper.startswith("DROP INDEX "):
        name, _, table = text[len("DROP INDEX ") :].partition(" ON ")
        if not table:
            raise AccessError("DROP INDEX needs ON <table>")
        db.drop_index(_unquote(table), _unquote(name), updated=updated)
    elif upper.startswith("ALTER TABLE "):
        _alter_table(db, text[len("ALTER TABLE ") :], created, updated, referenced_updated)
    else:
        raise AccessError(f"{text.split()[0]} is not a statement this can run")
    return 0


def _unquote(name: str) -> str:
    name = name.strip()
    if name.startswith("[") and name.endswith("]"):
        return name[1:-1]
    return name.strip("`").strip()


def _split_parenthesized(text: str) -> tuple[str, str]:
    """The head before a top-level ``(...)`` and what is inside it."""
    start = text.find("(")
    if start < 0 or not text.rstrip().endswith(")"):
        raise AccessError(f"expected a bracketed list in {text!r}")
    return text[:start].strip(), text[start + 1 : text.rstrip().rfind(")")].strip()


def column_spec(text: str) -> tuple[ColumnSpec, list[tuple[str, str, list[str], tuple[str, list[str]] | None]]]:
    """One column definition: its spec and the constraints written on it.
    ``NOT NULL`` is accepted and changes nothing, which is what the engine
    does with it (measured: a NOT NULL column's header is byte for byte
    the header of a nullable one)."""
    match = re.match(r"^\s*(\[[^\]]+\]|\w+)\s+(.*)$", text.strip(), re.DOTALL)
    if not match:
        raise AccessError(f"cannot read the column {text.strip()!r}")
    name = _unquote(match.group(1))
    rest = match.group(2).strip()
    type_match = re.match(r"^(\w+)\s*(\(([^)]*)\))?\s*(.*)$", rest, re.DOTALL)
    if not type_match:
        raise AccessError(f"column {name!r}: cannot read its type")
    word = type_match.group(1).lower()
    size_text = (type_match.group(3) or "").strip()
    tail = (type_match.group(4) or "").strip().upper()
    if word in REFUSED_TYPES:
        raise AccessError(f"column {name!r}: {REFUSED_TYPES[word]}")
    code = DDL_TYPES.get(word)
    if code is None:
        raise AccessError(f"column {name!r}: unknown type {type_match.group(1)!r}")
    size: int | None = None
    if size_text:
        try:
            size = int(size_text.split(",")[0])
        except ValueError as exc:
            raise AccessError(f"column {name!r}: cannot read the size {size_text!r}") from exc
    if "WITH COMPRESSION" in tail or "WITH COMP" in tail:
        raise AccessError(f"column {name!r}: the Jet parser refuses WITH COMPRESSION")
    spec = ColumnSpec(
        name=name,
        type=TYPE_NAMES[code].lower(),
        size=size,
        autonumber=word in AUTONUMBER_WORDS,
        # CREATE TABLE leaves Unicode compression off, where the Access
        # window turns it on: measured on the engine's own tables.
        compressed=False,
    )
    return spec, _column_constraints(name, type_match.group(4) or "")


def _column_constraints(column: str, tail: str) -> list[tuple[str, str, list[str], tuple[str, list[str]] | None]]:
    """The constraints written after a column's type.  A named one keeps
    its name; an unnamed one gets a made-up name, where the engine makes
    up a random ``Index_...`` instead, so those two cannot agree."""
    out: list[tuple[str, str, list[str], tuple[str, list[str]] | None]] = []
    rest = tail.strip()
    while rest:
        match = re.match(r"^(?:NOT\s+NULL|NULL)\s*", rest, re.IGNORECASE)
        if match:
            rest = rest[match.end() :]
            continue
        named = re.match(r"^CONSTRAINT\s+(\[[^\]]+\]|\w+)\s+", rest, re.IGNORECASE)
        name = _unquote(named.group(1)) if named else ""
        body = rest[named.end() :] if named else rest
        upper = body.upper()
        if upper.startswith("PRIMARY KEY"):
            out.append(("primary", name or f"PrimaryKey_{column}", [column], None))
            rest = body[len("PRIMARY KEY") :].strip()
        elif upper.startswith("UNIQUE"):
            out.append(("unique", name or f"Unique_{column}", [column], None))
            rest = body[len("UNIQUE") :].strip()
        elif upper.startswith("REFERENCES"):
            parent, columns = _references(" REFERENCES " + body[len("REFERENCES") :])
            out.append(("foreign", name or f"FK_{column}_{parent}", [column], (parent, columns)))
            rest = ""
        elif not body.strip():
            rest = ""
        else:
            raise AccessError(f"column {column!r}: cannot read {body.strip()!r}")
    return out


def _table_constraint(text: str) -> tuple[str, str, list[str], tuple[str, list[str]] | None]:
    """A ``CONSTRAINT name PRIMARY KEY|UNIQUE|FOREIGN KEY (cols)`` clause:
    its kind, name, columns, and for a foreign key the parent it names."""
    match = re.match(r"^CONSTRAINT\s+(\[[^\]]+\]|\w+)\s+(.*)$", text.strip(), re.IGNORECASE | re.DOTALL)
    if not match:
        raise AccessError(f"cannot read the constraint {text.strip()!r}")
    name = _unquote(match.group(1))
    body = match.group(2).strip()
    upper = body.upper()
    for word, kind in (("PRIMARY KEY", "primary"), ("FOREIGN KEY", "foreign"), ("UNIQUE", "unique")):
        if upper.startswith(word):
            rest = body[len(word) :].strip()
            head, inner = _split_parenthesized(rest) if kind != "foreign" else _foreign_head(rest)
            columns = [_unquote(c) for c in split_top_level(inner, ",")]
            if kind != "foreign":
                if head:
                    raise AccessError(f"constraint {name!r}: unexpected {head!r}")
                return kind, name, columns, None
            return kind, name, columns, _references(rest)
    raise AccessError(f"constraint {name!r}: only PRIMARY KEY, UNIQUE and FOREIGN KEY are written")


def _foreign_head(rest: str) -> tuple[str, str]:
    body, _, _references_text = rest.partition(" REFERENCES ")
    return _split_parenthesized(body)


def _references(rest: str) -> tuple[str, list[str]]:
    _, _, tail = rest.partition(" REFERENCES ")
    if not tail:
        raise AccessError("a FOREIGN KEY needs REFERENCES <table> (<columns>)")
    head, inner = _split_parenthesized(tail)
    return _unquote(head), [_unquote(c) for c in split_top_level(inner, ",")]


def _create_table(db: AccessDatabase, body: str, created: object | None, updated: object | None, referenced_updated: object | None = None) -> None:
    name, inner = _split_parenthesized(body)
    columns: list[ColumnSpec] = []
    indexes: list[IndexSpec] = []
    foreign: list[tuple[str, list[str], str, list[str]]] = []
    for item in split_top_level(inner, ","):
        item = item.strip()
        if item.upper().startswith("CONSTRAINT "):
            kind, constraint, cols, parent = _table_constraint(item)
            if kind == "foreign" and parent is not None:
                foreign.append((constraint, cols, parent[0], parent[1]))
            else:
                indexes.append(IndexSpec(constraint, tuple(cols), unique=True, primary=kind == "primary"))
            continue
        spec, constraints = column_spec(item)
        columns.append(spec)
        for kind, constraint, cols, parent in constraints:
            if kind == "foreign" and parent is not None:
                foreign.append((constraint, cols, parent[0], parent[1]))
            else:
                indexes.append(IndexSpec(constraint, tuple(cols), unique=True, primary=kind == "primary"))
    db.create_table(_unquote(name), columns, indexes, created=created, updated=updated)
    for constraint, cols, parent, parent_columns in foreign:
        db.create_relationship(
            constraint, _unquote(name), tuple(cols), parent, tuple(parent_columns),
            created=created, table_updated=updated, referenced_updated=referenced_updated if referenced_updated is not None else updated,
        )


def _create_index(db: AccessDatabase, text: str, updated: object | None) -> None:
    match = re.match(r"^CREATE\s+(UNIQUE\s+)?INDEX\s+(\[[^\]]+\]|\w+)\s+ON\s+(.*)$", text, re.IGNORECASE | re.DOTALL)
    if not match:
        raise AccessError(f"cannot read {text!r}; expected CREATE [UNIQUE] INDEX <name> ON <table> (<columns>)")
    unique = bool(match.group(1))
    name = _unquote(match.group(2))
    table, inner = _split_parenthesized(match.group(3).split(" WITH ")[0])
    tail = match.group(3).upper()
    columns: list[str | tuple[str, bool]] = []
    for item in split_top_level(inner, ","):
        item = item.strip()
        if item.upper().endswith(" DESC"):
            columns.append((_unquote(item[:-5]), False))
        elif item.upper().endswith(" ASC"):
            columns.append((_unquote(item[:-4]), True))
        else:
            columns.append(_unquote(item))
    spec = IndexSpec(
        name,
        tuple(columns),
        unique=unique or "WITH PRIMARY" in tail,
        primary="WITH PRIMARY" in tail,
        ignore_nulls="IGNORE NULL" in tail,
        required="DISALLOW NULL" in tail,
    )
    db.create_index(_unquote(table), spec, updated=updated)


def _alter_table(db: AccessDatabase, body: str, created: object | None, updated: object | None, referenced_updated: object | None = None) -> None:
    match = re.match(r"^(\[[^\]]+\]|\w+)\s+(.*)$", body.strip(), re.DOTALL)
    if not match:
        raise AccessError(f"cannot read ALTER TABLE {body.strip()!r}")
    table = db.table(_unquote(match.group(1)))
    action = match.group(2).strip()
    upper = action.upper()
    if upper.startswith("ADD CONSTRAINT "):
        kind, name, columns, parent = _table_constraint(action[len("ADD ") :])
        if kind == "foreign" and parent is not None:
            db.create_relationship(
                name, table.name, tuple(columns), parent[0], tuple(parent[1]),
                created=created, table_updated=updated, referenced_updated=referenced_updated if referenced_updated is not None else updated,
            )
            return
        db.create_index(table.name, IndexSpec(name, tuple(columns), unique=True, primary=kind == "primary"), updated=updated)
        return
    if upper.startswith("DROP CONSTRAINT "):
        db.drop_relationship(_unquote(action[len("DROP CONSTRAINT ") :]), table_updated=updated, referenced_updated=referenced_updated if referenced_updated is not None else updated)
        return
    for word, method in (("ADD COLUMN ", "add"), ("ALTER COLUMN ", "alter")):
        if upper.startswith(word):
            spec, constraints = column_spec(action[len(word) :])
            if constraints:
                raise AccessError("a key on an added column is a separate CONSTRAINT")
            if method == "add":
                table.add_column(spec, updated=updated)
            else:
                table.alter_column(spec.name, spec, updated=updated)
            return
    if upper.startswith("DROP COLUMN "):
        table.drop_column(_unquote(action[len("DROP COLUMN ") :]), updated=updated)
        return
    if upper.startswith("ADD "):
        spec, _constraints = column_spec(action[len("ADD ") :])
        table.add_column(spec, updated=updated)
        return
    raise AccessError(f"ALTER TABLE {action.split()[0] if action else ''} is not a statement this can run")
