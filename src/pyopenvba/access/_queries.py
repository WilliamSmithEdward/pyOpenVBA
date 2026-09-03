"""Saved queries: the rows of ``MSysQueries`` and the Jet SQL they encode.

A saved query is a catalog object of type 5 whose definition is a set of
MSysQueries rows keyed by ObjectId, each ``(Attribute, Order, Name1,
Name2, Expression, Flag)``.  Measured with DAO's ``CreateQueryDef`` on a
plain select, a join with DISTINCT, TOP, GROUP BY, HAVING and ORDER BY
DESC, a parameter query, DELETE, UPDATE, INSERT INTO ... SELECT, SELECT
... INTO and UNION:

    attribute   rows                                  Name1        Name2         Expression        Flag
    0           one, first                                                                          0
    255         one, second: the end marker
    1           the type, absent for a select         target table               .                 2 make-table, 3 append, 4 update, 5 delete, 9 union
    2           one per PARAMETERS entry              [name]                                        DAO type
    6           one per output column                 alias        SET / INSERT   the expression    0
                                                                   target column
    7           one per JOIN                          left table   right          the ON condition  1 inner, 2 left, 3 right
    5           one per source table                  table        alias          (UNION: a member  (UNION: Name2 X7YZ_____n)
                                                                                  SELECT verbatim)
    8           WHERE                                                             the condition
    9           one per GROUP BY expression                                                         0
    10          HAVING
    11          one per ORDER BY expression           d when DESC
    3           the select flags, last, when not 0    TOP count                                     0x01 *, 0x02 DISTINCT, 0x04 DISTINCTROW, 0x10 TOP; 3 on a UNION

``Order`` is a four-byte big-endian sequence within each attribute.  The
rows are inserted in a type-specific order (see :func:`rows_from_sql`).
Expressions are stored as written.  The catalog row's Flags is DAO's
QueryDefTypeEnum: 0 select, 32 delete, 48 update, 64 append, 80
make-table, 128 union.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError

ATTR_START = 0
ATTR_TYPE = 1
ATTR_PARAMETER = 2
ATTR_FLAGS = 3
ATTR_TABLE = 5
ATTR_COLUMN = 6
ATTR_JOIN = 7
ATTR_WHERE = 8
ATTR_GROUP = 9
ATTR_HAVING = 10
ATTR_ORDER = 11
ATTR_END = 255

QUERY_SELECT = 1
QUERY_MAKE_TABLE = 2
QUERY_APPEND = 3
QUERY_UPDATE = 4
QUERY_DELETE = 5
QUERY_CROSSTAB = 6
QUERY_PASSTHROUGH = 8
QUERY_UNION = 9

#: MSysObjects.Flags for each query type: DAO's QueryDefTypeEnum.
CATALOG_FLAGS = {QUERY_SELECT: 0, QUERY_CROSSTAB: 16, QUERY_PASSTHROUGH: 112, QUERY_DELETE: 32, QUERY_UPDATE: 48, QUERY_APPEND: 64, QUERY_MAKE_TABLE: 80, QUERY_UNION: 128}

#: A crosstab's column and group rows say what each one is for.
CROSSTAB_VALUE = 0
CROSSTAB_PIVOT = 1
CROSSTAB_HEADING = 2

FLAG_ALL_COLUMNS = 0x01
FLAG_DISTINCT = 0x02
FLAG_DISTINCTROW = 0x04
FLAG_TOP = 0x10
UNION_FLAGS = 0x03
UNION_MEMBER_PREFIX = "X7YZ_____"

JOIN_INNER = 1
JOIN_LEFT = 2
JOIN_RIGHT = 3
JOIN_WORDS = {JOIN_INNER: "INNER JOIN", JOIN_LEFT: "LEFT JOIN", JOIN_RIGHT: "RIGHT JOIN"}

# PARAMETERS type words to DAO types, and the word DAO prints for each.
PARAMETER_TYPES = {
    "BIT": 1, "YESNO": 1, "BYTE": 2, "SHORT": 3, "INTEGER": 3, "LONG": 4, "CURRENCY": 5, "SINGLE": 6,
    "IEEESINGLE": 6, "DOUBLE": 7, "IEEEDOUBLE": 7, "DATETIME": 8, "DATE": 8, "BINARY": 9, "TEXT": 10,
    "LONGBINARY": 11, "LONGTEXT": 12, "MEMO": 12, "GUID": 15, "VALUE": 12,
}
PARAMETER_WORDS = {1: "Bit", 2: "Byte", 3: "Short", 4: "Long", 5: "Currency", 6: "IEEESingle", 7: "IEEEDouble", 8: "DateTime", 9: "Binary", 10: "Text", 11: "LongBinary", 12: "LongText", 15: "Guid"}


@dataclass(frozen=True)
class QueryRow:
    attribute: int
    order: int
    name1: str | None = None
    name2: str | None = None
    expression: str | None = None
    flag: int | None = None


@dataclass
class SavedQuery:
    """A saved query: its catalog name and its MSysQueries rows in stored order."""

    name: str
    rows: list[QueryRow] = field(default_factory=lambda: [])

    def _rows(self, attribute: int) -> list[QueryRow]:
        return sorted((r for r in self.rows if r.attribute == attribute), key=lambda r: r.order)

    @property
    def type(self) -> int:
        typed = self._rows(ATTR_TYPE)
        return typed[0].flag if typed and typed[0].flag is not None else QUERY_SELECT

    @property
    def catalog_flags(self) -> int:
        return CATALOG_FLAGS.get(self.type, 0)

    @property
    def flags(self) -> int:
        flagged = self._rows(ATTR_FLAGS)
        return flagged[0].flag if flagged and flagged[0].flag is not None else 0

    @property
    def tables(self) -> list[tuple[str, str | None]]:
        """Each FROM entry as ``(name, alias)``; a derived table's name is
        its bracketed SELECT."""
        return [(r.name1 if r.name1 else f"({r.expression})", r.name2) for r in self._rows(ATTR_TABLE)]

    @property
    def connect(self) -> str | None:
        """The connect string of a pass-through query, else None."""
        typed = self._rows(ATTR_TYPE)
        return typed[0].name1 if typed and typed[0].flag == QUERY_PASSTHROUGH else None

    @property
    def sql(self) -> str:
        """The Jet SQL the rows spell, in DAO's own layout.  A pass-through
        query keeps whatever text was saved for the server."""
        kind = self.type
        if kind == QUERY_PASSTHROUGH:
            typed = self._rows(ATTR_TYPE)
            return typed[0].expression or "" if typed else ""
        if kind == QUERY_UNION:
            return " UNION ".join((r.expression or "").rstrip() for r in self._rows(ATTR_TABLE))
        parts: list[str] = []
        parameters = self._rows(ATTR_PARAMETER)
        if parameters:
            parts.append("PARAMETERS " + ", ".join(f"{p.name1} {PARAMETER_WORDS.get(p.flag or 12, 'Value')}" for p in parameters) + ";")
        columns = self._rows(ATTR_COLUMN)
        typed = self._rows(ATTR_TYPE)
        target = typed[0].name1 if typed else None
        if kind == QUERY_CROSSTAB:
            value = next((c for c in columns if c.flag == CROSSTAB_VALUE), None)
            if value is not None:
                parts.append("TRANSFORM " + (value.expression or "") + (f" AS {value.name1}" if value.name1 else ""))
        if kind == QUERY_DELETE:
            parts.append("DELETE")
        elif kind == QUERY_UPDATE:
            parts.append("UPDATE " + ", ".join(t + (f" AS {a}" if a else "") for t, a in self.tables))
            parts.append("SET " + ", ".join(f"{c.name2} = {c.expression}" for c in columns))
        elif kind == QUERY_APPEND:
            parts.append(f"INSERT INTO {target} ( " + ", ".join(c.name2 or "" for c in columns) + " )")
        if kind in (QUERY_SELECT, QUERY_APPEND, QUERY_MAKE_TABLE, QUERY_CROSSTAB):
            head = "SELECT"
            if self.flags & FLAG_DISTINCTROW:
                head += " DISTINCTROW"
            if self.flags & FLAG_DISTINCT:
                head += " DISTINCT"
            if self.flags & FLAG_TOP:
                head += " TOP " + (self._rows(ATTR_FLAGS)[0].name1 or "")
            listed = [c for c in columns if kind != QUERY_CROSSTAB or c.flag == CROSSTAB_HEADING]
            if listed:
                head += " " + ", ".join((c.expression or "") + (f" AS {c.name1}" if c.name1 else "") for c in listed)
            else:
                head += " *"
            if kind == QUERY_MAKE_TABLE:
                head += f" INTO {target}"
            parts.append(head)
        if kind != QUERY_UPDATE:
            joins = self._rows(ATTR_JOIN)
            tables = self.tables
            if joins:
                joined = [_named(tables, tables[0][1] or tables[0][0])]
                for j in joins:
                    joined.append(f"{JOIN_WORDS.get(j.flag or JOIN_INNER, 'INNER JOIN')} {_named(tables, j.name2 or '')} ON {j.expression}")
                parts.append("FROM " + " ".join(joined))
            elif tables:
                parts.append("FROM " + ", ".join(t + (f" AS {alias}" if alias else "") for t, alias in tables))
        for attribute, word in ((ATTR_WHERE, "WHERE"), (ATTR_GROUP, "GROUP BY"), (ATTR_HAVING, "HAVING"), (ATTR_ORDER, "ORDER BY")):
            rows = [r for r in self._rows(attribute) if kind != QUERY_CROSSTAB or attribute != ATTR_GROUP or r.flag != CROSSTAB_PIVOT]
            if rows:
                items = [(r.expression or "") + (" DESC" if attribute == ATTR_ORDER and r.name1 == "d" else "") for r in rows]
                parts.append(f"{word} " + ", ".join(items))
        if kind == QUERY_CROSSTAB:
            pivot = next((c for c in columns if c.flag == CROSSTAB_PIVOT), None)
            if pivot is not None:
                parts.append("PIVOT " + (pivot.expression or ""))
        return " ".join(parts)


def _named(tables: list[tuple[str, str | None]], reference: str) -> str:
    """``Table AS alias`` for the table a join names by its alias."""
    for name, alias in tables:
        if (alias or name).lower() == reference.lower():
            return name + (f" AS {alias}" if alias else "")
    return reference


def rows_from_sql(sql: str) -> list[QueryRow]:
    """The MSysQueries rows DAO writes for ``sql``, in its insertion order,
    which depends on the statement: a select puts its columns before its
    tables and the flags row last; DELETE and UPDATE put the type row and
    the table before the columns; INSERT INTO puts the type row first and
    the source tables after the columns; SELECT INTO puts the type row
    after the columns; UNION stores each member SELECT verbatim."""
    text = sql.strip().rstrip(";").strip()
    rows: list[QueryRow] = [QueryRow(ATTR_START, 1, flag=0), QueryRow(ATTR_END, 1)]
    parameters: list[QueryRow] = []
    if text.upper().startswith("PARAMETERS "):
        head, _, text = text.partition(";")
        for i, item in enumerate(split_top_level(head[len("PARAMETERS "):], ","), start=1):
            name, _, kind = item.strip().rpartition(" ")
            dao_type = PARAMETER_TYPES.get(kind.strip().upper())
            if not name or dao_type is None:
                raise AccessError(f"cannot read parameter {item.strip()!r}")
            parameters.append(QueryRow(ATTR_PARAMETER, i, name1=name.strip(), flag=dao_type))
        text = text.strip()
    members = split_top_level_words(text, "UNION")
    if len(members) > 1:
        if parameters:
            raise AccessError("PARAMETERS on a UNION query are not written")
        for i, member in enumerate(members, start=1):
            rows.append(QueryRow(ATTR_TABLE, i, name2=f"{UNION_MEMBER_PREFIX}{i}", expression=member))
            if i == 1:
                rows.append(QueryRow(ATTR_TYPE, 1, flag=QUERY_UNION))
        rows.append(QueryRow(ATTR_FLAGS, 1, flag=UNION_FLAGS))
        return rows
    clauses = split_clauses(text)
    verb = clauses[0][0]
    if verb == "TRANSFORM":
        return _crosstab_rows(rows, parameters, clauses)
    if verb == "UPDATE":
        rows.append(QueryRow(ATTR_TYPE, 1, flag=QUERY_UPDATE))
        rows.extend(parameters)
        tables, joins = parse_from(clauses[0][1])
        if joins:
            raise AccessError("UPDATE over a join is not written")
        rows.extend(tables)
        set_clause = next((b for w, b in clauses if w == "SET"), None)
        if set_clause is None:
            raise AccessError("UPDATE needs a SET clause")
        for i, item in enumerate(split_top_level(set_clause, ","), start=1):
            column, _, expression = item.partition("=")
            rows.append(QueryRow(ATTR_COLUMN, i, name2=column.strip(), expression=expression.strip(), flag=0))
        _append_tail(rows, clauses)
        return rows
    if verb == "DELETE":
        rows.append(QueryRow(ATTR_TYPE, 1, flag=QUERY_DELETE))
        rows.extend(parameters)
        body = clauses[0][1].strip()
        if body and body != "*":
            raise AccessError("DELETE takes no column list")
        from_clause = _clause(clauses, "FROM")
        tables, joins = parse_from(from_clause)
        rows.extend(joins)
        rows.extend(tables)
        _append_tail(rows, clauses)
        return rows
    if verb == "INSERT INTO":
        target, _, column_list = clauses[0][1].partition("(")
        if not column_list:
            raise AccessError("INSERT INTO needs a column list and a SELECT")
        targets = [c.strip() for c in split_top_level(column_list.rsplit(")", 1)[0], ",")]
        select = _clause(clauses, "SELECT")
        rows.append(QueryRow(ATTR_TYPE, 1, name1=target.strip(), flag=QUERY_APPEND))
        rows.extend(parameters)
        flags, top, expressions = select_list(select)
        if len(expressions) != len(targets):
            raise AccessError("INSERT INTO lists a different number of columns and expressions")
        for i, ((expression, _alias), column) in enumerate(zip(expressions, targets, strict=True), start=1):
            rows.append(QueryRow(ATTR_COLUMN, i, name2=column, expression=expression, flag=0))
        tables, joins = parse_from(_clause(clauses, "FROM"))
        rows.extend(joins)
        rows.extend(tables)
        _append_tail(rows, clauses)
        if flags:
            rows.append(QueryRow(ATTR_FLAGS, 1, name1=top, flag=flags))
        return rows
    if verb != "SELECT":
        raise AccessError(f"only SELECT, UPDATE, DELETE, INSERT INTO and UNION queries are written; got {verb}")
    rows.extend(parameters)
    flags, top, expressions = select_list(clauses[0][1])
    for i, (expression, alias) in enumerate(expressions, start=1):
        rows.append(QueryRow(ATTR_COLUMN, i, name1=alias, expression=expression, flag=0))
    into = next((b for w, b in clauses if w == "INTO"), None)
    if into is not None:
        rows.append(QueryRow(ATTR_TYPE, 1, name1=into.strip(), flag=QUERY_MAKE_TABLE))
    tables, joins = parse_from(_clause(clauses, "FROM"))
    rows.extend(joins)
    rows.extend(tables)
    _append_tail(rows, clauses)
    if flags:
        rows.append(QueryRow(ATTR_FLAGS, 1, name1=top, flag=flags))
    return rows


def _crosstab_rows(rows: list[QueryRow], parameters: list[QueryRow], clauses: list[tuple[str, str]]) -> list[QueryRow]:
    """A crosstab's rows in DAO's insertion order: the parameters, the type
    row, the value column and the row headings, the joins and tables, the
    tail, and last of all the pivot -- its group row, then its column.  The
    engine refuses HAVING on a crosstab, so this does too."""
    if any(word == "HAVING" for word, _ in clauses):
        raise AccessError("a crosstab query takes no HAVING clause")
    pivot = next((b for w, b in clauses if w == "PIVOT"), None)
    if pivot is None:
        raise AccessError("a TRANSFORM query needs a PIVOT clause")
    pivot = pivot.strip()
    rows.extend(parameters)
    rows.append(QueryRow(ATTR_TYPE, 1, flag=QUERY_CROSSTAB))
    value, alias = split_alias(clauses[0][1].strip())
    rows.append(QueryRow(ATTR_COLUMN, 1, name1=alias, expression=value, flag=CROSSTAB_VALUE))
    flags, top, expressions = select_list(_clause(clauses, "SELECT"))
    column_order = 1
    for expression, item_alias in expressions:
        column_order += 1
        rows.append(QueryRow(ATTR_COLUMN, column_order, name1=item_alias, expression=expression, flag=CROSSTAB_HEADING))
    tables, joins = parse_from(_clause(clauses, "FROM"))
    rows.extend(joins)
    rows.extend(tables)
    _append_tail(rows, clauses, group_flag=CROSSTAB_HEADING)
    # The pivot goes on the end: its group row carries the expression alone,
    # its column row the whole clause, IN list and all.
    group_order = sum(1 for r in rows if r.attribute == ATTR_GROUP) + 1
    rows.append(QueryRow(ATTR_GROUP, group_order, expression=_pivot_expression(pivot), flag=CROSSTAB_PIVOT))
    rows.append(QueryRow(ATTR_COLUMN, column_order + 1, expression=pivot, flag=CROSSTAB_PIVOT))
    if flags:
        rows.append(QueryRow(ATTR_FLAGS, 1, name1=top, flag=flags))
    return rows


def _pivot_expression(pivot: str) -> str:
    """A PIVOT clause without its ``IN (...)`` list."""
    for item in split_top_level_words(pivot, "IN"):
        return item.strip()
    return pivot.strip()


def select_list(body: str) -> tuple[int, str | None, list[tuple[str, str | None]]]:
    """Flags, TOP count and ``(expression, alias)`` pairs of a select list."""
    body = body.strip()
    flags = 0
    top: str | None = None
    while True:
        upper = body.upper()
        if upper.startswith("DISTINCTROW "):
            flags |= FLAG_DISTINCTROW
            body = body[len("DISTINCTROW "):].lstrip()
        elif upper.startswith("DISTINCT "):
            flags |= FLAG_DISTINCT
            body = body[len("DISTINCT "):].lstrip()
        elif upper.startswith("TOP "):
            match = re.match(r"TOP\s+(\d+)(\s+PERCENT)?\s+", body, re.IGNORECASE)
            if not match:
                raise AccessError("TOP needs a count")
            flags |= FLAG_TOP
            # PERCENT is kept with the count, which is how the engine
            # writes it back and what the executor reads.
            top = match.group(1) + (" PERCENT" if match.group(2) else "")
            body = body[match.end():]
        else:
            break
    if body == "*":
        return flags | FLAG_ALL_COLUMNS, top, []
    if not body:
        raise AccessError("the select list is empty")
    return flags, top, [split_alias(item.strip()) for item in split_top_level(body, ",")]


def _clause(clauses: list[tuple[str, str]], word: str) -> str:
    body = next((b for w, b in clauses if w == word), None)
    if body is None:
        raise AccessError(f"the query has no {word} clause")
    return body


def _append_tail(rows: list[QueryRow], clauses: list[tuple[str, str]], group_flag: int = 0) -> None:
    for word, attribute in (("WHERE", ATTR_WHERE), ("GROUP BY", ATTR_GROUP), ("HAVING", ATTR_HAVING), ("ORDER BY", ATTR_ORDER)):
        body = next((b for w, b in clauses if w == word), None)
        if body is None:
            continue
        if attribute in (ATTR_GROUP, ATTR_ORDER):
            for i, item in enumerate(split_top_level(body, ","), start=1):
                expression = item.strip()
                direction: str | None = None
                if attribute == ATTR_ORDER:
                    upper = expression.upper()
                    if upper.endswith(" DESC"):
                        expression, direction = expression[:-5].rstrip(), "d"
                    elif upper.endswith(" ASC"):
                        expression = expression[:-4].rstrip()
                rows.append(QueryRow(attribute, i, name1=direction, expression=expression, flag=group_flag if attribute == ATTR_GROUP else None))
        else:
            rows.append(QueryRow(attribute, 1, expression=body.strip()))


_CLAUSE_WORDS = ("TRANSFORM", "SELECT", "DELETE", "UPDATE", "INSERT INTO", "SET", "INTO", "FROM", "WHERE", "GROUP BY", "HAVING", "ORDER BY", "PIVOT")


def split_clauses(text: str) -> list[tuple[str, str]]:
    """Split a statement at its top-level clause keywords."""
    positions: list[tuple[int, int, str]] = []
    depth = 0
    quote: str | None = None
    i = 0
    upper = text.upper()
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
        elif depth == 0 and (i == 0 or not (text[i - 1].isalnum() or text[i - 1] in "_.")):
            for word in _CLAUSE_WORDS:
                end = i + len(word)
                if upper.startswith(word, i) and (end == len(text) or not (text[end].isalnum() or text[end] == "_")):
                    positions.append((i, end, word))
                    i = end - 1
                    break
        i += 1
    if not positions or positions[0][0] != 0:
        raise AccessError("the query must start with SELECT, UPDATE, DELETE or INSERT INTO")
    out: list[tuple[str, str]] = []
    for n, (_start, end, word) in enumerate(positions):
        stop = positions[n + 1][0] if n + 1 < len(positions) else len(text)
        out.append((word, text[end:stop].strip()))
    return out


def split_top_level(text: str, separator: str) -> list[str]:
    parts: list[str] = []
    depth = 0
    quote: str | None = None
    start = 0
    for i, ch in enumerate(text):
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
        elif ch == separator and depth == 0:
            parts.append(text[start:i])
            start = i + 1
    parts.append(text[start:])
    return [p for p in parts if p.strip()]


def split_top_level_words(text: str, word: str) -> list[str]:
    """Split at a top-level keyword, keeping each member's own spacing
    (DAO keeps the space before UNION on the member it ends)."""
    parts: list[str] = []
    depth = 0
    quote: str | None = None
    start = 0
    i = 0
    upper = text.upper()
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
        elif depth == 0 and upper.startswith(word, i) and i > 0 and text[i - 1] == " " and (i + len(word) == len(text) or text[i + len(word)] == " "):
            parts.append(text[start:i])
            start = i + len(word) + 1
            i += len(word)
        i += 1
    parts.append(text[start:])
    return parts


def split_alias(item: str) -> tuple[str, str | None]:
    """``Table AS t``, ``Table t`` or plain ``Table``."""
    match = re.search(r"\s+AS\s+(\[[^\]]+\]|\w+)\s*$", item, re.IGNORECASE)
    if match:
        return item[: match.start()].rstrip(), match.group(1)
    match = re.fullmatch(r"(\[[^\]]+\]|[\w.]+)\s+(\[[^\]]+\]|\w+)", item.strip())
    if match:
        return match.group(1), match.group(2)
    return item, None


def parse_from(clause: str) -> tuple[list[QueryRow], list[QueryRow]]:
    """Tables (attribute 5) and joins (attribute 7) of a FROM clause."""
    tables: list[QueryRow] = []
    joins: list[QueryRow] = []
    join_pattern = re.compile(r"\s+(INNER|LEFT|RIGHT)\s+JOIN\s+", re.IGNORECASE)
    for source in split_top_level(clause, ","):
        pieces = join_pattern.split(source.strip())
        first, alias = split_alias(pieces[0].strip())
        tables.append(_table_row(len(tables) + 1, first, alias))
        left = alias or first
        for k in range(1, len(pieces), 2):
            kind = {"INNER": JOIN_INNER, "LEFT": JOIN_LEFT, "RIGHT": JOIN_RIGHT}[pieces[k].upper()]
            right_part, condition = _split_on(pieces[k + 1])
            right, right_alias = split_alias(right_part.strip())
            tables.append(_table_row(len(tables) + 1, right, right_alias))
            joins.append(QueryRow(ATTR_JOIN, len(joins) + 1, name1=left, name2=right_alias or right, expression=condition.strip(), flag=kind))
            left = right_alias or right
    return tables, joins


def _table_row(order: int, name: str, alias: str | None) -> QueryRow:
    """A FROM entry: a table by name, or a bracketed SELECT whose text goes
    in the expression with only the alias naming it, as the engine
    stores it."""
    if name.startswith("("):
        return QueryRow(ATTR_TABLE, order, name2=alias, expression=name[1 : name.rfind(")")].strip())
    return QueryRow(ATTR_TABLE, order, name1=name, name2=alias)


def _split_on(text: str) -> tuple[str, str]:
    """A joined table and its ON condition."""
    match = re.search(r"\s+ON\s+", text, re.IGNORECASE)
    if not match:
        raise AccessError(f"a JOIN needs an ON condition: {text.strip()!r}")
    return text[: match.start()], text[match.end():]
