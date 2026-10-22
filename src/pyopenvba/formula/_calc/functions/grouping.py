"""GROUPBY and PIVOTBY: an array summarized by the values in some of its
columns, a function applied to the values of each group.

Measured against Excel 365, the formulas entered as dynamic arrays:

- A group's values reach the function as a column, blanks kept blank. A
  function with two required parameters, PERCENTOF among them, gets
  every row's values as its second, or in PIVOTBY the rows relative_to
  names: the column's (0), the row's (1), all of them (2), or the parent
  column's or row's (3, 4). One that gives an array, even of one item,
  makes the whole result ``#VALUE!``; an error it gives stays in its
  cell. A PIVOTBY cell no row falls in is empty text, the function not
  called.
- Keys match as SORT compares them, case never mattering and a blank
  apart from the empty text, and sort numbers, text, logicals, errors,
  with blanks last whichever way the key sorts. A key shows the spelling
  of the first row of its group's first leaf, the leaves in key order.
- Sorting by a result column orders every level of rows by its own
  group's result, ties kept in key order; in PIVOTBY only while there is
  a total column, and across the top only the groups with a column of
  their own, the leaves and those with subtotals.
- Headers, when the argument leaves them to Excel, are taken to be there
  when the first value column starts with text and holds numbers below
  it, and then any other text in that column is ``#VALUE!``. PIVOTBY
  shows the column fields' names joined into one text.
- Functions stacked one below another over several value columns name
  the value column beside the function, when headers are shown.
- One function over several value columns puts PIVOTBY's column totals
  in the wrong columns. Each total is written from the start of its
  group's columns, one column along for every group before it there, as
  if each group took one column, the grand total first; the leaves are
  written after, over them. What survives stands where Excel shows it,
  and the totals' own columns are otherwise empty text.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arrays import sort_key
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.registry import A, V, function
from pyopenvba.formula._calc.values import (
    EMPTY,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Lambda,
    Reference,
    Scalar,
    Value,
)
from pyopenvba.formula._calc.cells import CellError

Key = tuple[int, object]
Row = list[Scalar]

#: Where a blank sorts among keys: last, as sort_key has it.
_BLANK: Key = sort_key(EMPTY)

#: The key of a tree's root, which no value sorts to.
_ROOT: Key = (-1, None)


def _given(value: Value | None) -> bool:
    return value is not None and not isinstance(value, Empty)


def _blank(count: int) -> Row:
    """Empty text, ``count`` cells of it."""
    return [""] * count


# ----------------------------------------------------------------------
# Arguments
# ----------------------------------------------------------------------


def _whole(context: Context, value: Scalar | None, default: int) -> int:
    """An argument that has to be a whole number, or its default when left
    out."""
    if value is None or isinstance(value, Empty):
        return default
    number = context.number(value)
    if not number.is_integer():
        raise ExcelError(VALUE)
    return int(number)


def _depth(context: Context, value: Scalar | None, fields: int) -> int:
    """A total depth: 0 for no totals, 1 for a grand total, up to the
    number of fields for subtotals too, negative to put them first."""
    depth = _whole(context, value, 1)
    if not (abs(depth) <= 1 or 2 <= abs(depth) <= fields):
        raise ExcelError(VALUE)
    return depth


def _total(depth: int) -> str:
    """The grand total's label: Grand Total when there are subtotals."""
    return "Total" if abs(depth) == 1 else "Grand Total"


def _functions(value: Value) -> tuple[list[Lambda], bool]:
    """The functions to apply, and whether they were stacked in a column,
    one below another, as VSTACK(SUM,MAX) stacks them."""
    if isinstance(value, Lambda):
        items, stacked = [value], False
    elif not isinstance(value, Array):
        raise ExcelError(VALUE)
    elif value.height == 1:
        items, stacked = list(value.rows[0]), False
    elif value.width == 1:
        items, stacked = [row[0] for row in value.rows], True
    else:
        raise ExcelError(VALUE)
    found = [item for item in items if isinstance(item, Lambda)]
    # Measured: a function that needs three arguments is refused whole.
    if len(found) != len(items) or any(item.required > 2 for item in found):
        raise ExcelError(VALUE)
    return found, stacked and len(found) > 1


def _name(call: Lambda) -> str:
    """How a result is labelled: the function's own name, or CUSTOM for a
    LAMBDA."""
    return call.builtin if call.builtin is not None else "CUSTOM"


def _headers(context: Context, value: Scalar | None, values: Array) -> tuple[bool, bool | None]:
    """Whether the first row holds headers, and whether to show them: True,
    False, or None when Excel decides, which it does to show them only with
    more than one function side by side."""
    if not _given(value):
        column = [row[0] for row in values.rows]
        if not isinstance(column[0], str) or not any(isinstance(item, float) for item in column[1:]):
            return False, False
        if any(isinstance(item, str) for item in column[1:]):
            raise ExcelError(VALUE)
        return True, None
    mode = _whole(context, value, 0)
    if mode not in (0, 1, 2, 3):
        raise ExcelError(VALUE)
    return mode in (1, 3), mode in (2, 3)


def _kept(context: Context, value: Value | None, height: int, skipped: int) -> list[int]:
    """The data rows a filter keeps, given as a column as long as the data,
    or as the data and its header row, or as one value for every row."""
    rows = height - skipped
    if not _given(value):
        return list(range(rows))
    assert value is not None
    grid = matrix(context, value)
    if grid.width != 1:
        raise ExcelError(VALUE)
    flags = [row[0] for row in grid.rows]
    if len(flags) == 1:
        flags = flags * rows
    elif len(flags) == height:
        flags = flags[skipped:]
    elif len(flags) != rows:
        raise ExcelError(VALUE)
    kept: list[int] = []
    for index, flag in enumerate(flags):
        if isinstance(flag, CellError):
            raise ExcelError(VALUE)
        if not isinstance(flag, Empty) and context.logical(flag):
            kept.append(index)
    if not kept:
        raise ExcelError(VALUE)
    return kept


@dataclass(frozen=True)
class _Order:
    """How to sort: the key columns named, in order, each with whether it
    sorts descending; or one result column and its direction."""

    columns: list[tuple[int, bool]]
    by_result: tuple[int, bool] | None

    def descending(self, column: int) -> bool:
        return any(named == column and flag for named, flag in self.columns)


def _order(context: Context, value: Value | None, keys: int, results: int) -> _Order:
    """sort_order: one column, key or result, or a row of key columns; a
    negative number sorts that column descending. Measured: a column of
    numbers, a number twice, 0, a fraction, or a result column among
    others is #VALUE!."""
    if not _given(value):
        return _Order([], None)
    assert value is not None
    grid = matrix(context, value)
    if grid.height != 1:
        raise ExcelError(VALUE)
    found: list[tuple[int, bool]] = []
    for item in grid.rows[0]:
        index = _whole(context, item, 0)
        if index == 0 or abs(index) > keys + results:
            raise ExcelError(VALUE)
        found.append((abs(index) - 1, index < 0))
    named = [column for column, _ in found]
    if len(set(named)) != len(named) or (len(found) > 1 and any(column >= keys for column in named)):
        raise ExcelError(VALUE)
    if len(found) == 1 and named[0] >= keys:
        return _Order([], (named[0] - keys, found[0][1]))
    return _Order(found, None)


# ----------------------------------------------------------------------
# Groups
# ----------------------------------------------------------------------


@dataclass(eq=False)
class _Node:
    """A group: its key, its level (1 for the first key column), the data
    rows it holds in their order, and its groups one level down."""

    identity: Key
    level: int
    rows: list[int] = field(default_factory=list[int])
    children: dict[Key, _Node] = field(default_factory=dict[Key, "_Node"])


def _tree(keys: list[Row], rows: list[int]) -> _Node:
    root = _Node(_ROOT, 0)
    for row in rows:
        node = root
        node.rows.append(row)
        for level, value in enumerate(keys[row], start=1):
            identity = sort_key(value)
            child = node.children.get(identity)
            if child is None:
                child = node.children[identity] = _Node(identity, level)
            child.rows.append(row)
            node = child
    return root


def _parents(root: _Node) -> dict[int, _Node]:
    """Each group's parent, by the group's id."""
    found: dict[int, _Node] = {}
    waiting = [root]
    while waiting:
        node = waiting.pop()
        for child in node.children.values():
            found[id(child)] = node
            waiting.append(child)
    return found


def _by_key(nodes: list[_Node], descending: bool) -> list[_Node]:
    """Groups in key order, blanks last whichever way the rest sort."""
    filled = sorted(
        (node for node in nodes if node.identity != _BLANK), key=lambda node: node.identity, reverse=descending
    )
    return filled + [node for node in nodes if node.identity == _BLANK]


def _by_result(nodes: list[_Node], result: Callable[[_Node], Scalar], descending: bool) -> list[_Node]:
    """Groups in the order of one result, ties kept in key order."""
    ordered = _by_key(nodes, False)
    ordered.sort(key=lambda node: sort_key(result(node)), reverse=descending)
    return ordered


def _first_rows(root: _Node, order: _Order | None) -> dict[int, int]:
    """Each group's first row, which spells its key: the first data row of
    its first leaf, the leaves in key order, each key column sorting the
    way ``order`` has it, or ascending."""
    found: dict[int, int] = {}

    def visit(node: _Node) -> int:
        if not node.children:
            first = node.rows[0]
        else:
            descending = order is not None and order.descending(node.level)
            children = _by_key(list(node.children.values()), descending)
            first = visit(children[0])
            for child in children[1:]:
                visit(child)
        found[id(node)] = first
        return first

    visit(root)
    return found


def _head(keys: list[Row], first: dict[int, int], path: list[_Node], width: int) -> Row:
    """A group's keys, each as its first row spells it, then empty text out
    to the number of key columns."""
    head: Row = [keys[first[id(node)]][node.level - 1] for node in path]
    return head + _blank(width - len(path))


def _leaves(node: _Node, path: list[_Node]) -> list[tuple[list[_Node], _Node]]:
    """Every leaf with the groups from the first level down to it."""
    if not node.children:
        return [(path, node)]
    return [found for child in node.children.values() for found in _leaves(child, [*path, child])]


# ----------------------------------------------------------------------
# Results
# ----------------------------------------------------------------------


def _one(context: Context, value: Value) -> Scalar:
    """A function's result for one group: one value, or the whole summary
    is #VALUE!."""
    if isinstance(value, Reference):
        area = value.area
        if area is None or not area.is_cell:
            raise ExcelError(VALUE)
        return context.book.cell(area.sheet, area.top, area.left)
    if isinstance(value, (Array, Lambda)):
        raise ExcelError(VALUE)
    return value


def _pairs(functions: list[Lambda], columns: int) -> list[tuple[Lambda, int]]:
    """Which function goes with which value column: one function with each
    column, each function with the one column, or the nth with the nth."""
    if len(functions) == 1:
        return [(functions[0], column) for column in range(columns)]
    if columns == 1:
        return [(call, 0) for call in functions]
    if len(functions) != columns:
        raise ExcelError(VALUE)
    return list(zip(functions, range(columns), strict=True))


@dataclass
class _Grid:
    """The results where a row group meets a column group, over the data
    rows both hold, each worked out once. A function that takes a second
    argument gets the rows ``relative`` names: the column's (0), the
    row's (1), every row (2), or the parent column's or row's (3, 4).
    GROUPBY has one column, holding every row."""

    context: Context
    values: list[Row]
    pairs: list[tuple[Lambda, int]]
    relative: int
    rows: _Node
    columns: _Node
    parents: dict[int, _Node] = field(default_factory=dict[int, _Node])
    shared_rows: dict[tuple[int, int], list[int]] = field(default_factory=dict[tuple[int, int], list[int]])
    members: dict[int, frozenset[int]] = field(default_factory=dict[int, frozenset[int]])
    others: dict[tuple[int, int], Array] = field(default_factory=dict[tuple[int, int], Array])
    found: dict[tuple[int, int], Row] = field(default_factory=dict[tuple[int, int], Row])

    def shared(self, row: _Node, column: _Node) -> list[int]:
        """The data rows both groups hold, in their order."""
        if row is self.rows:
            return column.rows
        if column is self.columns:
            return row.rows
        key = (id(row), id(column))
        found = self.shared_rows.get(key)
        if found is None:
            members = self.members.get(id(column))
            if members is None:
                members = self.members[id(column)] = frozenset(column.rows)
            found = self.shared_rows[key] = [index for index in row.rows if index in members]
        return found

    def results(self, row: _Node, column: _Node) -> Row:
        """Each pair's result for the cell, or empty text where no data row
        falls."""
        key = (id(row), id(column))
        found = self.found.get(key)
        if found is None:
            rows = self.shared(row, column)
            found = self.found[key] = [
                self._apply(call, index, rows, row, column) if rows else "" for call, index in self.pairs
            ]
        return found

    def _apply(self, call: Lambda, index: int, rows: list[int], row: _Node, column: _Node) -> Scalar:
        args: list[Value] = [Array([[self.values[at][index]] for at in rows])]
        if call.required == 2:
            relative = self._relative(row, column)
            key = (id(relative), index)
            other = self.others.get(key)
            if other is None:
                other = self.others[key] = Array([[self.values[at][index]] for at in relative])
            args.append(other)
        return _one(self.context, self.context.apply(call, args))

    def _relative(self, row: _Node, column: _Node) -> list[int]:
        if self.relative == 0:
            return column.rows
        if self.relative == 1:
            return row.rows
        if self.relative == 2:
            return self.rows.rows
        if self.relative == 3:
            return self.shared(row, self.parents.get(id(column), column))
        return self.shared(self.parents.get(id(row), row), column)


def _table(
    leaves: list[tuple[list[_Node], _Node]], order: _Order, results: Callable[[_Node], Row]
) -> list[tuple[list[_Node], _Node]]:
    """Leaves sorted across the whole table, as field_relationship 1 sorts
    them: by the key columns named, in the order named, then the rest
    ascending, a descending column reversed whole, blanks and all; or by a
    result, ties in key order."""
    ordered = sorted(leaves, key=lambda item: tuple(node.identity for node in item[0]))
    if order.by_result is not None:
        column, descending = order.by_result
        ordered.sort(key=lambda item: sort_key(results(item[1])[column]), reverse=descending)
        return ordered
    for column, descending in reversed(order.columns):
        ordered.sort(key=lambda item: item[0][column].identity, reverse=descending)
    return ordered


def _labels(data: Array, has_headers: bool, pairs: list[tuple[Lambda, int]]) -> Row:
    """Each pair's value column by name: its header, or Value n."""
    return [data.rows[0][column] if has_headers else f"Value {column + 1}" for _, column in pairs]


def _field_names(fields: Array, has_headers: bool, kind: str) -> Row:
    """The key columns by name: their headers, or Row Field n or Column
    Field n, as ``kind`` says."""
    return [fields.rows[0][column] if has_headers else f"{kind} Field {column + 1}" for column in range(fields.width)]


# ----------------------------------------------------------------------
# GROUPBY
# ----------------------------------------------------------------------


@function("GROUPBY", A, A, A, V, V, A, A, V, minimum=3)
def GROUPBY(
    context: Context,
    row_fields: Value,
    values: Value,
    function_: Value,
    field_headers: Scalar | None = None,
    total_depth: Scalar | None = None,
    sort_order: Value | None = None,
    filter_array: Value | None = None,
    field_relationship: Scalar | None = None,
) -> Value:
    fields = matrix(context, row_fields)
    data = matrix(context, values)
    if fields.height != data.height:
        return VALUE
    functions, stacked = _functions(function_)
    pairs = _pairs(functions, data.width)
    has_headers, show = _headers(context, field_headers, data)
    if show is None:
        show = len(functions) > 1 and not stacked
    skipped = 1 if has_headers else 0
    kept = _kept(context, filter_array, data.height, skipped)
    width = fields.width
    depth = _depth(context, total_depth, width)
    relationship = _whole(context, field_relationship, 0)
    if relationship not in (0, 1) or (relationship == 1 and abs(depth) > 1):
        return VALUE
    order = _order(context, sort_order, width, 1 if stacked else len(pairs))
    keys = [list(row) for row in fields.rows[skipped:]]
    root = _tree(keys, kept)
    whole = _Node(_ROOT, 0, list(kept))
    grid = _Grid(context, [list(row) for row in data.rows[skipped:]], pairs, 2, root, whole)

    def results(node: _Node) -> Row:
        return grid.results(node, whole)

    first = _first_rows(root, None if relationship else order)
    named = stacked and show and data.width > 1
    labels = _labels(data, has_headers, pairs)
    body: list[Row] = []

    def emit(head: Row, node: _Node) -> None:
        """A group's row: its keys, then its results, one row a function
        when they are stacked."""
        found = results(node)
        if not stacked:
            body.append(head + found)
            return
        for number, (call, _) in enumerate(pairs):
            line: Row = [*head, _name(call)]
            if named:
                line.append(labels[number])
            line.append(found[number])
            body.append(line)

    def walk(node: _Node, path: list[_Node]) -> None:
        if not node.children:
            emit(_head(keys, first, path, width), node)
            return
        subtotal = 1 <= node.level < abs(depth)
        if subtotal and depth < 0:
            emit(_head(keys, first, path, width), node)
        children = list(node.children.values())
        if order.by_result is not None:
            column, descending = order.by_result
            children = _by_result(children, lambda child: results(child)[column], descending)
        else:
            children = _by_key(children, order.descending(node.level))
        for child in children:
            walk(child, [*path, child])
        if subtotal and depth > 0:
            emit(_head(keys, first, path, width), node)

    grand: Row = [_total(depth), *_blank(width - 1)]
    if depth < 0:
        emit(grand, root)
    if relationship:
        for path, leaf in _table(_leaves(root, []), order, results):
            emit(_head(keys, first, path, width), leaf)
    else:
        walk(root, [])
    if depth > 0:
        emit(grand, root)
    top: list[Row] = []
    if len(functions) > 1 and not stacked:
        top.append(_blank(width) + [_name(call) for call, _ in pairs])
    if show:
        names = _field_names(fields, has_headers, "Row")
        if named:
            top.append([*names, "", "", ""])
        elif stacked:
            top.append([*names, "", labels[0]])
        else:
            top.append(names + labels)
    return Array(top + body)


# ----------------------------------------------------------------------
# PIVOTBY
# ----------------------------------------------------------------------


@dataclass
class _Across:
    """PIVOTBY's columns: every group with a column, in the order Excel
    lays them out with each total after its group (``canonical``), the
    order shown, as indexes into that, and for each total the first
    column of its group's block and the order the totals are written in,
    the grand total first. ``paths`` holds each column's groups from the
    first level down."""

    canonical: list[_Node] = field(default_factory=list[_Node])
    paths: list[list[_Node]] = field(default_factory=list[list[_Node]])
    shown: list[int] = field(default_factory=list[int])
    firsts: dict[int, int] = field(default_factory=dict[int, int])
    writes: list[int] = field(default_factory=list[int])


def _across(root: _Node, depth: int, arrange: Callable[[_Node], list[_Node]]) -> _Across:
    across = _Across()

    def lay(node: _Node, path: list[_Node]) -> list[int]:
        if not node.children:
            across.canonical.append(node)
            across.paths.append(path)
            return [len(across.canonical) - 1]
        total = depth != 0 if node.level == 0 else node.level < abs(depth)
        start = len(across.canonical)
        slot = len(across.writes)
        if total:
            # Written before its groups' totals; its index comes after them.
            across.writes.append(-1)
        shown = [index for child in arrange(node) for index in lay(child, [*path, child])]
        if not total:
            return shown
        across.canonical.append(node)
        across.paths.append(path)
        index = len(across.canonical) - 1
        across.firsts[index] = start
        across.writes[slot] = index
        return [index, *shown] if depth < 0 else [*shown, index]

    across.shown = lay(root, [])
    return across


@function("PIVOTBY", A, A, A, A, V, V, A, V, A, A, V, minimum=4)
def PIVOTBY(
    context: Context,
    row_fields: Value,
    col_fields: Value,
    values: Value,
    function_: Value,
    field_headers: Scalar | None = None,
    row_total_depth: Scalar | None = None,
    row_sort_order: Value | None = None,
    col_total_depth: Scalar | None = None,
    col_sort_order: Value | None = None,
    filter_array: Value | None = None,
    relative_to: Scalar | None = None,
) -> Value:
    down = matrix(context, row_fields)
    across_fields = matrix(context, col_fields)
    data = matrix(context, values)
    if not down.height == across_fields.height == data.height:
        return VALUE
    functions, stacked = _functions(function_)
    pairs = _pairs(functions, data.width)
    has_headers, show = _headers(context, field_headers, data)
    if show is None:
        show = len(functions) > 1 and not stacked
    skipped = 1 if has_headers else 0
    kept = _kept(context, filter_array, data.height, skipped)
    row_depth = _depth(context, row_total_depth, down.width)
    column_depth = _depth(context, col_total_depth, across_fields.width)
    relative = _whole(context, relative_to, 0)
    if relative not in range(5):
        return VALUE
    row_order = _order(context, row_sort_order, down.width, 1 if stacked else len(pairs))
    column_order = _order(context, col_sort_order, across_fields.width, 1)
    row_keys = [list(row) for row in down.rows[skipped:]]
    column_keys = [list(row) for row in across_fields.rows[skipped:]]
    rows = _tree(row_keys, kept)
    columns = _tree(column_keys, kept)
    grid = _Grid(context, [list(row) for row in data.rows[skipped:]], pairs, relative, rows, columns)
    grid.parents = _parents(rows) | _parents(columns)

    def arrange_columns(node: _Node) -> list[_Node]:
        """A group's groups across the top. Sorted by a result, only those
        with a column of their own sort by it; the rest keep key order."""
        children = list(node.children.values())
        if column_order.by_result is None:
            return _by_key(children, column_order.descending(node.level))
        level = node.level + 1
        if abs(column_depth) <= level < across_fields.width:
            return _by_key(children, False)
        return _by_result(children, lambda child: grid.results(rows, child)[0], column_order.by_result[1])

    def arrange_rows(node: _Node) -> list[_Node]:
        """A group's groups down the side, sorted by a result only when
        there is a total column to sort by."""
        children = list(node.children.values())
        if row_order.by_result is None:
            return _by_key(children, row_order.descending(node.level))
        if column_depth == 0:
            return _by_key(children, False)
        column, descending = row_order.by_result
        return _by_result(children, lambda child: grid.results(child, columns)[column], descending)

    across = _across(columns, column_depth, arrange_columns)
    canonical, shown = across.canonical, across.shown
    leaves = [index for index, node in enumerate(canonical) if not node.children]
    width = len(pairs)
    stride = len(functions)

    def cells(row: _Node) -> Row:
        """A row's results across the top, placed as Excel places them:
        totals first, each from its block's start at ``stride`` columns
        a group, then the leaves over them."""
        placed = _blank(len(canonical) * width)
        for index in across.writes:
            first = across.firsts[index]
            start = first * width + (index - first) * stride
            placed[start : start + width] = grid.results(row, canonical[index])
        for index in leaves:
            placed[index * width : index * width + width] = grid.results(row, canonical[index])
        return [value for index in shown for value in placed[index * width : index * width + width]]

    named = stacked and show and data.width > 1
    labels = _labels(data, has_headers, pairs)
    body: list[Row] = []

    def emit(head: Row, row: _Node) -> None:
        if not stacked:
            body.append(head + cells(row))
            return
        for number, (call, _) in enumerate(pairs):
            line: Row = [*head, _name(call)]
            if named:
                line.append(labels[number])
            line.extend(grid.results(row, canonical[index])[number] for index in shown)
            body.append(line)

    row_first = _first_rows(rows, row_order)

    def walk(node: _Node, path: list[_Node]) -> None:
        if not node.children:
            emit(_head(row_keys, row_first, path, down.width), node)
            return
        subtotal = 1 <= node.level < abs(row_depth)
        if subtotal and row_depth < 0:
            emit(_head(row_keys, row_first, path, down.width), node)
        for child in arrange_rows(node):
            walk(child, [*path, child])
        if subtotal and row_depth > 0:
            emit(_head(row_keys, row_first, path, down.width), node)

    grand: Row = [_total(row_depth), *_blank(down.width - 1)]
    if row_depth < 0:
        emit(grand, rows)
    walk(rows, [])
    if row_depth > 0:
        emit(grand, rows)

    column_first = _first_rows(columns, column_order)

    def label(index: int, level: int) -> Scalar:
        """A column's key at one level, or the grand total's label."""
        path = across.paths[index]
        if not path:
            return _total(column_depth) if level == 0 else ""
        if level < len(path):
            return column_keys[column_first[id(path[level])]][level]
        return ""

    per = 1 if stacked else width
    lead = down.width + (1 if stacked else 0) + (1 if named else 0)
    names = _field_names(down, has_headers, "Row")
    top: list[Row] = []
    if show:
        heads = _field_names(across_fields, has_headers, "Column")
        top.append([*_blank(lead), ", ".join(context.text(item) for item in heads), *_blank(len(shown) * per - 1)])
    for level in range(across_fields.width):
        start = names + _blank(lead - down.width) if named and level == across_fields.width - 1 else _blank(lead)
        top.append(start + [label(index, level) for index in shown for _ in range(per)])
    if len(functions) > 1 and not stacked:
        top.append(_blank(lead) + [_name(call) for _ in shown for call, _ in pairs])
    if show and not named:
        top.append(names + _blank(lead - down.width) + [labels[number] for _ in shown for number in range(per)])
    return Array(top + body)
