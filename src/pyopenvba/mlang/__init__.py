"""An evaluator for M, the Power Query formula language.

    >>> from pyopenvba.mlang import evaluate_query
    >>> table = evaluate_query('''
    ...     let
    ...         Source = #table({"Item", "Qty"}, {{"a", 1}, {"b", 2}}),
    ...         Bigger = Table.SelectRows(Source, each [Qty] > 1)
    ...     in
    ...         Bigger
    ... ''')
    >>> table.columns, table.rows
    (['Item', 'Qty'], [['b', 2]])

A query that reaches a source this cannot get to says so rather than
guessing: see :func:`pyopenvba.mlang._library.missing`.  The workbook's
own wiring, which is what makes ``WorkbookQuery.Refresh`` land data on a
sheet, is in :mod:`pyopenvba.apps.excel._refresh`.
"""

from pyopenvba.mlang._eval import Scope, base_scope, evaluate, run
from pyopenvba.mlang._parse import MSyntaxError, Node, parse
from pyopenvba.mlang._values import (
    Builtin,
    Duration,
    Function,
    MError,
    MType,
    Record,
    Table,
    type_name,
)

__all__ = [
    "Builtin",
    "Duration",
    "Function",
    "MError",
    "MSyntaxError",
    "MType",
    "Node",
    "Record",
    "Scope",
    "Table",
    "base_scope",
    "evaluate",
    "evaluate_query",
    "parse",
    "run",
    "type_name",
]


def evaluate_query(formula: str, names: dict[str, object] | None = None) -> object:
    """Evaluate one query's formula, with extra names in view.

    ``names`` is how the other queries in a workbook are made visible to
    this one, and how a host adds its own sources.
    """
    return run(formula, names)
