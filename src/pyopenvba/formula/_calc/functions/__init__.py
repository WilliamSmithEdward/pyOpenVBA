"""The worksheet functions, one module per family.

Importing this package registers every function in
:data:`~pyopenvba.formula._calc.registry.FUNCTIONS`.
"""

from __future__ import annotations

from pyopenvba.formula._calc.functions import (
    aggregate,
    arithmetic,
    arrays,
    conditional,
    database,
    dates,
    distributions,
    engineering,
    financial,
    information,
    lambdas,
    logical,
    lookup,
    matrix,
    statistics,
    subtotal,
    text,
)

__all__ = [
    "aggregate",
    "arithmetic",
    "arrays",
    "conditional",
    "database",
    "dates",
    "distributions",
    "engineering",
    "financial",
    "information",
    "lambdas",
    "logical",
    "lookup",
    "matrix",
    "statistics",
    "subtotal",
    "text",
]
