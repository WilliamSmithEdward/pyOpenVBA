"""Formulas as Excel reads and calculates them: pyOfficeEditor's formula engine.

This package is a copy of ``pyofficeeditor.excel._calc`` from
pyOfficeEditor, as committed at 098e441, with its imports pointed here. It
reads a formula as Excel does, into a tree, and calculates it: 493 of
Excel's functions, its precedence, coercion and implicit intersection, and
arithmetic as the x87 does it, held to 10,958 formulas Excel calculated.

- :mod:`.lexer` breaks formula text into tokens, references whole
- :mod:`.nodes` is the tree
- :mod:`.parser` builds the tree with Excel's precedence, which is not the
  usual one: ``-2^2`` is 4 and ``2^3^2`` is 64
- :mod:`.evaluator` works a tree out against a workbook, which it reads
  through the :class:`~.evaluator.Book` protocol
- :mod:`.registry` and :mod:`.functions` are the functions

Every module keeps its name from pyOfficeEditor, so a later copy is a diff
of the same files: scripts/copy_formula_engine.py shows the differences
and makes the copy. The copy carries changes of pyOpenVBA's own, each for
one of its live measurements, that a later copy has to keep; they are
listed in docs/formula_engine.md. pyOfficeEditor's own workbook driver, ``engine.py``,
is not copied; :mod:`pyopenvba.apps.excel._engine_book` gives the engine
a workbook of the model to read. The rest of pyOfficeEditor the engine
used came along as :mod:`.collate` and :mod:`.reference`; :mod:`.cells`
and :mod:`.host` are pyOpenVBA's own, where the engine meets this library.
"""

from __future__ import annotations

from pyopenvba.formula._calc.parser import FormulaSyntaxError, parse

__all__ = ["FormulaSyntaxError", "parse"]
