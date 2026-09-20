"""A VBA interpreter in pure Python.

The language on its own, with no host: enough to run a module, keep its
state, and report what happened.  A host application plugs in through
:class:`~pyopenvba.interpreter._runtime.HostBridge`; see
:mod:`pyopenvba.apps.excel` for a real one.

    >>> from pyopenvba.interpreter import Interpreter
    >>> vba = Interpreter()
    >>> _ = vba.add_module('''
    ... Function Doubled(ByVal n As Long) As Long
    ...     Doubled = n * 2
    ... End Function
    ... ''', name="Module1")
    >>> int(vba.run("Doubled", [21]))
    42

Three errors, kept apart on purpose:
:class:`~pyopenvba.exceptions.VBACompileError` for VBA that does not
compile, :class:`~pyopenvba.exceptions.VBARuntimeError` for an error
VBA itself would raise while running, and
:class:`~pyopenvba.exceptions.VBAUnsupportedError` for VBA that is
real and that pyOpenVBA does not implement.
"""

from pyopenvba.exceptions import (
    VBACompileError,
    VBAError,
    VBARuntimeError,
    VBAUnsupportedError,
)
from pyopenvba.interpreter._lex import tokenize
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, ref_setter, setter
from pyopenvba.interpreter._parse import parse_expression, parse_module
from pyopenvba.interpreter._runtime import Collection, HostBridge, Interpreter, ModuleRuntime
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBADecimal,
    VBAErrorValue,
    VBAInt,
    VBASingle,
    to_bool,
    to_date,
    to_number,
    to_text,
    type_name,
)

__all__ = [
    "EMPTY",
    "MISSING",
    "NOTHING",
    "NULL",
    "Collection",
    "HostBridge",
    "Interpreter",
    "ModuleRuntime",
    "VBAArray",
    "VBACollection",
    "VBACompileError",
    "VBACurrency",
    "VBADate",
    "VBADecimal",
    "VBAError",
    "VBAErrorValue",
    "VBAInt",
    "VBAObject",
    "VBARuntimeError",
    "VBASingle",
    "VBAUnsupportedError",
    "member",
    "method",
    "parse_expression",
    "parse_module",
    "ref_setter",
    "setter",
    "to_bool",
    "to_date",
    "to_number",
    "to_text",
    "tokenize",
    "type_name",
]
