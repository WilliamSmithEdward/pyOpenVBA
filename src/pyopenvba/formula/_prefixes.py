"""How a file spells what Excel 2007 could not read: newer functions, and the names a LET or a LAMBDA binds.

Measured in live Excel (scripts/measure_formula_prefixes.py,
tests/fixtures/formula_prefixes/): a file writes a function Excel 2007
did not have as ``_xlfn.NAME``, FILTER and SORT as ``_xlfn._xlws.NAME``,
and each name a LET or a LAMBDA binds as ``_xlpm.name``, where it is
bound and wherever it stands for what it is bound to, called or not, in
an inner LET too. Range.Formula reads each without. Which functions take
``_xlfn.`` is Excel's own list, not the year they came:
NETWORKDAYS.INTL is newer than 2007 and takes none.

A file also marks a formula ``ca="1"``, worked out whenever anything
changes, when it calls one of a few functions, NOW, OFFSET, INDIRECT and
the like, or a function neither Excel nor the workbook's macros have. An
array formula like that is ``aca="1"`` too, unless it covers several
cells and calls RAND.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import Final

from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._calc.nodes import function_key
from pyopenvba.formula._parse import FormulaError, Token, tokenize

#: The functions a file writes as _xlfn.NAME: each Excel was measured writing so, of the 525 it has.
_NEWER: Final = frozenset("""
ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND
BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV
CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P
COVARIANCE.S CSC CSCH DAYS DECIMAL DETECTLANGUAGE DROP ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST
F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT
FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE
GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA
ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT
NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC
PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY
RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SINGLE SKEW.P
SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER
TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK
WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST
""".split())
#: The functions a file writes as _xlfn._xlws.NAME.
_SHEET_ONLY: Final = frozenset({"FILTER", "SORT"})
#: The functions whose formulas a file marks ca="1".
_ALWAYS: Final = frozenset({"CELL", "DETECTLANGUAGE", "EUROCONVERT", "FORMULATEXT", "INDIRECT", "INFO", "NOW",
                            "OFFSET", "RAND", "RANDARRAY", "RANDBETWEEN", "SHEET", "SHEETS", "TODAY", "TRANSLATE"})
#: The functions that bind names, and the prefixes a file gives what they bind and a function.
_BINDERS: Final = frozenset({"LET", "LAMBDA"})
_BOUND = "_xlpm."
_FUNCTION = "_xlfn."
_SHEET = "_xlfn._xlws."


@dataclass
class _Bracket:
    """A bracket open while a formula is read: a call's, LET's or LAMBDA's among them, or an array constant's."""

    binder: str = ""
    #: What a LET or a LAMBDA has bound so far, by the name in capitals.
    bound: set[str] = field(default_factory=lambda: set())
    #: Which argument is being read, and whether none of it has been yet.
    argument: int = 0
    starting: bool = True


def _roles(tokens: list[Token]) -> Iterator[tuple[Token, str]]:
    """Each name of a formula with what it is: "bound" for a name a LET or a LAMBDA binds, where it is bound and
    where it stands for that, "function" for a function called, "name" for anything else."""
    open_: list[_Bracket] = []
    for index, token in enumerate(tokens):
        following = tokens[index + 1] if index + 1 < len(tokens) else token
        # A call's bracket follows its name straight away; after a space it opens an intersection.
        calling = following.kind == "open" and following.at == token.at + len(token.text)
        if token.kind == "name":
            key = function_key(token.text.removeprefix(_BOUND))
            inner = open_[-1] if open_ else None
            declaring = (inner is not None and inner.binder and inner.starting and following.kind == "comma"
                         and "!" not in token.text and (inner.binder == "LAMBDA" or inner.argument % 2 == 0))
            if declaring:
                assert inner is not None
                inner.bound.add(key)
                yield token, "bound"
            elif "!" not in token.text and _is_bound(open_, key):
                yield token, "bound"
            elif calling:
                yield token, "function"
            else:
                yield token, "name"
        if open_ and token.kind != "eof":
            open_[-1].starting = False
        if token.kind in ("open", "lbrace"):
            previous = tokens[index - 1] if index else None
            called = previous is not None and previous.kind == "name" and token.kind == "open" \
                and token.at == previous.at + len(previous.text)
            key = function_key(previous.text) if called and previous is not None else ""
            open_.append(_Bracket(binder=key if key in _BINDERS and not _is_bound(open_, key) else ""))
        elif token.kind in ("close", "rbrace") and open_:
            open_.pop()
        elif token.kind == "comma" and open_:
            open_[-1].argument += 1
            open_[-1].starting = True


def _is_bound(open_: list[_Bracket], key: str) -> bool:
    return any(key in bracket.bound for bracket in open_)


def _tokens(formula: str) -> list[Token] | None:
    try:
        return tokenize(formula[1:] if formula.startswith("=") else formula)
    except FormulaError:
        return None


def _replaced(formula: str, changes: list[tuple[Token, str]]) -> str:
    """``formula`` with each token given replaced by the text beside it."""
    offset = 1 if formula.startswith("=") else 0
    out = formula
    for token, text in sorted(changes, key=lambda change: change[0].at, reverse=True):
        start = offset + token.at
        out = out[:start] + text + out[start + len(token.text):]
    return out


def in_file(formula: str) -> str:
    """``formula`` as a file spells it: each newer function and each bound name with its prefix."""
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    changes: list[tuple[Token, str]] = []
    for token, role in _roles(tokens):
        if role == "bound":
            changes.append((token, _BOUND + token.text))
        elif role == "function" and "!" not in token.text and token.text.upper() == function_key(token.text):
            key = token.text.upper()
            if key in _SHEET_ONLY:
                changes.append((token, _SHEET + token.text))
            elif key in _NEWER:
                changes.append((token, _FUNCTION + token.text))
    return _replaced(formula, changes)


def from_file(formula: str) -> str:
    """``formula``, as a file spells it, as Range.Formula spells it: the prefixes of newer functions and of bound
    names dropped."""
    if "_xl" not in formula.lower():
        return formula
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    changes: list[tuple[Token, str]] = []
    for token in tokens:
        if token.kind != "name":
            continue
        text = token.text
        for prefix in (_SHEET, _FUNCTION, _BOUND):
            if text.lower().startswith(prefix.lower()):
                text = text[len(prefix):]
                break
        if text != token.text:
            changes.append((token, text))
    return _replaced(formula, changes)


def calculated_always(formula: str, known: Callable[[str], bool]) -> bool:
    """Whether a file marks a formula ca="1": it calls NOW, OFFSET or another of the functions a file marks so, or
    a function neither Excel nor the workbook has, ``known`` saying whether the workbook has a name: a macro's
    function, or a defined name."""
    tokens = _tokens(formula)
    if tokens is None:
        return False
    for token, role in _roles(tokens):
        if role != "function":
            continue
        key = function_key(token.text)
        if key in _ALWAYS or not (is_excel_function(key) or known(token.text)):
            return True
    return False


def calls(formula: str, name: str) -> bool:
    """Whether a formula calls the function ``name``."""
    tokens = _tokens(formula)
    return tokens is not None and any(role == "function" and function_key(token.text) == name
                                      for token, role in _roles(tokens))


__all__ = ["calculated_always", "calls", "from_file", "in_file"]
