"""How a file spells what Excel 2007 could not read: newer functions, and the names a LET or a LAMBDA binds.

Measured in live Excel (scripts/measure_formula_prefixes.py,
tests/fixtures/formula_prefixes/): a file writes a function Excel 2007
did not have as ``_xlfn.NAME``, FILTER and SORT as ``_xlfn._xlws.NAME``,
and each name a LET or a LAMBDA binds as ``_xlpm.name``, where it is
bound and wherever it stands for what it is bound to, called or not, in
an inner LET too. Range.Formula reads each without. Which functions take
``_xlfn.`` is Excel's own list, not the year they came:
NETWORKDAYS.INTL is newer than 2007 and takes none.

An @, which cuts what it holds to one value, is a call of SINGLE in a
file, =@A1 being _xlfn.SINGLE(A1); and Excel writes a formula's call of
SINGLE as an @, its argument as it stands (tests/fixtures/at_sign.json).

A file also marks a formula ``ca="1"``, worked out whenever anything
changes, when it calls one of a few functions, NOW, OFFSET, INDIRECT and
the like, or a function neither Excel nor the workbook's macros have. An
array formula like that is ``aca="1"`` too, unless it covers several
cells and calls RAND.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from typing import Final

from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._calc.nodes import function_key
from pyopenvba.formula._parse import FormulaError, Token, tokenize
from pyopenvba.formula._r1c1 import is_cell

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
#: A LAMBDA's parameter a call may leave out, [y], which a file writes _xlop.y; and a name that looks like a cell,
#: 'A1', which a file writes _xlnm.A1 (tests/fixtures/bound_names.json).
_OPTIONAL = "_xlop."
_NAMED = "_xlnm."
#: What spilled from a cell, A1#, as a file writes it.
_ANCHORED = "_xlfn.ANCHORARRAY"
_OPTIONAL_NAME = re.compile(r"\[([^\[\]\s]+)\]")


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
    where it stands for that, "optional" for a LAMBDA's parameter in brackets, [y], "function" for a function
    called, "name" for anything else."""
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
        elif token.kind == "structured":
            inner = open_[-1] if open_ else None
            optional = _OPTIONAL_NAME.fullmatch(token.text)
            if optional is not None and inner is not None and inner.binder == "LAMBDA" and inner.starting \
                    and following.kind == "comma":
                inner.bound.add(function_key(optional.group(1)))
                yield token, "optional"
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


def _closing(tokens: list[Token], index: int) -> int:
    """The index of the bracket or brace closing the one at ``index``."""
    depth = 0
    for place in range(index, len(tokens)):
        if tokens[place].kind in ("open", "lbrace"):
            depth += 1
        elif tokens[place].kind in ("close", "rbrace"):
            depth -= 1
            if depth == 0:
                return place
    raise FormulaError("a bracket is not closed")


def _called(tokens: list[Token], index: int) -> bool:
    """Whether the token at ``index`` is a function's name, its bracket straight after it."""
    if index + 1 >= len(tokens):
        return False
    token, following = tokens[index], tokens[index + 1]
    return token.kind == "name" and following.kind == "open" and following.at == token.at + len(token.text)


def _primary_end(tokens: list[Token], index: int) -> int:
    """Just past the primary at ``index``: a call with its brackets, something in brackets or braces, or one token."""
    if _called(tokens, index):
        return _closing(tokens, index + 1) + 1
    if tokens[index].kind in ("open", "lbrace"):
        return _closing(tokens, index) + 1
    return index + 1


def _joined(tokens: list[Token], index: int) -> bool:
    """Whether a reference operator joins more to what ends just before ``index``: a range's colon, a spill's #, or
    a space between two operands, which intersects them."""
    token, before = tokens[index], tokens[index - 1]
    if token.kind == "op":
        return token.text in (":", "#")
    return token.kind in ("ref", "structured", "name", "open") and token.at > before.at + len(before.text)


def _operand_end(tokens: list[Token], index: int) -> int:
    """Just past what an @ at ``index - 1`` holds, as Excel reads it: the signs before an operand, the operand, and
    what a range, an intersection or a spill joins to it; a percent sign after it is outside, @A1:A3% being
    (@A1:A3)% (tests/fixtures/at_sign.json)."""
    while tokens[index].kind == "op" and tokens[index].text in ("+", "-", "@"):
        index += 1
    index = _primary_end(tokens, index)
    while _joined(tokens, index):
        token = tokens[index]
        if token.kind != "op":
            # A space, and what it intersects with.
            index = _primary_end(tokens, index)
        elif token.text == "#":
            index += 1
        else:
            # A colon, and the range's other end.
            index = _primary_end(tokens, index + 1)
    return index


def _as_single(formula: str) -> str:
    """``formula`` with each @ written as the SINGLE a file keeps it as, =@A1 as =SINGLE(A1)."""
    tokens = _tokens(formula)
    if tokens is None or not any(token.kind == "op" and token.text == "@" for token in tokens):
        return formula
    offset = 1 if formula.startswith("=") else 0
    # Each @ becomes SINGLE and its bracket, and a bracket closes after what it holds: each change is a place, how
    # many characters it replaces and with what. The last is made first, so the places before stay where they are.
    changes: list[tuple[int, int, str]] = []
    for index, token in enumerate(tokens):
        if token.kind == "op" and token.text == "@":
            last = tokens[_operand_end(tokens, index + 1) - 1]
            changes += [(offset + token.at, 1, "SINGLE("), (offset + last.at + len(last.text), 0, ")")]
    out = formula
    for position, length, text in sorted(changes, reverse=True):
        out = out[:position] + text + out[position + length:]
    return out


def _singles(tokens: list[Token]) -> Iterator[tuple[int, int]]:
    """Where SINGLE is called with one argument: the index of its name and of the bracket closing its call."""
    for index in range(len(tokens) - 1):
        if _called(tokens, index) and function_key(tokens[index].text.removeprefix(_FUNCTION)) == "SINGLE":
            close = _closing(tokens, index + 1)
            if close > index + 2 and not any(token.kind == "comma" for token in _level(tokens, index + 1, close)):
                yield index, close


def _level(tokens: list[Token], opening: int, close: int) -> Iterator[Token]:
    """The tokens standing directly inside the brackets at ``opening`` and ``close``, not in brackets inside them."""
    depth = 0
    for token in tokens[opening + 1:close]:
        if token.kind in ("close", "rbrace"):
            depth -= 1
        if depth == 0:
            yield token
        if token.kind in ("open", "lbrace"):
            depth += 1


def _at_for_single(formula: str, *, kept: bool) -> str:
    """``formula`` with a call of SINGLE written as the @ Excel writes it as, its argument as it stands; ``kept`` for
    only the calls whose @ holds just what SINGLE does."""
    offset = 1 if formula.startswith("=") else 0
    while True:
        tokens = _tokens(formula)
        if tokens is None:
            return formula
        found = next(((name, close) for name, close in _singles(tokens)
                      if not kept or _holds_the_same(tokens, name, close)), None)
        if found is None:
            return formula
        # One at a time, as each changes where the rest stand.
        name, close = found
        start, argument, end = offset + tokens[name].at, offset + tokens[name + 1].at + 1, offset + tokens[close].at
        formula = formula[:start] + "@" + formula[argument:end] + formula[end + 1:]


def _holds_the_same(tokens: list[Token], name: int, close: int) -> bool:
    """Whether an @ in place of the SINGLE called at ``name`` would hold what the call does: all of its argument,
    and nothing a reference operator joins to it on either side."""
    inner = [*tokens[name + 2:close], Token("eof", "", tokens[close].at)]
    if _operand_end(inner, 0) != len(inner) - 1:
        return False
    if close + 1 < len(tokens) and _joined(tokens, close + 1):
        return False
    if not name:
        return True
    before = tokens[name - 1]
    if before.kind == "op":
        return before.text != ":"
    # A space after an operand intersects it with the call.
    return before.kind not in ("ref", "structured", "name", "close") or tokens[name].at == before.at + len(before.text)


def at_kept(formula: str) -> str:
    """``formula`` with each call of SINGLE written as an @, as Excel writes it, where the @ holds just what SINGLE
    does: =SINGLE(A1:A3) is =@A1:A3. SINGLE(A1:A3%) stays as it is, since @A1:A3% is (@A1:A3)%; :func:`at_shown`
    writes it as Excel shows it (tests/fixtures/at_sign.json)."""
    return _at_for_single(formula, kept=True) if "SINGLE" in formula.upper() else formula


def at_shown(formula: str) -> str:
    """``formula`` as Excel shows it: each call of SINGLE an @ and its argument as it stands, brackets added nowhere,
    =SINGLE(A1:A3%) shown =@A1:A3% (tests/fixtures/at_sign.json)."""
    return _at_for_single(formula, kept=False) if "SINGLE" in formula.upper() else formula


def intersected(formula: str) -> bool:
    """Whether ``formula`` cuts something to one value with an @ or a call of SINGLE."""
    if "@" not in formula and "SINGLE" not in formula.upper():
        return False
    tokens = _tokens(formula)
    return tokens is not None and (any(token.kind == "op" and token.text == "@" for token in tokens)
                                   or next(_singles(tokens), None) is not None)


def in_file(formula: str) -> str:
    """``formula`` as a file spells it: each @ as SINGLE, =@A1 as =_xlfn.SINGLE(A1) (tests/fixtures/at_sign.json),
    and each newer function and each bound name with its prefix."""
    formula = _as_single(formula)
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    changes: list[tuple[Token, str]] = []
    for token, role in _roles(tokens):
        head, bang, bare = token.text.rpartition("!")
        if role == "bound":
            changes.append((token, _BOUND + token.text))
        elif role == "optional":
            changes.append((token, _OPTIONAL + token.text[1:-1]))
        elif role == "name" and bare.startswith("'"):
            changes.append((token, head + bang + _NAMED + bare[1:-1].replace("''", "'")))
        elif role == "function" and "!" not in token.text and token.text.upper() == function_key(token.text):
            key = token.text.upper()
            if key in _SHEET_ONLY:
                changes.append((token, _SHEET + token.text))
            elif key in _NEWER:
                changes.append((token, _FUNCTION + token.text))
    for index, token in enumerate(tokens[:-1]):
        following = tokens[index + 1]
        if token.kind == "ref" and following.kind == "op" and following.text == "#":
            # What spilled from a cell, A1#, which a file writes _xlfn.ANCHORARRAY(A1) (tests/fixtures/dynamic_arrays.json).
            changes += [(token, f"{_ANCHORED}({token.text})"), (following, "")]
    return _replaced(formula, changes)


def from_file(formula: str) -> str:
    """``formula``, as a file spells it, as Range.Formula spells it: the prefixes of newer functions and of bound
    names dropped, and SINGLE the @ Excel reads it as (see :func:`at_kept`)."""
    return at_kept(_unprefixed(formula))


def _unprefixed(formula: str) -> str:
    if "_xl" not in formula.lower():
        return formula
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    changes: list[tuple[Token, str]] = []
    spills = [index for index in range(len(tokens) - 3) if tokens[index].kind == "name"
              and tokens[index].text.lower() == _ANCHORED.lower() and tokens[index + 1].kind == "open"
              and tokens[index + 2].kind == "ref" and tokens[index + 3].kind == "close"]
    for index in spills:
        # _xlfn.ANCHORARRAY(A1) is what spilled from A1, A1#.
        changes += [(tokens[index], ""), (tokens[index + 1], ""), (tokens[index + 3], "#")]
    for index, token in enumerate(tokens):
        if token.kind != "name" or index in spills:
            continue
        head, bang, text = token.text.rpartition("!")
        lower = text.lower()
        if lower.startswith(_OPTIONAL):
            text = f"[{text[len(_OPTIONAL):]}]"
        elif lower.startswith(_NAMED) and is_cell(text[len(_NAMED):]):
            text = "'" + text[len(_NAMED):] + "'"
        else:
            for prefix in (_SHEET, _FUNCTION, _BOUND):
                if lower.startswith(prefix.lower()):
                    text = text[len(prefix):]
                    break
        if head + bang + text != token.text:
            changes.append((token, head + bang + text))
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


def bound(formula: str) -> set[int]:
    """Where in ``formula``, without its =, each name a LET or a LAMBDA binds stands: where it is bound, a LAMBDA's
    [y] among them, and where it stands for that."""
    tokens = _tokens("=" + formula)
    return set() if tokens is None else {token.at for token, role in _roles(tokens) if role in ("bound", "optional")}


def newer(name: str) -> bool:
    """Whether a file writes a call of the function ``name``, in capitals, behind _xlfn.: whether it is newer than
    Excel 2007."""
    return name in _NEWER or name in _SHEET_ONLY


def calls(formula: str, name: str) -> bool:
    """Whether a formula calls the function ``name``."""
    tokens = _tokens(formula)
    return tokens is not None and any(role == "function" and function_key(token.text) == name
                                      for token, role in _roles(tokens))


__all__ = ["at_kept", "at_shown", "bound", "calculated_always", "calls", "from_file", "in_file", "intersected", "newer"]
