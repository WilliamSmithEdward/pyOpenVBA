"""Formula2's @: where a formula written through Range.Formula cuts cells to one value.

A formula Range.Formula writes is worked out as Excel worked formulas out
before dynamic arrays: where one value is wanted, a range is cut to the
formula's own row or column and an array to its first item. Formula2
reads such a formula back with an @ at each place that happens, which is
how the same formula is written for dynamic arrays. Measured in live
Excel (scripts/measure_implicit_intersection.py,
tests/fixtures/implicit_intersection.json): each of Excel's functions
written with each argument a cell, a range, and an operation on a range,
with its optional and repeating arguments too, and 150 formulas besides.

What can give several values: a range of more than one cell as written,
A1:A1 among them, a whole column or row, a range or an intersection the
operators make, a defined name for such cells, an array or a formula, an
array constant, a call of one of the functions that give arrays -- FILTER,
SEQUENCE, TRANSPOSE, INDIRECT, IFS and the rest -- and of a function
Excel has not got or a name defines, and some calls by what they are
given: INDEX without a row and column that are numbers other than 0,
OFFSET of a range or with a height or width other than 1, XLOOKUP
returning more than one cell, IF, CHOOSE and IFERROR choosing among such,
LET ending in one, ROW, COLUMN, ISFORMULA and FORMULATEXT of a range.

Where one value is wanted, such a thing takes the @: the whole formula,
an operand of an operator, and an argument a function takes as one value.
A function takes each argument in one of three ways, measured function by
function: as one value (ABS, a lookup's value); as a range, which takes
cells whole but cuts an operation on them (SUM, COUNT, MATCH's range):
SUM(@A1:A3*2); or whole, cutting nothing inside (SUMPRODUCT, MMULT and
the functions that give arrays). IF, CHOOSE, IFERROR and IFNA hand a
choice on untouched, and so do LET for what it binds and returns and a
LAMBDA for its result, a name they bind standing for what it is given.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Final

from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._calc.nodes import function_key
from pyopenvba.formula._calc.registry import FUNCTIONS
from pyopenvba.formula._deep import deep
from pyopenvba.formula._parse import (ArrayLiteral, Binary, Call, FormulaError, Invoke, Literal, NameNode, Node,
                                      Reference, Structured, Unary, parse, parse_placed, tokenize)

#: How a function takes each argument, by position, where it is not as one value (V) throughout: R as a range, A
#: whole. Past the last position given the last ones repeat, as many as the function's repeating arguments.
_ROLES: Final[dict[str, str]] = {
    "ACCRINT": "RRRRRRRR", "ACCRINTM": "RRRRR", "AGGREGATE": "VVARR", "AMORDEGRC": "RRRRRRR",
    "AMORLINC": "RRRRRRR", "AND": "RRR", "AREAS": "R", "ARRAYTOTEXT": "AV", "AVEDEV": "RRR", "AVERAGE": "RRR",
    "AVERAGEA": "RRR", "AVERAGEIF": "RVR", "AVERAGEIFS": "RRVRVRV", "BESSELI": "RR", "BESSELJ": "RR",
    "BESSELK": "RR", "BESSELY": "RR", "BIN2DEC": "R", "BIN2HEX": "RR", "BIN2OCT": "RR", "BYCOL": "AV",
    "BYROW": "AV", "CELL": "VR", "CHISQ.TEST": "AA", "CHITEST": "AA", "CHOOSE": "VRRR", "CHOOSECOLS": "AVVV",
    "CHOOSEROWS": "AVVV", "COLUMN": "R", "COLUMNS": "A", "COMPLEX": "RRR", "CONCAT": "RRR", "CONVERT": "RRR",
    "CORREL": "AA", "COUNT": "RRR", "COUNTA": "RRR", "COUNTBLANK": "R", "COUNTIF": "RV", "COUNTIFS": "RVRVRV",
    "COUPDAYBS": "RRRR", "COUPDAYS": "RRRR", "COUPDAYSNC": "RRRR", "COUPNCD": "RRRR", "COUPNUM": "RRRR",
    "COUPPCD": "RRRR", "COVAR": "AA", "COVARIANCE.P": "AA", "COVARIANCE.S": "AA", "CUBEMEMBER": "VR",
    "CUBESET": "VR", "CUMIPMT": "RRRRRR", "CUMPRINC": "RRRRRR", "DAVERAGE": "RRR", "DCOUNT": "RRR",
    "DCOUNTA": "RRR", "DEC2BIN": "RR", "DEC2HEX": "RR", "DEC2OCT": "RR", "DELTA": "RR", "DEVSQ": "RRR",
    "DGET": "RRR", "DISC": "RRRRR", "DMAX": "RRR", "DMIN": "RRR", "DOLLARDE": "RR", "DOLLARFR": "RR",
    "DPRODUCT": "RRR", "DROP": "AVV", "DSTDEV": "RRR", "DSTDEVP": "RRR", "DSUM": "RRR", "DURATION": "RRRRRR",
    "DVAR": "RRR", "DVARP": "RRR", "EDATE": "RR", "EFFECT": "RR", "EOMONTH": "RR", "ERF": "RR", "ERF.PRECISE": "R",
    "ERFC": "R", "ERFC.PRECISE": "R", "EXPAND": "AVVV", "F.TEST": "AA", "FACTDOUBLE": "R", "FILTER": "AAV",
    "FORECAST": "VAA", "FORECAST.ETS": "VAA", "FORECAST.ETS.CONFINT": "VAA", "FORECAST.ETS.SEASONALITY": "AA",
    "FORECAST.ETS.STAT": "AAV", "FORECAST.LINEAR": "VAA", "FORMULATEXT": "R", "FREQUENCY": "AA", "FTEST": "AA",
    "FVSCHEDULE": "RR", "GCD": "RRR", "GEOMEAN": "RRR", "GESTEP": "RR", "GETPIVOTDATA": "RR", "GROUPBY": "AAA",
    "GROWTH": "AAA", "HARMEAN": "RRR", "HEX2BIN": "RR", "HEX2DEC": "R", "HEX2OCT": "RR", "HLOOKUP": "VRVV",
    "HSTACK": "AAA", "IF": "VRR", "IFERROR": "VR", "IFNA": "VR", "IFS": "VRVRVR", "IMABS": "R", "IMAGINARY": "R",
    "IMARGUMENT": "R", "IMCONJUGATE": "R", "IMCOS": "R", "IMCOSH": "R", "IMCOT": "R", "IMCSC": "R", "IMCSCH": "R",
    "IMDIV": "RR", "IMEXP": "R", "IMLN": "R", "IMLOG10": "R", "IMLOG2": "R", "IMPOWER": "RR", "IMPRODUCT": "RRR",
    "IMREAL": "R", "IMSEC": "R", "IMSECH": "R", "IMSIN": "R", "IMSINH": "R", "IMSQRT": "R", "IMSUB": "RR",
    "IMSUM": "RRR", "IMTAN": "R", "INDEX": "AVV", "INTERCEPT": "AA", "INTRATE": "RRRRR", "IRR": "AV", "ISEVEN": "R",
    "ISFORMULA": "R", "ISODD": "R", "ISOMITTED": "R", "ISREF": "R", "KURT": "RRR", "LAMBDA": "RR", "LARGE": "RV",
    "LCM": "RRR", "LET": "RR", "LINEST": "AAR", "LOGEST": "AAR", "LOOKUP": "VAA", "MAKEARRAY": "VVR", "MAP": "ARR",
    "MATCH": "VRR", "MAX": "RRR", "MAXA": "RRR", "MAXIFS": "RRVRVRV", "MDETERM": "A", "MDURATION": "RRRRRR",
    "MEDIAN": "RRR", "MIN": "RRR", "MINA": "RRR", "MINIFS": "RRVRVRV", "MINVERSE": "A", "MIRR": "AVV", "MMULT": "AA",
    "MODE": "AAA", "MODE.MULT": "AAA", "MODE.SNGL": "AAA", "MROUND": "RR", "MULTINOMIAL": "RRR", "N": "R",
    "NETWORKDAYS": "RRR", "NETWORKDAYS.INTL": "RRVR", "NOMINAL": "RR", "NPV": "VRRR", "OCT2BIN": "RR",
    "OCT2DEC": "R", "OCT2HEX": "RR", "OFFSET": "RVVVV", "OR": "RRR", "PEARSON": "AA", "PERCENTILE": "RV",
    "PERCENTILE.EXC": "RV", "PERCENTILE.INC": "RV", "PERCENTOF": "RR", "PERCENTRANK": "RVV",
    "PERCENTRANK.EXC": "RVV", "PERCENTRANK.INC": "RVV", "PHONETIC": "R", "PIVOTBY": "AAAA", "PRICE": "RRRRRRR",
    "PRICEDISC": "RRRRR", "PRICEMAT": "RRRRRR", "PROB": "AAVV", "PRODUCT": "RRR", "QUARTILE": "RV",
    "QUARTILE.EXC": "RV", "QUARTILE.INC": "RV", "QUOTIENT": "RR", "RANDBETWEEN": "RR", "RANK": "VRV",
    "RANK.AVG": "VRV", "RANK.EQ": "VRV", "RECEIVED": "RRRRR", "REDUCE": "RAV", "ROW": "R", "ROWS": "A", "RSQ": "AA",
    "SCAN": "RAV", "SERIESSUM": "RRRR", "SHEET": "R", "SHEETS": "A", "SINGLE": "R", "SKEW": "RRR", "SKEW.P": "RRR",
    "SLOPE": "AA", "SMALL": "RV", "SORT": "AAA", "SORTBY": "AA", "SQRTPI": "R", "STDEV": "RRR", "STDEV.P": "RRR",
    "STDEV.S": "RRR", "STDEVA": "RRR", "STDEVP": "RRR", "STDEVPA": "RRR", "STEYX": "AA", "SUBTOTAL": "VRRR",
    "SUM": "RRR", "SUMIF": "RVR", "SUMIFS": "RRVRVRV", "SUMPRODUCT": "AAA", "SUMSQ": "RRR", "SUMX2MY2": "AA",
    "SUMX2PY2": "AA", "SUMXMY2": "AA", "SWITCH": "VVRVRVR", "T": "R", "T.TEST": "AAVV", "TAKE": "AVV",
    "TBILLEQ": "RRR", "TBILLPRICE": "RRR", "TBILLYIELD": "RRR", "TEXTJOIN": "RVRRR", "TOCOL": "AVV", "TOROW": "AVV",
    "TREND": "AAA", "TRIMMEAN": "RV", "TRIMRANGE": "AVV", "TTEST": "AAVV", "UNIQUE": "AVV", "VAR": "RRR",
    "VAR.P": "RRR", "VAR.S": "RRR", "VARA": "RRR", "VARP": "RRR", "VARPA": "RRR", "VLOOKUP": "VRVV",
    "VSTACK": "AAA", "WEEKNUM": "RR", "WORKDAY": "RRR", "WORKDAY.INTL": "RRVR", "WRAPCOLS": "AVV",
    "WRAPROWS": "AVV", "XIRR": "RRR", "XLOOKUP": "VAARV", "XMATCH": "VAVV", "XNPV": "RRR", "XOR": "RRR",
    "YEARFRAC": "RRR", "YIELD": "RRRRRRR", "YIELDDISC": "RRRRR", "YIELDMAT": "RRRRRR", "Z.TEST": "RVV",
    "ZTEST": "RVV",
}
#: The functions whose call can always give several values.
_ARRAYS: Final = frozenset("""
BYCOL BYROW CELL CHOOSECOLS CHOOSEROWS DROP EUROCONVERT EXPAND FILTER FILTERXML FREQUENCY GROUPBY GROWTH HSTACK IFS
INDIRECT LAMBDA LINEST LOGEST MAKEARRAY MAP MINVERSE MMULT MODE.MULT MUNIT PIVOTBY RANDARRAY REDUCE REGEXEXTRACT
SCAN SEQUENCE SINGLE SORT SORTBY STOCKHISTORY SWITCH TAKE TEXTSPLIT TOCOL TOROW TRANSPOSE TREND TRIMRANGE UNIQUE
VSTACK WRAPCOLS WRAPROWS
""".split())
#: The functions that hand some of their arguments on untouched, and which: CHOOSE's from the second on.
_CHOICES: Final[dict[str, Callable[[int], bool]]] = {
    "IF": lambda index: index in (1, 2), "IFERROR": lambda index: index == 1, "IFNA": lambda index: index == 1,
    "CHOOSE": lambda index: index > 0, "XLOOKUP": lambda index: index == 3,
}

#: Where a thing stands: one value wanted, a range taken, taken whole, or handed on untouched.
VALUE, RANGE, WHOLE, HANDED = "V", "R", "A", "H"


class UnreadFormula2Error(FormulaError):
    """A formula whose Formula2 the model does not work out."""


@dataclass
class _Walk:
    starts: dict[int, int]
    #: A defined name's formula, or None for a name the workbook does not define.
    named: Callable[[str], str | None]
    marks: set[int] = field(default_factory=lambda: set())
    #: What each name a LET or a LAMBDA binds stands for: whether it can give several values.
    bound: list[dict[str, bool]] = field(default_factory=lambda: [])

    def visit(self, node: Node | None, where: str) -> bool:
        """Mark where ``node`` takes an @ standing ``where``, and say whether it can give several values then."""
        many = self._many(node, where)
        if many and where == VALUE and self._takes(node):
            self.marks.add(self.starts[id(node)])
            return False
        return many

    def _takes(self, node: Node | None) -> bool:
        """Whether a thing that can give several values takes an @ where one is wanted: all but a union."""
        return not (isinstance(node, Binary) and node.op == ",")

    def _many(self, node: Node | None, where: str) -> bool:
        if node is None or isinstance(node, Literal):
            return False
        if isinstance(node, Reference):
            return ":" in node.text
        if isinstance(node, Structured):
            raise UnreadFormula2Error("Formula2 of a formula with a structured reference is not implemented")
        if isinstance(node, ArrayLiteral):
            for row in node.rows:
                for item in row:
                    self.visit(item, WHOLE)
            return True
        if isinstance(node, NameNode):
            return self._name(node)
        if isinstance(node, Unary):
            if node.op == "@":
                self.visit(node.operand, WHOLE)
                return False
            return self.visit(node.operand, WHOLE if where == WHOLE else VALUE) and where == WHOLE
        if isinstance(node, Binary):
            if node.op in (":", " ", ","):
                self.visit(node.left, WHOLE)
                self.visit(node.right, WHOLE)
                return True
            inside = WHOLE if where == WHOLE else VALUE
            left = self.visit(node.left, inside)
            right = self.visit(node.right, inside)
            return where == WHOLE and (left or right)
        if isinstance(node, Invoke):
            return self._invoke(node)
        if isinstance(node, Call):
            return self._call(node, where)
        return False

    def _name(self, node: NameNode) -> bool:
        key = node.name.upper()
        if not node.sheet:
            for scope in reversed(self.bound):
                if key in scope:
                    return scope[key]
        formula = self.named(node.name if not node.sheet else f"{node.sheet}!{node.name}")
        if formula is None:
            return False
        try:
            tree = parse(formula)
        except FormulaError:
            return True
        if isinstance(tree, Literal):
            return False
        if isinstance(tree, Reference):
            return ":" in tree.text
        return not (isinstance(tree, Call) and function_key(tree.name) == "LAMBDA")

    def _invoke(self, node: Invoke) -> bool:
        target = node.target
        if isinstance(target, Call) and function_key(target.name) == "LAMBDA":
            self._lambda(target)
        else:
            self.visit(target, WHOLE)
        for argument in node.args:
            # Handed on to the parameters as it is; only an operation on cells is cut: (@A1:A2*1).
            self.visit(argument, HANDED)
        return True

    def _lambda(self, node: Call) -> None:
        """A LAMBDA's parameters stand for what it is given, cells or an array: its body cuts them."""
        *parameters, body = node.args or [None]
        names = [one.name for one in parameters if isinstance(one, NameNode)]
        # A parameter a call may leave out is written in brackets, [y].
        names += [one.first for one in parameters if isinstance(one, Structured) and one.first is not None]
        self.bound.append({name.upper(): True for name in names})
        try:
            self.visit(body, HANDED)
        finally:
            self.bound.pop()

    def _let(self, node: Call) -> bool:
        scope: dict[str, bool] = {}
        self.bound.append(scope)
        try:
            *pairs, result = node.args or [None]
            for index in range(0, len(pairs) - 1, 2):
                name = pairs[index]
                many = self.visit(pairs[index + 1], HANDED)
                if isinstance(name, NameNode):
                    scope[name.name.upper()] = many
            return self.visit(result, HANDED)
        finally:
            self.bound.pop()

    def _call(self, node: Call, where: str) -> bool:
        key = function_key(node.name)
        if key == "LET":
            return self._let(node)
        if key == "LAMBDA":
            # A LAMBDA given to MAP and the like is one function, not several values.
            self._lambda(node)
            return False
        if not is_excel_function(key) or self._defined(node.name):
            # A function Excel has not got, or one a name defines: what it gives is not known.
            for argument in node.args:
                self.visit(argument, WHOLE)
            return True
        chosen = _CHOICES.get(key)
        entry = FUNCTIONS.get(key)
        many: list[bool] = []
        handed = False
        for index, argument in enumerate(node.args):
            if chosen is not None and chosen(index):
                handed = self.visit(argument, HANDED) or handed
                many.append(False)
                continue
            role = {"V": VALUE, "R": RANGE, "A": WHOLE}[_role(key, index)]
            if where == WHOLE and role == VALUE and not _as_cell(entry, index):
                # Worked out whole, a function runs over an array given where it takes one value, as ABS does
                # inside SUMPRODUCT; only the places a legacy formula cuts even there take an @.
                role = WHOLE
            many.append(self.visit(argument, role))
        return handed or key in _ARRAYS or self._many_by_arguments(key, node.args, many)

    def _defined(self, name: str) -> bool:
        return self.named(name) is not None and not is_excel_function(function_key(name))

    def _many_by_arguments(self, key: str, args: list[Node], many: list[bool]) -> bool:
        """Whether a function that gives several values only when given them does so here; ``many`` says which
        arguments can give several."""
        if key == "INDEX":
            if not args or not isinstance(args[0], (Reference, NameNode)):
                return True
            return any(not _whole_number(argument) for argument in args[1:3])
        if key == "OFFSET":
            return bool(many and many[0]) or any(not _is_one(argument) for argument in args[3:5])
        if key in ("ROW", "COLUMN", "ISFORMULA", "FORMULATEXT"):
            return bool(many and many[0])
        if key == "XLOOKUP":
            return _xlookup_many(args)
        return False


def _role(key: str, index: int) -> str:
    roles = _ROLES.get(key)
    if not roles:
        return "V"
    if index < len(roles):
        return roles[index]
    entry = FUNCTIONS.get(key)
    repeat = max(1, entry.repeat if entry is not None else 1)
    tail = roles[-repeat:]
    return tail[(index - len(roles)) % len(tail)]


def _as_cell(entry: object, index: int) -> bool:
    """Whether a legacy formula cuts an argument to one value even inside one worked out whole: the engine's
    legacy_first and legacy_cell places, and the functions that work their arguments out as a cell does."""
    from pyopenvba.formula._calc.registry import Function

    if not isinstance(entry, Function):
        return False
    return entry.legacy_as_cell or index in entry.legacy_first or index in entry.legacy_cell


def _whole_number(node: Node | None) -> bool:
    """A row or column INDEX takes as one cell: a number other than 0 written out."""
    return isinstance(node, Literal) and isinstance(node.value, float) and node.value != 0.0


def _is_one(node: Node | None) -> bool:
    return isinstance(node, Literal) and node.value == 1.0


def _shape(node: Node | None) -> tuple[int, int] | None:
    """The rows and columns of a range written as one, or None."""
    from pyopenvba._a1 import parse_area

    if not isinstance(node, Reference):
        return None
    try:
        area = parse_area(node.text, sheet="")
    except ValueError:
        return None
    return area.rows, area.columns


def _xlookup_many(args: list[Node]) -> bool:
    """XLOOKUP gives several values when what it returns is not a range, is wider than one cell across the way it
    looks, or when what it gives when nothing is found can be several."""
    if len(args) < 3:
        return False
    back = _shape(args[2])
    if back is None:
        return True
    looked = _shape(args[1])
    across = back[1] if looked is None or looked[1] == 1 else back[0]
    return across > 1


def formula2(formula: str, named: Callable[[str], str | None], *, whole: bool = False) -> str:
    """``formula``, as Range.Formula spells it, as Formula2 reads it back: with an @ where it cuts cells to one
    value. ``named`` gives a defined name's formula, or None; ``whole`` for an array formula, worked out whole,
    which cuts nothing at its top."""
    return deep(lambda: _formula2(formula, named, whole=whole))


def _formula2(formula: str, named: Callable[[str], str | None], *, whole: bool) -> str:
    tree, starts = parse_placed(formula)
    walk = _Walk(starts, named)
    walk.visit(tree, WHOLE if whole else VALUE)
    body = formula[1:] if formula.startswith("=") else formula
    for start in sorted(walk.marks, reverse=True):
        body = body[:start] + "@" + body[start:]
    return "=" + body


def legacy(formula: str) -> str:
    """A formula as Formula2 spells it, its @ taken out: what Range.Formula would write for it."""
    body = formula[1:] if formula.startswith("=") else formula
    tokens = tokenize(body, spaces=True)
    return "=" + "".join(token.text for token in tokens if not (token.kind == "op" and token.text == "@"))


__all__ = ["UnreadFormula2Error", "formula2", "legacy"]
