"""pyOfficeEditor's formula corpus, worked out by the engine taken from it, over the model's workbook.

tests/fixtures/formula_corpus/inputs.xlsx is pyOfficeEditor's
formulas.xlsx, as its scripts/measure_formulas.py built it (commit
098e441): inputs written as exact doubles, and 10,958 formulas Excel
typed in and calculated, each on a sheet named for what it probes. The
copy keeps the name Excel calculated it under, inputs.xlsx, which CELL
shows, and the sheet About records the day it was built, the year a date
typed without one falls in.

Each formula is read from the model's workbook and worked out by
:mod:`pyopenvba.formula._calc` over the model's cells, through
:mod:`pyopenvba.apps.excel._engine_book`, and held to the value Excel
cached: to the bit, or within the units in the last place pyOfficeEditor
allows the functions its corpus narrows no further.
"""

from __future__ import annotations

import datetime as dt
import math
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._engine_book import EngineBook, model_value
from pyopenvba.apps.excel._model import Workbook
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.nodes import Call, walk
from pyopenvba.formula._values import ExcelError

FIXTURE = Path(__file__).parent / "fixtures" / "formula_corpus" / "inputs.xlsx"

#: Sheets that hold inputs, not formulas to check.
INPUTS = frozenset({"About", "Data", "Numbers", "Pairs", "FileLiterals", "Powers", "Rounding", "Divisions",
                    "DatePairs", "Reals", "Wholes", "Times", "Ties", "Crit", "Stats", "Db", "Samples"})

#: Formulas whose result the engine does not match, and why.
KNOWN: dict[str, str] = {
    # The rate is exactly 0; Excel's iteration stops at -1.96E-10.
    "Financial!R19C1": "RATE(10,-100,1000,0,0,0.1)",
}

#: Functions the corpus holds to within a number of units in the last place of Excel's result rather than to the
#: bit, as pyOfficeEditor holds them: the most any formula calling one is off by. Every other function is exact.
NEAR: dict[str, int] = {
    "SINH": 1, "TANH": 1, "ASIN": 1, "COTH": 2, "CSCH": 1,
    "PMT": 1, "IPMT": 1, "PPMT": 1, "CUMIPMT": 2, "CUMPRINC": 4,
    "NORM.DIST": 1, "NORM.S.DIST": 4, "NORM.S.INV": 3, "T.DIST": 6, "T.DIST.2T": 13, "T.DIST.RT": 13,
    "TDIST": 13, "T.INV": 1, "T.INV.2T": 35, "TINV": 35, "CHISQ.DIST": 52, "CHISQ.DIST.RT": 2, "CHIDIST": 2,
    "CHISQ.INV": 16, "F.DIST": 4, "F.INV": 1, "GAMMA": 15, "GAMMALN": 1, "GAMMA.DIST": 11, "GAMMADIST": 1,
    "GAMMA.INV": 1, "GAMMAINV": 1, "BETA.DIST": 2, "BETADIST": 2, "BETA.INV": 1, "BETAINV": 1,
    "LOGNORM.DIST": 3, "LOGNORMDIST": 1, "LOGNORM.INV": 2, "LOGINV": 2, "HYPGEOM.DIST": 3, "HYPGEOMDIST": 2,
    "NEGBINOM.DIST": 1, "NEGBINOMDIST": 1, "BINOM.DIST": 38, "POISSON.DIST": 9, "CONFIDENCE": 3,
    "CONFIDENCE.NORM": 3, "CONFIDENCE.T": 1, "T.TEST": 23, "TTEST": 20, "F.TEST": 34, "Z.TEST": 6,
    "LINEST": 4, "TREND": 1, "GEOMEAN": 1,
}  # fmt: skip

#: Functions Excel solves by iteration, stopping short of the root; the engine converges fully, and the two agree to
#: this fraction of the result.
ITERATIVE: dict[str, float] = {"RATE": 1e-12, "IRR": 1e-12, "XIRR": 1e-8, "YIELD": 1e-12}


@pytest.fixture(scope="module")
def book() -> Workbook:
    return ExcelApplication.open(FIXTURE, with_vba=False).workbook


def _sheets() -> list[str]:
    return [sheet.name for sheet in ExcelApplication.open(FIXTURE, with_vba=False).workbook.sheets_
            if sheet.name not in INPUTS]


def _same(got: object, want: object, units: int, relative: float) -> bool:
    if isinstance(want, float) and isinstance(got, float):
        difference = abs(got - want)
        return got == want or difference <= units * math.ulp(want) or difference <= relative * abs(want)
    if isinstance(want, ExcelError) and isinstance(got, ExcelError):
        return got.name == want.name
    return type(got) is type(want) and got == want


@pytest.mark.parametrize("name", _sheets())
def test_each_formula_comes_to_what_excel_cached(book: Workbook, name: str) -> None:
    serial = book.sheet_named("About").cells_[(2, 1)].value  # type: ignore[union-attr]
    assert isinstance(serial, float)
    built = dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))
    engine_book = EngineBook(book.calculator)
    sheet = book.sheet_named(name)
    assert sheet is not None
    wrong: list[str] = []
    for (row, column), cell in sorted(sheet.cells_.items()):
        if not cell.formula:
            continue
        address = f"{name}!R{row}C{column}"
        node = engine_book.read(cell.formula)
        called = {one.function for one in walk(node) if isinstance(one, Call)}
        array = (row, column) in sheet.array_formulas
        context = Context(engine_book, sheet.name, row, column, array=array, today=built,
                          now=dt.datetime.combine(built, dt.time()))
        value = context.formula(node)
        got: Any = model_value(context.first(context.array_of(value) if array else value))
        units = max((NEAR.get(one, 0) for one in called), default=0)
        relative = max((ITERATIVE.get(one, 0.0) for one in called), default=0.0)
        if not _same(got, cell.value, units, relative) and address not in KNOWN:
            wrong.append(f"{address} {cell.formula}: Excel {cell.value!r}, ours {got!r}")
        elif address in KNOWN and _same(got, cell.value, units, relative):
            wrong.append(f"{address} matches now; take it off KNOWN")
    assert not wrong, f"{len(wrong)} differ:\n" + "\n".join(wrong[:20])
