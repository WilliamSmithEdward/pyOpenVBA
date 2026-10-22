"""pyOfficeEditor's engine tests for what pyOpenVBA copied from its commits 81481d2 to 4723383.

pyOfficeEditor measured each of these against Excel (its
tests/test_excel_calc.py). Here they run through pyOpenVBA's copy of the
engine, over the model's workbook, so both projects' measurements hold at
once. ``_Book`` and ``_Sheet`` give the tests the calls they make on a
pyOfficeEditor workbook: cells set and read, formulas entered as Excel took
them before dynamic arrays, and ``evaluate``, which works a formula out in
a cell as pyOfficeEditor's engine does, with none of the legacy-formula
rules pyOpenVBA adds. The commits also taught pyOfficeEditor's own
workbook driver circular references, data tables, linked workbooks' cached
values and error checking; pyOpenVBA does not use that driver, so those
tests are not here.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pytest

from pyopenvba._a1 import column_letter
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Range
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.dates import parse_date_time
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.parser import parse
from pyopenvba.formula._calc.values import Empty, text_to_number

CellValue = str | float | bool | CellError | None
DATE_TEXTS = Path(__file__).parent / "fixtures" / "formula" / "date_texts.json"


class _Cell:
    def __init__(self, sheet: _Sheet, reference: str) -> None:
        self.sheet, self.reference = sheet, reference

    @property
    def value(self) -> object:
        """What the cell holds once worked out, as pyOfficeEditor reads a cell: a blank is 0."""
        range_ = self.sheet.range(self.reference)
        found = self.sheet.view.sheet.book.calculator.engine_book.cell(
            self.sheet.view.name, range_.first.top, range_.first.left)
        return 0.0 if isinstance(found, Empty) else found

    @value.setter
    def value(self, value: object) -> None:
        self.sheet.view.set_value(self.reference, value)

    @property
    def formula(self) -> str:
        return self.sheet.view.formula(self.reference).removeprefix("=")

    @formula.setter
    def formula(self, text: str) -> None:
        # Range.Formula enters it as Excel did before dynamic arrays.
        self.sheet.range(self.reference).vba_set("Formula", "=" + text)


class _Sheet:
    def __init__(self, app: ExcelApplication, name: str) -> None:
        self.view = app.sheet(name)

    def range(self, reference: str) -> Range:
        found = self.view.sheet.vba_get("Range", [reference])
        assert isinstance(found, Range)
        return found

    def __getitem__(self, reference: str) -> _Cell:
        return _Cell(self, reference)

    def __setitem__(self, reference: str, value: object) -> None:
        self.view.set_value(reference, value)
        assert self.view.value(reference) == value, reference

    def cell(self, row: int, column: int) -> _Cell:
        return _Cell(self, f"{column_letter(column)}{row}")

    def evaluate(self, formula: str, at: str = "A1", *, today: dt.date | None = None) -> object:
        """What ``formula`` gives in the cell ``at``, without putting it there, as pyOfficeEditor's evaluate has it."""
        book = self.view.sheet.book
        calculator = book.calculator
        now = calculator.now() if today is None else dt.datetime.combine(today, dt.time())
        place = self.range(at).first
        context = Context(calculator.engine_book, self.view.name, place.top, place.left, today=now.date(), now=now)
        context.legacy = False
        value = context.first(context.formula(parse(formula)))
        if isinstance(value, Empty):
            return 0
        if isinstance(value, float) and value.is_integer() and abs(value) < 2**53:
            return int(value)
        return value


class _Book:
    def __init__(self) -> None:
        self.app = ExcelApplication()
        self.app.add_workbook()
        self.app.workbook.sheets_[0].vba_set("Name", "Data")

    def __getitem__(self, name: str) -> _Sheet:
        return _Sheet(self.app, name)

    def add_sheet(self, name: str) -> _Sheet:
        self.app.workbook.add_sheet(name)
        return _Sheet(self.app, name)

    def calculate(self) -> None:
        """The model works a formula out when it is read, so there is nothing to start."""


@pytest.fixture
def book() -> _Book:
    return _Book()


def test_text_reads_as_a_date_or_a_time_as_excel_reads_it() -> None:
    # Strings measured in Excel by pyOfficeEditor's scripts/measure_date_texts.py, each with VALUE of it and
    # whether DATEVALUE and TIMEVALUE read it, which they do as VALUE's days and the rest.
    record = json.loads(DATE_TEXTS.read_text(encoding="utf-8"))
    measured = dt.date.fromisoformat(record["measured"])
    wrong: list[str] = []
    for key, epoch_1904 in (("cases", False), ("cases_1904", True)):
        for text, value, read in record[key]:
            found = parse_date_time(text, measured, epoch_1904=epoch_1904)
            as_measured = found == value if read else found is None
            if text_to_number(text, measured, epoch_1904=epoch_1904) != value or not as_measured:
                wrong.append(text)
    assert len(record["cases"]) > 18000
    assert wrong == []


@pytest.mark.parametrize(
    ("formula", "value"),
    [
        # Measured: a number that is no day of the month is a year.
        ('DATEVALUE("1/2020")', 43831.0),
        ('DATEVALUE("2/29")', 47150.0),
        ('DATEVALUE("Jan/99")', 36161.0),
        # A name, a day and a year take a comma.
        ('ISERROR(DATEVALUE("Jan 1 99"))', True),
        # One separator may trail a time, a comma with a space after it.
        ('TIMEVALUE("12:30, ")', 0.5208333333333334),
        ('ISERROR(TIMEVALUE("12:30,"))', True),
        # Minutes and seconds, and a field past its range.
        ('VALUE("93:22.15")', 0.06483969907407407),
        ('VALUE("25:00")', 1.0416666666666667),
        ('ISERROR(VALUE("24:60"))', True),
        # A month's name as a field is the negative of its number, unsigned.
        ('VALUE("Jan:5")', 178956970.6284722),
        # Numbers after a date and a time are dropped.
        ('VALUE("1/1/2020 12:00 4")', 43831.5),
        ('VALUE("12:00 PM1/1/2020")', 43831.5),
    ],
)
def test_a_formula_reads_a_date_or_a_time_from_text(book: _Book, formula: str, value: CellValue) -> None:
    assert book["Data"].evaluate(formula, today=dt.date(2026, 9, 26)) == value


def _bond(book: _Book, values: list[float]) -> str:
    """The arguments in row 1, where no formula literal cuts them to 15 digits."""
    sheet = book["Data"]
    for column, value in enumerate(values, start=1):
        sheet.cell(1, column).value = value
    return ",".join(f"{column_letter(column)}1" for column in range(1, len(values) + 1))


def test_yield_stops_up_to_1e_10_short_of_the_root(book: _Book) -> None:
    sheet = book["Data"]
    at_price = "YIELD(A1,B1,C1,PRICE(A1,B1,C1,D1,E1,F1,G1),E1,F1,G1)"
    # One Newton step from the start lands 1e-10 below the root, 0.15; the
    # next step would be under 1e-10, so that is the answer.
    _bond(book, [40000.0, 41533.0, 0.141, 0.15, 100.0, 2.0, 4.0])
    assert sheet.evaluate(at_price) == 0.14999999990009072
    # On a zero-coupon bond the start is one Halley step from 0, a cubic in
    # the root away from it, so near 0 it is the answer itself.
    _bond(book, [40000.0, 47300.0, 0.0, 0.00010000000000000026, 100.0, 2.0, 0.0])
    assert sheet.evaluate(at_price) == 9.999996673483998e-05
    # The start's years count 30/360 as YEARFRAC does: from 28 February,
    # the month's last day, to 31 May is 91 days, not 90.
    args = _bond(book, [44985.0, 47999.0, 0.045, 91.5, 100.0, 1.0, 0.0])
    assert sheet.evaluate(f"YIELD({args})") == 0.058221415821190996


def test_yield_with_one_coupon_left(book: _Book) -> None:
    sheet = book["Data"]
    # Actual/360: the period is the calendar's 182 days, not 360 / 2.
    args = _bond(book, [36505.0, 36596.0, 0.0, 91.36144230982995, 100.0, 2.0, 2.0])
    assert sheet.evaluate(f"YIELD({args})") == 0.3782145934550592
    # 30/360 from 28 April to 28 May is 30 days to redemption, not the 32
    # left of a period that began on 28 February.
    args = _bond(book, [38835.0, 38865.0, 0.0, 105.93784568694602, 92.60490729347165, 4.0, 0.0])
    assert sheet.evaluate(f"YIELD({args})") == -1.5102748190150097


def test_oddlprice_steps_quasi_coupon_dates_from_the_last_interest_date(book: _Book) -> None:
    sheet = book["Data"]
    last = "DATE(2021,8,31)"
    # Each date a quarter after the one before: 30 November, 28 February,
    # then 28 May. From a month end, that last covers any maturity in May.
    assert sheet.evaluate(f"ODDLPRICE(DATE(2021,9,10),DATE(2022,5,30),{last},0.1,0.1,100,4,1)") == 100.0338820483653
    assert sheet.evaluate(f"ODDLPRICE(DATE(2021,9,10),DATE(2022,5,27),{last},0.1,0.1,100,4,1)") == 99.98155503258074
    assert sheet.evaluate("ODDLPRICE(DATE(2011,10,26),DATE(2012,7,17),DATE(2010,12,31),0.05,0.06,100,2,1)") \
        == 99.13490142042163


def test_oddlprice_on_basis_0_counts_month_ends_as_the_30th_in_dc(book: _Book) -> None:
    sheet = book["Data"]
    # 13 October to 31 March is 167 days, where YEARFRAC's rules count 168;
    # the quasi period from 28 February is 178 days long, not 180.
    assert sheet.evaluate("ODDLPRICE(DATE(2032,11,19),DATE(2033,3,31),DATE(2032,10,13),0.1,0,100,2,0)") \
        == 103.63888888888889
    assert sheet.evaluate("ODDLPRICE(DATE(2018,4,2),DATE(2018,5,14),DATE(2018,2,28),0.1,0.1,100,2,0)") \
        == 99.9895189314301


def test_oddlyield(book: _Book) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("ODDLYIELD(DATE(2021,2,3),DATE(2021,7,23),DATE(2021,1,1),0.05,98,100,2,0)") \
        == 0.0938122427300492
    # 30/360 counts no days from 30 July to 31 July: the yield is 0.
    assert sheet.evaluate("ODDLYIELD(DATE(2021,7,30),DATE(2021,7,31),DATE(2021,6,15),0.05,98,100,2,0)") == 0


def test_oddfprice_with_a_short_or_long_first_period(book: _Book) -> None:
    sheet = book["Data"]
    short = "DATE(2021,2,3),DATE(2023,6,30),DATE(2021,1,1),DATE(2021,6,30)"
    assert sheet.evaluate(f"ODDFPRICE({short},0.05,0.06,100,2,0)") == 97.78308404154141
    assert sheet.evaluate(f"ODDFYIELD({short},0.05,98,100,2,0)") == 0.059007229164258115
    # The coupon period comes from the first coupon's schedule, 30 November
    # to 28 February, not the maturity's 29 November.
    assert sheet.evaluate("ODDFPRICE(DATE(2019,2,5),DATE(2023,11,29),DATE(2018,12,1),DATE(2019,2,28),0.1,0,100,4,1)") \
        == 148.13888888888889
    # A long first period to a month-end first coupon discounts one period
    # more, unless settlement shares its month with a quasi-coupon date.
    long = "DATE(2027,4,30),DATE(2021,10,29),DATE(2023,4,30),0.1,0.1,100,1,0"
    assert sheet.evaluate(f"ODDFPRICE(DATE(2021,11,3),{long})") == 90.60114237553395
    assert sheet.evaluate(f"ODDFPRICE(DATE(2022,4,12),{long})") == 99.54578216866932
    mid_month = "DATE(2022,4,12),DATE(2027,6,15),DATE(2021,10,29),DATE(2023,6,15),0.1"
    assert sheet.evaluate(f"ODDFPRICE({mid_month},0.1,100,1,0)") == 99.43079730491951
    assert sheet.evaluate(f"ODDFYIELD({mid_month},90,100,1,0)") == 0.12551573318812853


def test_oddfprice_refusals(book: _Book) -> None:
    sheet = book["Data"]
    # An odd first period exactly one period long.
    odd = "DATE(2021,2,1),DATE(2023,6,30),DATE(2020,12,31),DATE(2021,6,30)"
    assert sheet.evaluate(f"ODDFPRICE({odd},0.05,0.06,100,2,0)") == CellError("#NUM!")
    # A first coupon that is not one of the maturity's coupon dates.
    off = "DATE(2021,2,3),DATE(2023,6,30),DATE(2021,1,1),DATE(2021,6,29)"
    assert sheet.evaluate(f"ODDFPRICE({off},0.05,0.06,100,2,0)") == CellError("#NUM!")


def test_pmt_divides_by_1_less_the_discount_and_multiplies_by_the_rate(book: _Book) -> None:
    sheet = book["Data"]
    # w = 1 - (1 + r)^-n, and -((pv + fv)/w - fv) r: the exact payment,
    # rounded once, is a unit or two away in each of these.
    assert sheet.evaluate("PMT(0.01,12,1000,0,0)") == -88.84878867834168
    assert sheet.evaluate("PMT(-0.02,60,1000,0,0)") == -8.471904731198892
    assert sheet.evaluate("PMT(1E-20,12,1000)") == -83.33333333333334
    # At the start of each period r becomes 1/(1/r + 1).
    assert sheet.evaluate("PMT(0.01,360,0,5000,1)") == -1.4164651943319049
    assert sheet.evaluate("PMT(0.3,30,250000,5000,1)") == -57714.77666967648


def test_pmt_errors(book: _Book) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("PMT(0.05,0,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PMT(-1,12,1000)") == CellError("#NUM!")
    # (1 + r)^-n overflows.
    assert sheet.evaluate("PMT(-0.05,1000000,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PMT(0,7,0.1,0.2,1)") == -0.042857142857142864


def test_ipmt_and_ppmt_form_the_balance_and_the_discount_as_pmt_does(book: _Book) -> None:
    sheet = book["Data"]
    # The balance going into the period, (pv + fv)/w_n w_m - fv with m
    # payments left, times the rate: the exact interest, rounded once, is
    # 0.9990000000000009 and -8.562690475450857.
    assert sheet.evaluate("IPMT(-0.999,2,10,1000)") == 0.9990000000000073
    assert sheet.evaluate("IPMT(0.01,2.5,10,1000)") == -8.562690475450859
    # The principal is (pv + fv)/w_n (1 + r)^-m times the rate, not the
    # payment less the interest.
    assert sheet.evaluate("PPMT(0.01,1,10,1000)") == -95.58207655117134
    assert sheet.evaluate("PPMT(1E-20,2,12,1000)") == -83.33333333333334
    # Paid at the start of each period, the first payment is all principal.
    assert sheet.evaluate("IPMT(0.01,1,10,1000,0,1)") == 0
    assert sheet.evaluate("PPMT(0.01,1,10,1000,0,1)") == -104.5367094566053


def test_ipmt_takes_a_period_up_to_the_term_plus_one(book: _Book) -> None:
    sheet = book["Data"]
    assert sheet.evaluate("IPMT(0.01,10.5,10,1000)") == -0.5239837631578298
    assert sheet.evaluate("IPMT(0.01,11,10,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PPMT(0.01,0.99,10,1000)") == CellError("#NUM!")
    # A rate of -1 or below is refused, as PMT refuses it.
    assert sheet.evaluate("IPMT(-1,2,10,1000)") == CellError("#NUM!")
    assert sheet.evaluate("PPMT(-1.5,2,10,1000)") == CellError("#NUM!")


def test_cumipmt_and_cumprinc_in_closed_form(book: _Book) -> None:
    sheet = book["Data"]
    # The principal repaid is pv/w_n (1 + r)^-(n - end + 1) w_count, times
    # 1 + r when payments fall at the end of each period, and the interest
    # is the payments less it: the exact sum of the periods, rounded once,
    # is -55.82076551171361, -1000 and -6.5e-17.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1,10,0)") == -55.82076551171349
    assert sheet.evaluate("CUMPRINC(5,10,1000,1,10,0)") == -1000.0000000000001
    assert sheet.evaluate("CUMIPMT(1E-20,12,1000,1,12,0)") == -1.1368683772161603e-13
    # At the start of each period a first payment in the span is all
    # principal, added on its own.
    assert sheet.evaluate("CUMPRINC(0.01,10,1000,1,2,1)") == -200.11878600777663
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1,2,1)") == -8.95463290543394


def test_cumipmt_rounds_a_fractional_start_up_and_end_down(book: _Book) -> None:
    sheet = book["Data"]
    # From period 2 and to period 9.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,1.1,10,0)") == -45.820765511713375
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,2,9.9,0)") == -44.77539841714736
    assert sheet.evaluate("CUMPRINC(0.01,10,1000,1.4,1.4,1)") == 0
    # The checks read the periods as given.
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,0.9,10,0)") == CellError("#NUM!")
    assert sheet.evaluate("CUMIPMT(0.01,10.5,1000,1,10.7,0)") == CellError("#NUM!")
    assert sheet.evaluate("CUMIPMT(0.01,10,1000,2.5,2.4,0)") == CellError("#NUM!")


def test_below_1_the_hyperbolic_functions_take_e_to_the_x_less_1_by_kahans_trick(book: _Book) -> None:
    sheet = book["Data"]
    # a = e^x - 1 and b = e^-x - 1, each (u - 1) x / ln u with u = EXP(x):
    # SINH is (a - b)/2, TANH (a - b)/(a + b + 2).
    assert sheet.evaluate("SINH(0.7)") == 0.7585837018395336
    assert sheet.evaluate("SINH(-0.01)") == -0.010000166667500001
    assert sheet.evaluate("TANH(0.001)") == 0.0009999996666668002
    assert sheet.evaluate("TANH(0.4)") == 0.3799489622552249


def test_weibull_takes_its_distribution_as_1_less_e_to_the_minus_t_below_1(book: _Book) -> None:
    sheet = book["Data"]
    # -(e^-t - 1) by Kahan's trick, not 1 - EXP(-t), up to t = 1.
    assert sheet.evaluate("WEIBULL.DIST(0.1,1,1,TRUE)") == 0.09516258196404045
    assert sheet.evaluate("WEIBULL.DIST(0.85,1,1,TRUE)") == 0.5725850680512734


#: Sheet In: orders with every kind of key the grouping functions meet.
ORDERS: list[tuple[CellValue, ...]] = [
    ("Region", "Product", "Year", "Qty", "Price", "Flag"),
    ("East", "Pen", 2023, 3, 1.5, True), ("West", "Book", 2024, 1, 12.0, False),
    ("East", "Ink", 2024, 5, 4.25, True), ("North", "Pen", 2023, 2, 1.5, True),
    ("West", "Pen", 2023, 4, 1.25, True), ("east", "Book", 2023, 2, 11.5, False),
    ("South", "Ink", 2024, 7, 4.0, True), ("North", "Book", 2024, 1, 13.0, True),
    ("West", "Ink", 2023, 6, 4.5, False), ("East", "Pen", 2024, 8, 1.75, True),
    (None, "Pen", 2024, 2, 1.5, True), ("South", None, 2023, 3, 4.0, True),
    ("North", "Ink", 2023, None, 4.25, True), ("East", "Book", 2024, 2, 12.5, False),
    ("West", "Book", 2023, "x", 12.0, True), (7, "Pen", 2024, 1, 1.5, True),
    ("South", "Pen", 2024, 9, 1.25, False), ("North", "Pen", 2024, 3, 1.5, True),
]  # fmt: skip

#: Sheet Pv: groups whose key order, total order and leaves' orders all
#: differ. a: x = 4 + 6, y = 1; b: x = 3, y = 4; c: x = 2, y = 20; d: x = 9,
#: y = 6.
LEVELS: list[tuple[CellValue, ...]] = [
    ("K1", "K2", "T", "V"), ("c", "y", "t", 20), ("a", "x", "t", 4), ("d", "y", "t", 6), ("b", "x", "t", 3),
    ("a", "y", "t", 1), ("c", "x", "t", 2), ("d", "x", "t", 9), ("b", "y", "t", 4), ("a", "x", "t", 6),
]  # fmt: skip


def _grouped(book: _Book) -> None:
    """The sheets the grouping tests read: In, Pv, and Mw, where each leaf
    of A/B/C by p/q has one row and value columns worth 1, 100, 10000 and
    1000000 times the powers of two, so that every sum says what it added."""
    for name, rows in (("In", ORDERS), ("Pv", LEVELS)):
        sheet = book.add_sheet(name)
        for row, values in enumerate(rows, start=1):
            for column, value in enumerate(values, start=1):
                if value is not None:
                    sheet[f"{column_letter(column)}{row}"] = value
    sheet = book.add_sheet("Mw")
    for row, (first, second) in enumerate([("A", "p"), ("A", "q"), ("B", "p"), ("B", "q"), ("C", "p"), ("C", "q")],
                                          start=2):
        sheet[f"A{row}"] = "r"
        sheet[f"B{row}"] = first
        sheet[f"C{row}"] = second
        for offset, column in enumerate("DEFG"):
            sheet[f"{column}{row}"] = 2 ** (row - 2) * 100**offset


def _text(sheet: _Sheet, formula: str) -> object:
    """The formula's result as ARRAYTOTEXT(...,1) writes it, which is how
    Excel's results were read."""
    return sheet.evaluate(f"ARRAYTOTEXT({formula},1)")


def test_groupby_matches_keys_as_sort_does_and_spells_each_by_its_first_leaf(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    # "east" and "East" are one group, spelled from the first row of its
    # first leaf: Book, found in "east" 2023; with products descending,
    # Pen, found first in "East". A blank key sorts last either way.
    assert _text(sheet, "GROUPBY(In!A2:B19,In!D2:D19,_xleta.SUM,0,2)") == (
        '{7,"Pen",1;7,"",1;"east","Book",4;"east","Ink",5;"east","Pen",11;"east","",20;"North","Book",1;'
        '"North","Ink",0;"North","Pen",5;"North","",6;"South","Ink",7;"South","Pen",9;"South",,3;"South","",19;'
        '"West","Book",1;"West","Ink",6;"West","Pen",4;"West","",11;,"Pen",2;,"",2;"Grand Total","",59}'
    )
    assert _text(sheet, "GROUPBY(In!A2:B19,In!D2:D19,_xleta.SUM,0,1,{1,-2})") == (
        '{7,"Pen",1;"East","Pen",11;"East","Ink",5;"East","Book",4;"North","Pen",5;"North","Ink",0;'
        '"North","Book",1;"South","Pen",9;"South","Ink",7;"South",,3;"West","Pen",4;"West","Ink",6;'
        '"West","Book",1;,"Pen",2;"Total","",59}'
    )


def test_groupby_sorts_every_level_by_its_own_groups_result(book: _Book) -> None:
    _grouped(book)
    assert _text(book["Data"], "GROUPBY(Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,3)") == (
        '{"b","x",3;"b","y",4;"a","y",1;"a","x",10;"d","y",6;"d","x",9;"c","x",2;"c","y",20;"Total","",55}'
    )


def test_groupby_takes_headers_when_the_values_start_with_text_above_numbers(book: _Book) -> None:
    sheet = book["Data"]
    rows: list[tuple[CellValue, CellValue, CellValue]] = [("k", "v", "v"), ("a", None, 1), ("b", 1, "z"), ("a", 2, 2),
                                                          ("b", 3, 3)]
    for row, (key, number, text) in enumerate(rows, start=1):
        sheet[f"A{row}"] = key
        if number is not None:
            sheet[f"B{row}"] = number
        sheet[f"C{row}"] = text
    assert _text(sheet, "GROUPBY(A1:A5,B1:B5,_xleta.SUM)") == '{"a",2;"b",4;"Total",6}'
    # Text below the header is #VALUE!; text all the way down is no header.
    assert _text(sheet, "GROUPBY(A1:A5,C1:C5,_xleta.SUM)") == CellError("#VALUE!")
    assert _text(sheet, 'GROUPBY({1;2;1},{"a";"b";"c"},_xleta.SUM)') == '{1,0;2,0;"Total",0}'


def test_groupby_names_the_value_column_beside_each_stacked_function(book: _Book) -> None:
    _grouped(book)
    assert _text(book["Data"], "GROUPBY(In!A1:A19,In!D1:E19,VSTACK(_xleta.SUM,_xleta.MAX),3)") == (
        '{"Region","","","";7,"SUM","Qty",1;7,"MAX","Price",1.5;"East","SUM","Qty",20;"East","MAX","Price",12.5;'
        '"North","SUM","Qty",6;"North","MAX","Price",13;"South","SUM","Qty",19;"South","MAX","Price",4;'
        '"West","SUM","Qty",11;"West","MAX","Price",12;,"SUM","Qty",2;,"MAX","Price",1.5;"Total","SUM","Qty",59;'
        '"Total","MAX","Price",13}'
    )


def test_groupby_keeps_an_error_in_its_cell_but_refuses_an_array(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, 'GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,IF(ROWS(x)>3,VALUE("a"),1)))') == (
        '{7,1;"East",#VALUE!;"North",#VALUE!;"South",1;"West",#VALUE!;,1;"Total",#VALUE!}'
    )
    # A group reaches the function as an array, even a group of one row,
    # and a result that is an array refuses the whole summary.
    assert _text(sheet, "GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,x))") == CellError("#VALUE!")
    # COUNTBLANK takes only a range: given the group's array it gives an
    # array, one #VALUE! an item, and an array result refuses it all.
    assert _text(sheet, "GROUPBY(In!A2:A19,In!D2:D19,LAMBDA(x,COUNTBLANK(x)))") == CellError("#VALUE!")


def test_pivotby_spreads_a_field_across_the_top_with_subtotals_first(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "PIVOTBY(In!A2:A19,In!C2:C19,In!E2:E19,_xleta.SUM)") == (
        '{"",2023,2024,"Total";7,"",1.5,1.5;"East",13,18.5,31.5;"North",5.75,14.5,20.25;"South",4,5.25,9.25;'
        '"West",17.75,12,29.75;,"",1.5,1.5;"Total",40.5,53.25,93.75}'
    )
    assert _text(sheet, "PIVOTBY(In!A2:A19,In!B2:C19,In!E2:E19,_xleta.SUM,0,1,,-2)") == (
        '{"","Grand Total","Book","Book","Book","Ink","Ink","Ink","Pen","Pen","Pen",,;'
        '"","","",2023,2024,"",2023,2024,"",2023,2024,"",2023;7,1.5,"","","","","","",1.5,"",1.5,"","";'
        '"East",31.5,24,11.5,12.5,4.25,"",4.25,3.25,1.5,1.75,"","";"North",20.25,13,"",13,4.25,4.25,"",3,1.5,1.5,"","";'
        '"South",9.25,"","","",4,"",4,1.25,"",1.25,4,4;"West",29.75,24,12,12,4.5,4.5,"",1.25,1.25,"","","";'
        ',1.5,"","","","","","",1.5,"",1.5,"","";"Total",93.75,61,23.5,37.5,17,8.75,8.25,11.75,4.25,7.5,4,4}'
    )


def test_pivotby_shows_the_column_fields_names_as_one_text(book: _Book) -> None:
    _grouped(book)
    assert _text(book["Data"], "PIVOTBY(In!A1:B19,In!B1:C19,In!E1:E19,_xleta.SUM,3)") == (
        '{"","","Product, Year","","","","","","","";"","","Book","Book","Ink","Ink","Pen","Pen",,"Total";'
        '"","",2023,2024,2023,2024,2023,2024,2023,"";"Region","Product","Price","Price","Price","Price","Price",'
        '"Price","Price","Price";7,"Pen","","","","","",1.5,"",1.5;"east","Book",11.5,12.5,"","","","","",24;'
        '"east","Ink","","","",4.25,"","","",4.25;"east","Pen","","","","",1.5,1.75,"",3.25;'
        '"North","Book","",13,"","","","","",13;"North","Ink","","",4.25,"","","","",4.25;'
        '"North","Pen","","","","",1.5,1.5,"",3;"South","Ink","","","",4,"","","",4;'
        '"South","Pen","","","","","",1.25,"",1.25;"South",,"","","","","","",4,4;"West","Book",12,12,"","","","","",24;'
        '"West","Ink","","",4.5,"","","","",4.5;"West","Pen","","","","",1.25,"","",1.25;'
        ',"Pen","","","","","",1.5,"",1.5;"Total","",23.5,37.5,8.75,8.25,4.25,7.5,4,93.75}'
    )


def test_pivotby_relative_to_the_parent_row(book: _Book) -> None:
    _grouped(book)
    # ROWS(y) counts the rows of the cell's parent row, the grand total's
    # for a subtotal, its own for the grand total.
    formula = "PIVOTBY(In!A2:B19,In!B2:C19,In!E2:E19,LAMBDA(x,y,ROWS(y)),0,2,,2,,,4)"
    assert _text(book["Data"], formula) == (
        '{"","","Book","Book","Book","Ink","Ink","Ink","Pen","Pen","Pen",,,"Grand Total";'
        '"","",2023,2024,"",2023,2024,"",2023,2024,"",2023,"","";7,"Pen","","","","","","","",1,1,"","",1;'
        '7,"","","","","","","","",5,8,"","",18;"east","Book",1,1,2,"","","","","","","","",5;'
        '"east","Ink","","","","",1,1,"","","","","",5;"east","Pen","","","","","","",1,1,2,"","",5;'
        '"east","",2,3,5,"",2,4,3,5,8,"","",18;"North","Book","",1,1,"","","","","","","","",4;'
        '"North","Ink","","","",1,"",1,"","","","","",4;"North","Pen","","","","","","",1,1,2,"","",4;'
        '"North","","",3,5,2,"",4,3,5,8,"","",18;"South","Ink","","","","",1,1,"","","","","",3;'
        '"South","Pen","","","","","","","",1,1,"","",3;"South",,"","","","","","","","","",1,1,3;'
        '"South","","","","","",2,4,"",5,8,1,1,18;"West","Book",1,1,2,"","","","","","","","",4;'
        '"West","Ink","","","",1,"",1,"","","","","",4;"West","Pen","","","","","","",1,"",1,"","",4;'
        '"West","",2,3,5,2,"",4,3,"",8,"","",18;,"Pen","","","","","","","",1,1,"","",1;'
        ',"","","","","","","","",5,8,"","",18;"Grand Total","",2,3,5,2,2,4,3,5,8,1,1,18}'
    )


def test_pivotby_sorts_by_a_value_only_where_there_is_a_total_to_sort_by(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    # Across the top the leaves sort by their totals, but a, b, c and d,
    # having no columns of their own, keep key order until subtotals show.
    assert _text(sheet, "PIVOTBY(Pv!C2:C10,Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,,1,3)") == (
        '{"","a","a","b","b","c","c","d","d","Total";"","y","x","x","y","x","y","y","x","";'
        '"t",1,10,3,4,2,20,6,9,55;"Total",1,10,3,4,2,20,6,9,55}'
    )
    assert _text(sheet, "PIVOTBY(Pv!C2:C10,Pv!A2:B10,Pv!D2:D10,_xleta.SUM,0,1,,2,3)") == (
        '{"","b","b","b","a","a","a","d","d","d","c","c","c","Grand Total";"","x","y","","y","x","","y","x","","x","y","","";'
        '"t",3,4,7,1,10,11,6,9,15,2,20,22,55;"Total",3,4,7,1,10,11,6,9,15,2,20,22,55}'
    )
    # Down the side every level sorts, but only beside a total column.
    assert _text(sheet, "PIVOTBY(Pv!A2:B10,Pv!C2:C10,Pv!D2:D10,_xleta.SUM,0,1,-3)") == (
        '{"","","t","Total";"c","y",20,20;"c","x",2,2;"d","x",9,9;"d","y",6,6;"a","x",10,10;"a","y",1,1;'
        '"b","y",4,4;"b","x",3,3;"Total","",55,55}'
    )
    assert _text(sheet, "PIVOTBY(Pv!A2:A10,Pv!B2:B10,Pv!D2:D10,_xleta.SUM,0,1,2,0)") == (
        '{"","x","y";"a",10,1;"b",3,4;"c",2,20;"d",9,6;"Total",24,31}'
    )


def test_pivotby_misplaces_one_functions_totals_over_several_value_columns(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    # Each total goes one column along a node, the grand total first, and
    # the leaves are written over them: of the grand total (63 and 6300)
    # only 6300 survives, in B's subtotal.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!B2:C7,Mw!D2:E7,_xleta.SUM,0,0,,2)") == (
        '{"","A","A","A","A","A","A","B","B","B","B","B","B","C","C","C","C","C","C","Grand Total","Grand Total";'
        '"","p","p","q","q","","","p","p","q","q","","","p","p","q","q","","","","";'
        '"r",1,100,2,200,"","",4,400,8,800,6300,"",16,1600,32,3200,"","","",""}'
    )
    # Totals first move the columns as they were placed.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!B2:C7,Mw!D2:G7,_xleta.SUM,0,0,,-2)") == (
        '{"","Grand Total","Grand Total","Grand Total","Grand Total","A","A","A","A","A","A","A","A","A","A","A","A",'
        '"B","B","B","B","B","B","B","B","B","B","B","B","C","C","C","C","C","C","C","C","C","C","C","C";'
        '"","","","","","","","","","p","p","p","p","q","q","q","q","","","","","p","p","p","p","q","q","q","q",'
        '"","","","","p","p","p","p","q","q","q","q";"r","","","","","",63,6300,630000,1,100,10000,1000000,'
        '2,200,20000,2000000,"","","","",4,400,40000,4000000,8,800,80000,8000000,"","","","",'
        '16,1600,160000,16000000,32,3200,320000,32000000}'
    )
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!A2:A7,Mw!D2:E7,_xleta.SUM)") == (
        '{"","r","r","Total","Total";"r",63,6300,6300,"";"Total",63,6300,6300,""}'
    )
    # Two functions over the two columns total as they should.
    assert _text(sheet, "PIVOTBY(Mw!A2:A7,Mw!A2:A7,Mw!D2:E7,HSTACK(_xleta.SUM,_xleta.SUM),0,0)") == (
        '{"","r","r","Total","Total";"","SUM","SUM","SUM","SUM";"r",63,6300,63,6300}'
    )


def test_pivotby_names_functions_and_their_value_columns(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "PIVOTBY(In!A1:A19,In!C1:C19,In!E1:E19,HSTACK(_xleta.SUM,_xleta.MAX),3)") == (
        '{"","Year","","","","","";"",2023,2023,2024,2024,"Total","Total";"","SUM","MAX","SUM","MAX","SUM","MAX";'
        '"Region","Price","Price","Price","Price","Price","Price";7,"","",1.5,1.5,1.5,1.5;'
        '"East",13,11.5,18.5,12.5,31.5,12.5;"North",5.75,4.25,14.5,13,20.25,13;"South",4,4,5.25,4,9.25,4;'
        '"West",17.75,12,12,12,29.75,12;,"","",1.5,1.5,1.5,1.5;"Total",40.5,12,53.25,13,93.75,13}'
    )
    # Stacked over two value columns, each row names its value column and
    # the row fields' names move up beside the last row of keys.
    formula = "PIVOTBY(In!A1:A19,In!C1:C19,In!D1:E19,VSTACK(_xleta.SUM,_xleta.MAX),3)"
    assert _text(sheet, formula) == (
        '{"","","","Year","","";"Region","","",2023,2024,"Total";7,"SUM","Qty","",1,1;7,"MAX","Price","",1.5,1.5;'
        '"East","SUM","Qty",5,15,20;"East","MAX","Price",11.5,12.5,12.5;"North","SUM","Qty",2,4,6;'
        '"North","MAX","Price",4.25,13,13;"South","SUM","Qty",3,16,19;"South","MAX","Price",4,4,4;'
        '"West","SUM","Qty",10,1,11;"West","MAX","Price",12,12,12;,"SUM","Qty","",2,2;,"MAX","Price","",1.5,1.5;'
        '"Total","SUM","Qty",20,39,59;"Total","MAX","Price",12,13,13}'
    )


def test_a_lambdas_arguments_are_counted_and_may_be_optional(book: _Book) -> None:
    sheet = book["Data"]
    optional = "_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),\"om\",_xlpm.y))"
    assert sheet.evaluate(f"{optional}(1)") == "om"
    assert sheet.evaluate(f"{optional}(1,5)") == 5
    assert sheet.evaluate(f"{optional}(1,)") == "om"
    assert sheet.evaluate(f"{optional}(1,2,3)") == CellError("#VALUE!")
    assert sheet.evaluate("LAMBDA(x,x+1)()") == CellError("#VALUE!")
    assert sheet.evaluate("REDUCE(0,{1,2},_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlop.c,_xlpm.a+_xlpm.b))") == 3
    # PERCENTOF needs two arguments where BYROW gives one.
    assert sheet.evaluate("BYROW({1,2;3,4},_xleta.PERCENTOF)") == CellError("#VALUE!")


def test_counta_counts_every_item_of_an_array_but_only_filled_cells(book: _Book) -> None:
    sheet = book["Data"]
    sheet["C1"] = "a"
    sheet["C3"] = "c"
    assert sheet.evaluate("COUNTA(C1:C3)") == 2
    assert sheet.evaluate("COUNTA(VSTACK(C1:C3))") == 3
    # A blank cell for ignore_empty is FALSE; left out, it is TRUE.
    assert sheet.evaluate('TEXTJOIN("|",,C1:C3)') == "a|c"
    assert sheet.evaluate('TEXTJOIN("|",C2,C1:C3)') == "a||c"
    assert sheet.evaluate('TEXTJOIN("|",C5:C6,C1:C3)') == CellError("#VALUE!")


def test_a_parameter_only_a_range_will_do_takes_an_array_one_item_at_a_time(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "LET(q,SEQUENCE(2,2),COUNTIF(q,1))") == "{#VALUE!,#VALUE!;#VALUE!,#VALUE!}"
    assert _text(sheet, "LET(a,SEQUENCE(3),COUNTIF(a,{1,2}))") == "{#VALUE!,#VALUE!;#VALUE!,#VALUE!;#VALUE!,#VALUE!}"
    assert sheet.evaluate("LET(s,5,COUNTIF(s,1))") == CellError("#VALUE!")
    assert sheet.evaluate("COUNTBLANK((In!D2:D4,In!E2:E4))") == CellError("#VALUE!")


def test_take_and_drop_of_a_range_are_ranges(book: _Book) -> None:
    _grouped(book)
    sheet = book["Data"]
    assert _text(sheet, "ROW(TAKE(In!D2:D10,3))") == "{2;3;4}"
    assert sheet.evaluate("COUNTIF(DROP(In!D2:E10,1,-1),1)") == 2
    assert sheet.evaluate("ISREF(TAKE(In!D2:D10,20))") is True
    assert sheet.evaluate("ISREF(TAKE({1,2},1))") is False
    assert sheet.evaluate("TAKE((In!D2:D4,In!E2:E4),1)") == CellError("#VALUE!")


def test_trimrange_trims_an_array_of_its_blank_edges(book: _Book) -> None:
    sheet = book["Data"]
    assert _text(sheet, "TRIMRANGE(VSTACK(Z1,1,Z2))") == "{1}"
    assert _text(sheet, "TRIMRANGE(HSTACK(Z1,1,Z2),,2)") == "{,1}"
    assert sheet.evaluate("TRIMRANGE(VSTACK(Z1,Z2))") == CellError("#VALUE!")


def test_transpose_in_a_formula_before_dynamic_arrays_cuts_its_range_to_the_row(book: _Book) -> None:
    sheet = book["Data"]
    for row in (1, 2, 3):
        sheet[f"A{row}"] = row
    # Measured with the formulas entered as Excel before dynamic arrays
    # took them: the range is cut to the formula's row, as a single value.
    sheet["C1"].formula = "TRANSPOSE(A1:A2)"
    sheet["C2"].formula = "TRANSPOSE(A1:A3)"
    sheet["C5"].formula = "SUM(TRANSPOSE(A1:A3))"
    sheet["C6"].formula = "INDEX(TRANSPOSE(A1:A3),1,2)"
    book.calculate()
    assert [sheet[f"C{row}"].value for row in (1, 2, 5, 6)] == [1, 2, CellError("#VALUE!"), CellError("#VALUE!")]


def test_bahttext_spells_an_amount_as_excel_does(book: _Book) -> None:
    sheet = book["Data"]
    # A 1 in the ones is "et" after any other digit, those before a
    # million included; a million's multiple is spelled on its own.
    assert sheet.evaluate("BAHTTEXT(21)") == "ยี่สิบเอ็ดบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1000001)") == "หนึ่งล้านเอ็ดบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1000001000000)") == "หนึ่งล้านเอ็ดล้านบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(1E+100)") == "หนึ่งหมื่น" + "ล้าน" * 16 + "บาทถ้วน"
    # Satang are rounded as ROUND rounds, and stand alone below a baht.
    assert sheet.evaluate("BAHTTEXT(1.005)") == "หนึ่งบาทหนึ่งสตางค์"
    assert sheet.evaluate("BAHTTEXT(0.995)") == "หนึ่งบาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(0.21)") == "ยี่สิบเอ็ดสตางค์"
    # A negative amount keeps its minus even when it rounds to nothing.
    assert sheet.evaluate("BAHTTEXT(-0.004)") == "ลบศูนย์บาทถ้วน"
    assert sheet.evaluate("BAHTTEXT(Z99)") == CellError("#VALUE!")


def test_sheet_and_sheets_given_what_is_not_a_range(book: _Book) -> None:
    sheet = book["Data"]
    assert sheet.evaluate('SHEET("Data")') == 1
    assert sheet.evaluate("SHEET(1)") == CellError("#N/A")
    assert sheet.evaluate('SHEET({"Data"})') == CellError("#N/A")
    assert sheet.evaluate('SHEETS("Data")') == CellError("#N/A")
    assert sheet.evaluate("SHEET(1/0)") == CellError("#DIV/0!")
