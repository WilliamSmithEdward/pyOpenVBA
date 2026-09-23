"""Excel's arithmetic, where it is not IEEE's.

Excel computes in doubles, and most of the time a formula's result is the
double IEEE arithmetic gives. Five rules, each measured in
``tests/fixtures/excel/formulas.xlsx``, make it differ:

**Two roundings.** Each operation happens in an x87 register, rounded to a
64-bit mantissa, and is then rounded again to a double; once in about two
thousand operations that lands on the other double. :mod:`.precise` has
the operations.

**Fifteen digits.** A number becomes text with fifteen significant digits,
so ``0.1+0.2`` is ``"0.3"`` after ``&""``. The text is plain while it fits
in twenty characters and scientific after that: ``1E+20``, but
``12345678901234600000``; ``0.000000000000000001``, but ``1E-19``.

**Equality to fifteen digits.** ``=`` and ``<`` compare two numbers as
they read to fifteen significant digits, so ``0.1+0.2=0.3`` is TRUE while
the two doubles differ. Rounding both, rather than allowing a tolerance, is
what the measurements need: ``1`` equals the double four units below it and
not the one five below, and equals the one twenty units above it and not
twenty-four, because those are where the fifteenth digit changes.

**A last subtraction near zero is zero.** When the operation a formula ends
with is ``+`` or ``-``, a result less than eight units in the last place of
the left operand becomes 0: ``=0.1+0.2-0.3`` is 0. The same subtraction
anywhere else keeps its rounding error, so ``=(0.1+0.2-0.3)`` and
``=1*(0.1+0.2-0.3)`` are ``5.55E-17``. SUM does the same with its last
addition, wherever it stands.

**Neither too small nor too big.** A result below the smallest normal
double is 0, and one that overflows is ``#NUM!``. A number literal keeps
fifteen significant digits, cut rather than rounded: ``=2.9999999999999996``
is ``2.99999999999999``. So does a number read from text.
"""

from __future__ import annotations

import math
import sys
from collections.abc import Iterable
from decimal import ROUND_DOWN, ROUND_HALF_DOWN, ROUND_HALF_UP, Decimal

from pyopenvba.formula._calc.precise import add

#: The smallest positive normal double. Excel has no subnormals.
SMALLEST = 2.2250738585072014e-308
#: The largest number Excel accepts as typed.
LARGEST_TYPED = 9.99999999999999e307
#: How many units in the last place a final sum may differ from zero and
#: still be taken for it.
_NEAR_ZERO_UNITS = 8
#: The longest a number's plain text may be before it is written with an
#: exponent, not counting a minus sign.
_PLAIN_WIDTH = 20
_DOUBLE_MAX = Decimal(sys.float_info.max)


def normal(value: float) -> float:
    """``value`` with a subnormal result flushed to zero, as Excel does."""
    if value != 0.0 and -SMALLEST < value < SMALLEST:
        return 0.0
    return value


def _significant(value: float, digits: int) -> tuple[str, int]:
    """``value``'s first ``digits`` significant digits, without trailing
    zeros, and the power of ten of the first.

    An exact half rounds toward zero: 1234567890123445 is
    ``1234567890123440`` and 1234567890123455 ``1234567890123450``, both
    measured. Anything past a half rounds up.
    """
    exact = Decimal(abs(value))
    exponent = exact.adjusted()
    rounded = exact.quantize(Decimal(1).scaleb(exponent - digits + 1), rounding=ROUND_HALF_DOWN)
    if rounded > _DOUBLE_MAX:
        # Fifteen digits of the largest double overflow it; Excel shows
        # fourteen there.
        return _significant(value, digits - 1)
    exponent = rounded.adjusted()
    text = str(rounded.scaleb(-exponent).quantize(Decimal(1).scaleb(-(digits - 1))))
    mantissa = text.replace(".", "").rstrip("0") or "0"
    return mantissa, exponent


def number_text(value: float) -> str:
    """A number as Excel turns it into text: for ``&``, and wherever a
    function wants text and is given a number."""
    if value == 0.0:
        return "0"
    digits, exponent = _significant(value, 15)
    sign = "-" if value < 0 else ""
    if exponent >= 0:
        whole = digits[: exponent + 1].ljust(exponent + 1, "0")
        fraction = digits[exponent + 1 :]
        plain = whole + ("." + fraction if fraction else "")
    else:
        plain = "0." + "0" * (-exponent - 1) + digits
    if len(plain) <= _PLAIN_WIDTH:
        return sign + plain
    mantissa = digits[0] + ("." + digits[1:] if len(digits) > 1 else "")
    return f"{sign}{mantissa}E{'+' if exponent >= 0 else '-'}{abs(exponent):02d}"


def _fifteen(value: float) -> Decimal:
    if value == 0.0:
        return Decimal(0)
    exact = Decimal(value)
    return exact.quantize(Decimal(1).scaleb(exact.adjusted() - 14), rounding=ROUND_HALF_UP)


def compare_numbers(left: float, right: float) -> int:
    """-1, 0 or 1 as ``left`` is below, level with or above ``right``, read
    to fifteen significant digits."""
    if left == right:
        return 0
    # Far enough apart that fifteen digits cannot bring them together.
    scale = max(abs(left), abs(right))
    if abs(left - right) > scale * 1e-13:
        return -1 if left < right else 1
    a, b = _fifteen(left), _fifteen(right)
    return (a > b) - (a < b)


def near_zero(left: float, right: float, result: float) -> float:
    """``result`` of ``left + right``, zero when it is only rounding error:
    under eight units in the last place of ``left``.

    The left operand, not the smaller: ``1-(1-8u)`` is 0 while ``(1-8u)-1``
    is not, ``u`` being a unit in the last place of 1 below it.
    """
    if result == 0.0:
        return result
    if abs(result) < _NEAR_ZERO_UNITS * math.ulp(left):
        return 0.0
    return result


def total(values: Iterable[float]) -> float:
    """Numbers added as SUM adds them: in order, with only the last
    addition taking a rounding error for zero. ``SUM(0.1,0.2,-0.3)`` is 0,
    and ``SUM(0.1,0.2,-0.3,1E-20)`` is ``5.55E-17`` plus ``1E-20``."""
    items = list(values)
    if not items:
        return 0.0
    running = 0.0
    for value in items[:-1]:
        running = add(running, value)
    last = items[-1]
    return near_zero(running, last, add(running, last))


def _cut(exact: Decimal) -> Decimal:
    """Fifteen significant digits, the rest cut off."""
    return exact.quantize(Decimal(1).scaleb(exact.adjusted() - 14), rounding=ROUND_DOWN)


def literal_value(digits: str) -> float:
    """What a number written in a formula stands for: fifteen significant
    digits of it, the rest cut off, and nothing below the smallest normal
    double."""
    exact = Decimal(digits)
    if exact == 0:
        return 0.0
    return normal(float(_cut(exact)))


#: Below this a number read from text is no number: "1e-308" is 0 and
#: "1e-400" is #VALUE!.
_TEXT_FLOOR = Decimal("1e-308")
_TEXT_CEILING = Decimal("9.99999999999999e307")


def text_value(digits: str) -> float | None:
    """What unsigned digits read from text stand for, fifteen of them as a
    literal's, or ``None`` for a number too large or too small to be one:
    ``"9.999999999999995e307"`` is ``9.99999999999999E+307`` and
    ``"1e308"`` no number."""
    exact = Decimal(digits)
    if exact == 0:
        return 0.0
    cut = _cut(exact)
    if cut > _TEXT_CEILING or cut < _TEXT_FLOOR:
        return None
    return normal(float(cut))


__all__ = [
    "LARGEST_TYPED",
    "SMALLEST",
    "compare_numbers",
    "literal_value",
    "near_zero",
    "normal",
    "number_text",
    "text_value",
    "total",
]
