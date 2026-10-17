"""VBA financial kernels, distinct from worksheet convergence and rounding.

Secant iteration structure adapted from Microsoft.VisualBasic Financial.vb:
https://github.com/dotnet/runtime/blob/2201016c1e13bdb9abf49e2e38cadf4ee0568df2/src/libraries/Microsoft.VisualBasic.Core/src/Microsoft/VisualBasic/Financial.vb
Native VBA uses forward extended-precision NPV rather than .NET's newer
reverse evaluation. Arithmetic support is shared with our worksheet engine.

Copyright (c) .NET Foundation and Contributors. All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
"""
from __future__ import annotations

import math
from fractions import Fraction

from pyopenvba.formula._calc.precise import register
from pyopenvba.interpreter import _ast as A
from pyopenvba.interpreter._values import VBAArray, error, to_integer, to_number


def number(value: object) -> float:
    result = float(to_number(value))
    if not math.isfinite(result):
        raise error(6)
    return result


def due(value: object) -> int:
    return int(to_integer(value, 'Long') != 0)


def flows(value: object) -> list[float]:
    if not isinstance(value, VBAArray) or value.element_type != 'Double':
        raise error(13)
    if len(value.bounds) != 1 or not value.items:
        raise error(5)
    return [number(item) for item in value.items]


def npv(rate: float, values: list[float], sign: int = 0) -> float:
    if rate == -1 or not values:
        raise error(5)
    discount, total = Fraction(1), Fraction(0)
    fraction_rate = Fraction(rate)
    for value in values:
        discount = register(discount + register(discount * fraction_rate))
        if sign == 0 or (sign > 0 and value > 0) or (sign < 0 and value < 0):
            total = register(total + register(Fraction(value) / discount))
    return float(total)


def irr(values: list[float], guess: float) -> float:
    if len(values) < 2 or guess <= -1:
        raise error(5)
    tolerance = max(map(abs, values)) * 1e-7 * 0.01
    previous, old_value = guess, npv(guess, values)
    current = guess + (1e-5 if old_value > 0 else -1e-5)
    if current <= -1:
        raise error(5)
    value = npv(current, values)
    for _ in range(40):
        if value == old_value:
            previous += -1e-5 if current > previous else 1e-5
            old_value = npv(previous, values)
            if value == old_value:
                raise error(5)
        found = current - (current - previous) * value / (value - old_value)
        if found <= -1:
            found = (current - 1) * 0.5
        result = npv(found, values)
        if abs(result) < tolerance and abs(found - current) < 1e-7:
            return found
        previous, current = current, found
        old_value, value = value, result
    raise error(5)


def rate(periods: float, payment: float, present: float, future: float, timing: int, guess: float) -> float:
    if periods <= 0:
        raise error(5)

    def residual(value: float) -> float:
        if value == 0:
            return present + payment * periods + future
        growth = Fraction(math.pow(1 + value, periods))
        factor = Fraction(1 + value if timing else 1)
        paid = register(register(Fraction(payment) * factor) * register(growth - 1))
        paid = register(paid / Fraction(value))
        total = register(register(Fraction(present) * growth) + paid)
        return float(register(total + Fraction(future)))

    previous, old_value = guess, residual(guess)
    current = guess / 2 if old_value > 0 else guess * 2
    value = residual(current)
    for _ in range(40):
        if value == old_value:
            previous += -1e-5 if current > previous else 1e-5
            old_value = residual(previous)
            if value == old_value:
                raise error(5)
        found = current - (current - previous) * value / (value - old_value)
        result = residual(found)
        if abs(result) < 1e-7:
            return found
        previous, current = current, found
        old_value, value = value, result
    raise error(5)


def mirr(values: list[float], finance: float, reinvest: float) -> float:
    if len(values) < 2 or finance == -1 or reinvest == -1:
        raise error(5)
    negative, positive = npv(finance, values, -1), npv(reinvest, values, 1)
    if negative == 0:
        raise error(11)
    ratio = -positive * math.pow(1 + reinvest, len(values)) / (negative * (1 + finance))
    if ratio < 0:
        raise error(5)
    return math.pow(ratio, 1 / (len(values) - 1)) - 1


def payment(rate: float, periods: float, present: float, future: float, timing: int) -> float:
    if periods == 0:
        raise error(5)
    if rate == 0:
        return -(present + future) / periods
    growth = math.pow(1 + rate, periods)
    return -(future + present * growth) / ((1 + rate * timing) * (growth - 1) / rate)


def future_value(rate: float, periods: float, paid: float, present: float) -> float:
    if rate == 0:
        return -present - paid * periods
    return -present * math.pow(1 + rate, periods) - (paid / rate) * (math.pow(1 + rate, periods) - 1)


def interest(rate: float, period: float, periods: float, present: float, future: float, timing: int) -> float:
    if period <= 0 or period >= periods + 1:
        raise error(5)
    if timing and period == 1:
        return 0.0
    paid = payment(rate, periods, present, future, timing)
    balance = future_value(rate, period - (2 if timing else 1), paid, present + (paid if timing else 0))
    return balance * rate


def _signature(name: str, required: str, optional: str = '', array: str = '') -> A.Procedure:
    params = [A.Param(name=item, declared='Double', by_val=True) for item in required.split()]
    for item in optional.split():
        default = 0.1 if item == 'Guess' else 0.0
        params.append(A.Param(name=item, declared='Long' if item == 'Due' else 'Double', by_val=True,
                              optional=True, default=A.Literal(value=default)))
    for param in params:
        if param.name == array:
            param.is_array, param.by_val = True, False
    return A.Procedure(name=name, kind='function', returns='Double', params=params)


SIGNATURES = {procedure.name.lower(): procedure for procedure in (
    _signature('Rate', 'NPer Pmt PV', 'FV Due Guess'),
    _signature('IRR', 'ValueArray', 'Guess', array='ValueArray'),
    _signature('MIRR', 'ValueArray FinanceRate ReinvestRate', array='ValueArray'),
    _signature('NPV', 'Rate ValueArray', array='ValueArray'),
    _signature('IPmt', 'Rate Per NPer PV', 'FV Due'),
    _signature('PPmt', 'Rate Per NPer PV', 'FV Due'),
    _signature('Pmt', 'Rate NPer PV', 'FV Due'),
    _signature('FV', 'Rate NPer Pmt', 'PV Due'),
    _signature('PV', 'Rate NPer Pmt', 'FV Due'),
    _signature('NPer', 'Rate Pmt PV', 'FV Due'),
)}
