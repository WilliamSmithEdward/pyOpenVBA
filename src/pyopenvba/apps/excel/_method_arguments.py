"""How Range's methods read their arguments, measured in live Excel (scripts/measure_range_methods.py,
tests/fixtures/range_methods_*.json).

An argument declared Variant reaches Excel as it was given, and Excel reads it itself: Null counts as left out, text that
is no number is error 1004 where VBA's own conversion would be 13, True is 1 and a fraction is cut, not rounded. An
argument declared as an enum is converted by VBA first where the variable is typed, so Null is 94 and a fraction rounds;
the model follows that (docs/host_completeness.md).
"""

from __future__ import annotations

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._values import MISSING, NULL, error, to_bool, to_integer


def given(value: object) -> object:
    """A Variant argument, Null counting as left out."""
    return MISSING if value is NULL else value


def enum_number(value: object, refusal: str) -> int:
    """An argument declared as an enum: VBA's conversion, but text that is no number is 1004 with ``refusal``."""
    if isinstance(value, str):
        try:
            return int(to_integer(value, 'Long'))
        except VBARuntimeError:
            raise error(1004, refusal) from None
    return int(to_integer(value, 'Long'))


def variant_number(value: object, refusal: str) -> int:
    """A number in a Variant argument as Excel reads it: True is 1 and a fraction is cut."""
    if value is True:
        return 1
    if isinstance(value, float):
        return int(value)
    return enum_number(value, refusal)


def switch(value: object, refusal: str) -> bool:
    """A Boolean in a Variant argument: text that is no Boolean is 1004 with ``refusal``."""
    try:
        return bool(to_bool(value))
    except VBARuntimeError:
        raise error(1004, refusal) from None


def search_number(value: object) -> int:
    """LookIn or SearchOrder: Null is 13, and any other value VBA's conversion, a fraction rounding."""
    if value is NULL:
        raise error(13)
    return int(to_integer(value, 'Long'))


def look_at_number(value: object) -> int:
    """LookAt: as search_number, but True is xlWhole, 1."""
    return 1 if value is True else search_number(value)


def search_switch(value: object) -> bool:
    """MatchCase: Null is 13."""
    if value is NULL:
        raise error(13)
    return bool(to_bool(value))
