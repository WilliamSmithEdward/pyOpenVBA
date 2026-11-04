"""What Excel makes of a Currency, a Date or a range of separate areas given to a property, before the property's own setter
sees it (scripts/measure_given_values.py, tests/fixtures/given_values.json). Typed and late bound answer alike.

The type library declares these properties Variant, so what a property is shows in what it answers:

- A colour or a shade converts as VBA does, so a Currency rounds and separate areas read as the first: it is left alone.
- A property that holds a formula or a value reads separate areas as the first, too: it is left alone.
- Any other property refuses separate areas with 1004.
- One that answers a Boolean also refuses a Currency and a Date with 1004.
- One that answers text takes a Currency's text, $1.50 with a space after it. NumberFormat refuses a Date.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from typing import Final

from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import NULL, VBACurrency, VBADate, error

#: The classes the rule was measured on.
MEASURED: Final = frozenset({'Range', 'Font', 'Interior', 'Border', 'Borders', 'Style', 'Characters'})
#: Convert as VBA does, whatever they answer.
CONVERTING: Final = frozenset({'color', 'colorindex', 'themecolor', 'tintandshade', 'patterncolor', 'patterncolorindex', 'patternthemecolor', 'patterntintandshade', 'themefont'})
#: Hold what they are given.
HOLDING: Final = frozenset({'value', 'value2', 'formula', 'formular1c1', 'formulalocal', 'formular1c1local', 'formula2', 'formula2r1c1', 'formula2local', 'formula2r1c1local', 'formulaarray'})
NUMBER_FORMATS: Final = frozenset({'numberformat', 'numberformatlocal'})


def currency_text(value: VBACurrency) -> str:
    """A Currency as Excel reads it: $4.00 with a space after it, ($4.00) below nothing, to the nearest cent, half away from
    nothing."""
    cents = abs(Decimal(value)).quantize(Decimal('0.01'), ROUND_HALF_UP)
    return f'(${cents:,.2f})' if value < 0 else f'${cents:,.2f} '


def given(target: VBAObject, name: str, value: object) -> object:
    """``value`` as the property ``name`` of ``target`` is to be given it."""
    lower = name.lower()
    if target.vba_type_name not in MEASURED or lower in CONVERTING or lower in HOLDING:
        return value
    separate = len(getattr(value, 'areas', ())) > 1
    if not (separate or isinstance(value, (VBACurrency, VBADate))):
        return value
    refusal = error(1004, f'Unable to set the {name} property of the {target.vba_type_name} class')
    if separate:
        raise refusal
    try:
        answer = target.vba_get(name)
    except VBARuntimeError:
        return value
    if answer is NULL:
        return value
    if isinstance(answer, bool):
        raise refusal
    if isinstance(answer, str):
        if isinstance(value, VBACurrency):
            return currency_text(value)
        if lower in NUMBER_FORMATS:
            raise refusal
    return value
