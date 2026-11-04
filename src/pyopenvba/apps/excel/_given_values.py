"""What Excel makes of a Currency, a Date or a range of separate areas given to a property, before the property's own setter
sees it (scripts/measure_given_values.py, tests/fixtures/given_values.json). Typed and late bound answer alike.

The type library declares these properties Variant, so what a property is shows in what it answers:

- A colour or a shade converts as VBA does, so a Currency rounds and separate areas read as the first: it is left alone.
- A property that holds a formula or a value reads separate areas as the first, too: it is left alone.
- Any other property refuses separate areas with 1004, and takes a block of cells as its last cell's value.
- A border's Weight and LineStyle, and a font's Background, read a number cut to a whole one, a block as its first cell.
- One that answers a Boolean also refuses a Currency and a Date with 1004.
- One that answers text takes a Currency's text, $1.50 with a space after it. NumberFormat refuses a Date.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from typing import Final

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY, NULL, VBACurrency, VBADate, VBASingle, error, to_number

#: The classes the rule was measured on.
MEASURED: Final = frozenset({'Range', 'Font', 'Interior', 'Border', 'Borders', 'Style', 'Characters'})
#: Convert as VBA does, whatever they answer.
CONVERTING: Final = frozenset({'color', 'colorindex', 'themecolor', 'tintandshade', 'patterncolor', 'patterncolorindex', 'patternthemecolor', 'patterntintandshade', 'themefont'})
#: Hold what they are given.
HOLDING: Final = frozenset({'value', 'value2', 'formula', 'formular1c1', 'formulalocal', 'formular1c1local', 'formula2', 'formula2r1c1', 'formula2local', 'formula2r1c1local', 'formulaarray'})
NUMBER_FORMATS: Final = frozenset({'numberformat', 'numberformatlocal'})
#: Read a number cut to a whole one, and a block as its first cell.
CUTTING: Final = frozenset({'weight', 'linestyle', 'background', 'underline'})
#: Excel's own numbers for its alignments.
ALIGNMENT: Final = {'horizontalalignment': -4131, 'verticalalignment': -4108}
BACKGROUNDS: Final = frozenset({1, 2, 3, -4105})


def currency_text(value: VBACurrency) -> str:
    """A Currency as Excel reads it: $4.00 with a space after it, ($4.00) below nothing, to the nearest cent, half away from
    nothing."""
    cents = abs(Decimal(value)).quantize(Decimal('0.01'), ROUND_HALF_UP)
    return f'(${cents:,.2f})' if value < 0 else f'${cents:,.2f} '


def given(target: VBAObject, name: str, value: object) -> object:
    """``value`` as the property ``name`` of ``target`` is to be given it."""
    from pyopenvba.apps.excel._model import Range

    lower = name.lower()
    if target.vba_type_name not in MEASURED:
        return value
    if lower == 'value' and target.vba_type_name in ('Border', 'Borders'):
        # A border's Value is its LineStyle.
        lower = 'linestyle'
    if isinstance(value, Range) and value.areas[0].rows * value.areas[0].columns > 1:
        # A first area of several cells is nothing to hold, and no shade.
        if lower in HOLDING:
            return EMPTY
        if lower == 'tintandshade':
            raise error(5)
    if lower in CONVERTING or lower in HOLDING:
        return value
    refusal = error(1004, f'Unable to set the {name} property of the {target.vba_type_name} class')
    if isinstance(value, Range):
        if len(value.areas) > 1:
            # A style's name is looked for in the first cell: nothing there says Normal, and a name it lacks is 450.
            if lower == 'style':
                first = Range(value.sheet, list(value.areas[:1])).vba_value()
                if first is EMPTY:
                    return 'Normal'
                raise error(450)
            raise refusal
        if value.single:
            value = value.vba_value()
        else:
            area = value.first
            row, column = (area.top, area.left) if lower in CUTTING else (area.bottom, area.right)
            value = Range(value.sheet, [Area(row, column, row, column, value.sheet.name)]).vba_value()
    if lower in CUTTING and isinstance(value, (float, VBACurrency, VBADate, VBASingle, Decimal)) and not isinstance(value, bool):
        # A date's number, as it stands in a cell.
        whole = int(float(to_number(value)))
        if lower == 'background' and whole not in BACKGROUNDS:
            raise refusal
        return whole
    if lower in ALIGNMENT and not isinstance(value, bool) and value == 2:
        return ALIGNMENT[lower]
    if not isinstance(value, (VBACurrency, VBADate)):
        return value
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
