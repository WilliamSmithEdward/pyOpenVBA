"""What Excel makes of a Currency, a Date or a range of separate areas given to a property, before the property's own setter
sees it (scripts/measure_given_values.py, tests/fixtures/given_values.json). Typed and late bound answer alike.

The type library declares these properties Variant, so what a property is shows in what it answers:

- A colour or a shade converts as VBA does, so a Currency rounds and separate areas read as the first: it is left alone.
- A property that holds a formula or a value reads separate areas as the first, too: it is left alone.
- Any other property refuses separate areas with 1004, and is given a block of cells one cell at a time, row by row. It tries
  every cell, so what the others did stays done, and a refusal is reported at the end.
- A border's Weight and LineStyle read a block as its first cell.
- One that answers a Boolean also refuses a Currency and a Date with 1004.
- One that answers text takes a Currency's text, $1.50 with a space after it. NumberFormat refuses a Date.
- A number given to a pattern, a line, an underline, an alignment or a reading order goes through Excel's own numbering for it,
  so 3 is xlDot to a LineStyle and xlCenter to a HorizontalAlignment (tests/fixtures/given_value_codes.json).
- A font's Size keeps twentieths of a point and its TintAndShade thirty-two thousand and sixty-sevenths, cutting what is past them.
"""

from __future__ import annotations

import math
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Final

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import EMPTY, NULL, VBACurrency, VBADate, VBAInt, VBASingle, error, to_number

#: The classes the rule was measured on.
MEASURED: Final = frozenset({'Range', 'Font', 'Interior', 'Border', 'Borders', 'Style', 'Characters'})
#: Convert as VBA does, whatever they answer.
CONVERTING: Final = frozenset({'color', 'colorindex', 'themecolor', 'tintandshade', 'patterncolor', 'patterncolorindex', 'patternthemecolor', 'patterntintandshade', 'themefont'})
#: Hold what they are given.
HOLDING: Final = frozenset({'value', 'value2', 'formula', 'formular1c1', 'formulalocal', 'formular1c1local', 'formula2', 'formula2r1c1', 'formula2local', 'formula2r1c1local', 'formulaarray'})
NUMBER_FORMATS: Final = frozenset({'numberformat', 'numberformatlocal'})
#: Read a block as its first cell.
CUTTING: Final = frozenset({'weight', 'linestyle'})
#: Excel's own numbering, member by member, measured from -3 to 20 and over its named values.
CODES: Final = {
    'pattern': {-1: 1, 2: -4125, 3: -4126, 4: -4124, 5: -4128, 6: -4166, 7: -4121, 8: -4162, -4105: 1},
    'linestyle': {0: -4142, 2: -4115, 3: -4118, 6: 13, 7: 1, 8: -4115, 9: -4119, 10: 4, 11: 5, 12: -4119, -4125: 13, -4126: 1, -4124: -4115},
    'weight': {3: -4138},
    'underline': {1: -4142, 3: -4119},
    'horizontalalignment': {2: -4131, 3: -4108, 4: -4152, 6: -4130, 8: -4117},
    'verticalalignment': {1: -4160, 2: -4108, 3: -4107, 4: -4130, 5: -4117},
    'readingorder': {1: -5002, 2: -5003, 3: -5004},
}
#: What each takes after that numbering; anything else is refused. Pattern's refusal is the setter's own, 9.
ACCEPTED: Final = {
    'linestyle': frozenset({-4142, 1, -4115, -4118, 4, 5, 13, -4119}),
    'weight': frozenset({1, 2, -4138, 4}),
    'underline': frozenset({-4142, 2, -4119, 4, 5}),
    'horizontalalignment': frozenset({1, -4131, -4108, -4152, 5, -4130, 7, -4117}),
    'verticalalignment': frozenset({-4160, -4108, -4107, -4130, -4117}),
    'readingorder': frozenset({-5002, -5003, -5004}),
}
#: Read a Double as text.
AS_TEXT: Final = frozenset({'name', 'numberformat', 'numberformatlocal'})
#: Colours, which refuse the largest Long on a font or a border.
LONG_LIMITED: Final = frozenset({'color', 'colorindex', 'themecolor'})
#: Past this many cells a block is walked by what its sheet uses.
LARGE_BLOCK: Final = 4096
BACKGROUNDS: Final = frozenset({1, 2, 3, -4105})


def currency_text(value: VBACurrency) -> str:
    """A Currency as Excel reads it: $4.00 with a space after it, ($4.00) below nothing, to the nearest cent, half away from
    nothing."""
    cents = abs(Decimal(value)).quantize(Decimal('0.01'), ROUND_HALF_UP)
    return f'(${cents:,.2f})' if value < 0 else f'${cents:,.2f} '


def number_as_text(value: float) -> str:
    """A Double given where text is wanted: fifteen digits, written out while that takes no more than twenty characters."""
    if value == 0:
        return '0'
    digits, _, exponent = f'{abs(value):.14e}'.partition('e')
    power = int(exponent)
    significant = digits.replace('.', '').rstrip('0') or '0'
    if power >= 0:
        whole = significant.ljust(power + 1, '0')
        written = whole[:power + 1] + ('.' + whole[power + 1:] if len(whole) > power + 1 else '')
    else:
        written = '0.' + '0' * (-power - 1) + significant
    if len(written) > 20:
        mantissa = significant[0] + ('.' + significant[1:] if len(significant) > 1 else '')
        written = f'{mantissa}E{"+" if power >= 0 else "-"}{abs(power):02d}'
    return f'-{written}' if value < 0 else written


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float, VBAInt, VBACurrency, VBADate, VBASingle, Decimal)) and not isinstance(value, bool)


def _block_cells(sheet: Any, area: Any) -> list[tuple[int, int]]:
    """The cells of a block in the order they are tried. Blank cells past what a large block's sheet uses all do the same, so
    one of them, the block's last, stands for them."""
    rows, columns = range(area.top, area.bottom + 1), range(area.left, area.right + 1)
    if len(rows) * len(columns) <= LARGE_BLOCK:
        return [(row, column) for row in rows for column in columns]
    used = sheet.used_bounds()
    found: list[tuple[int, int]] = []
    if used is not None:
        top, left, bottom, right = used
        found = [(row, column) for row in range(max(top, area.top), min(bottom, area.bottom) + 1)
                 for column in range(max(left, area.left), min(right, area.right) + 1)]
    last = (area.bottom, area.right)
    return found if last in found else [*found, last]


def _pattern(value: object) -> object:
    """A pattern converts as VBA does: separate areas read as the first, a block is an array, and a number rounds to even."""
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        first = value.areas[0]
        if first.rows * first.columns > 1:
            raise error(13)
        value = Range(value.sheet, [first]).vba_value()
    if not _is_number(value):
        return value
    whole = round(float(to_number(value)))
    return CODES['pattern'].get(whole, whole)


def given(target: VBAObject, name: str, value: object) -> object:
    """``value`` as the property ``name`` of ``target`` is to be given it."""
    from pyopenvba.apps.excel._model import Range

    lower = name.lower()
    if target.vba_type_name not in MEASURED:
        return value
    if lower == 'value' and target.vba_type_name in ('Border', 'Borders'):
        # A border's Value is its LineStyle, but for the negative numbers on Borders (tests/fixtures/format_properties.json).
        if target.vba_type_name == 'Borders' and _is_number(value) and float(to_number(value)) <= -1:
            raise error(1004, 'Unable to set the Value property of the Borders class')
        lower = 'linestyle'
    if lower == 'pattern':
        return _pattern(value)
    if isinstance(value, Range) and value.areas[0].rows * value.areas[0].columns > 1:
        # A first area of several cells is nothing to hold, and no shade.
        if lower in HOLDING:
            return EMPTY
        if lower == 'tintandshade':
            raise error(5)
    if lower == 'tintandshade' and target.vba_type_name == 'Font' and _is_number(value):
        amount = float(to_number(value))
        return math.trunc(amount * 32767) / 32767 if -1 <= amount <= 1 else value
    if lower in LONG_LIMITED and target.vba_type_name in ('Font', 'Border', 'Borders'):
        # The largest Long stands for no value, as a colour, once it is rounded; a number past it overflows.
        number = value
        if isinstance(value, Range) and value.areas[0].rows * value.areas[0].columns == 1:
            number = Range(value.sheet, [value.areas[0]]).vba_value()
        if _is_number(number) and round(float(to_number(number))) == 2147483647:
            raise error(1004, f'Unable to set the {name} property of the {target.vba_type_name} class')
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
            if lower == 'style' and value is EMPTY:
                return 'Normal'
        elif lower in CUTTING:
            area = value.first
            value = Range(value.sheet, [Area(area.top, area.left, area.top, area.left, value.sheet.name)]).vba_value()
        else:
            area = value.first
            cells = [Range(value.sheet, [Area(row, column, row, column, value.sheet.name)]) for row, column in _block_cells(value.sheet, area)]
            spec = target.vba_member(name)
            failed: VBARuntimeError | None = None
            last: object = EMPTY
            for cell in cells:
                if spec is None or spec.setter is None:
                    break
                try:
                    last = given(target, name, cell.vba_value())
                    spec.setter(target, last)
                except VBARuntimeError as failure:
                    failed = failure
            if failed is not None:
                raise failed
            # The assignment that called this sets it once more, to the same.
            return last
    if lower == 'size' and target.vba_type_name == 'Font' and _is_number(value):
        twentieths = Decimal(repr(float(to_number(value)))) * 20
        return float(twentieths.quantize(Decimal(1), ROUND_HALF_UP) / 20) if twentieths >= Decimal('19.5') else value
    if lower in CODES and _is_number(value):
        # A date's number, as it stands in a cell. A line's rounds down, so -0.5 is no line; the rest are cut.
        number = float(to_number(value))
        whole = math.floor(number) if lower == 'linestyle' else math.trunc(number)
        whole = CODES[lower].get(whole, whole)
        if lower in ACCEPTED and whole not in ACCEPTED[lower]:
            raise refusal
        return whole
    if lower in AS_TEXT and isinstance(value, float) and not isinstance(value, bool):
        return number_as_text(value)
    if lower == 'indentlevel' and _is_number(value):
        return math.floor(float(to_number(value)))
    if lower == 'orientation' and _is_number(value):
        return math.floor(float(to_number(value)))
    if lower == 'background' and _is_number(value):
        whole = math.trunc(float(to_number(value)))
        if whole not in BACKGROUNDS:
            raise refusal
        return whole
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
        serial = float(to_number(value))
        if 1 <= serial < 61:
            # Before March 1900 its text is of the day the sheet's serials call it.
            day = int(serial)
            return f'2/29/1900' if day == 60 else f'{"1" if day <= 31 else "2"}/{day if day <= 31 else day - 31}/1900'
    return value
