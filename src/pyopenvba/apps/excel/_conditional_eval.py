"""Evaluate conditional rules without changing cells or their formats."""
from __future__ import annotations

from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _styles as styles
from pyopenvba.apps.excel._conditional import NS, Rule, formula_anchor
from pyopenvba.apps.excel._differential import read_color
from pyopenvba.apps.excel._engine_book import scalar
from pyopenvba.apps.excel._xstring import decode
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.values import ExcelError, Scalar
from pyopenvba.formula._parse import FormulaError, shift_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Worksheet

FORMULA_PREDICATES = frozenset(('containsText', 'notContainsText', 'beginsWith', 'endsWith',
                              'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors',
                              'timePeriod'))


def _matches(sheet: Worksheet, rule: Rule, row: int, column: int) -> bool:
    calculator = sheet.book.calculator
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, row, column, today=now.date(), now=now)
    kind = rule.root.get('type', '')
    anchor = formula_anchor(rule.areas) if kind in FORMULA_PREDICATES or kind in ('expression', 'cellIs') else rule.areas[0]

    def evaluated(text: str) -> Scalar:
        formula = shift_text('=' + decode(text).removeprefix('='), row - anchor.top, column - anchor.left)
        from pyopenvba.apps.excel._udf import evaluation_origin

        with evaluation_origin(sheet.book.application, "conditional"):
            return context.first(context.formula(calculator.engine_book.read(formula)))

    formulas = [entry.text or '' for entry in rule.root.findall(NS + 'formula')]
    if kind in ('top10', 'aboveAverage', 'uniqueValues', 'duplicateValues'):
        from pyopenvba.apps.excel._conditional_stats import matches

        return matches(sheet, rule, row, column)
    if kind not in FORMULA_PREDICATES and kind not in ('expression', 'cellIs'):
        raise VBAUnsupportedError(f'Display evaluation of conditional rule {kind!r} is not implemented')
    try:
        if not formulas:
            return False
        # Native notContainsBlanks paints error cells even though its stored
        # LEN(TRIM(...))>0 compatibility formula would itself return an error.
        if kind == 'notContainsBlanks' and isinstance(scalar(calculator.value_of(sheet.name, row, column)), CellError):
            return True
        first = evaluated(formulas[0])
        if kind != 'cellIs':
            return context.logical(first)
        value = scalar(calculator.value_of(sheet.name, row, column))
        operator = rule.root.get('operator', '')
        comparisons = {'equal': '=', 'notEqual': '<>', 'greaterThan': '>', 'lessThan': '<',
                       'greaterThanOrEqual': '>=', 'lessThanOrEqual': '<='}
        if operator in comparisons:
            return context.logical(context.first(context.operate(comparisons[operator], value, first)))
        if operator not in ('between', 'notBetween'):
            raise VBAUnsupportedError(f'Conditional comparison {operator!r} is not implemented')
        if len(formulas) < 2:
            return False
        second = evaluated(formulas[1])
        # Native cellIs treats between bounds as an unordered pair.
        if context.logical(context.first(context.operate('>', first, second))):
            first, second = second, first
        lower = context.logical(context.first(context.operate('>=', value, first)))
        upper = context.logical(context.first(context.operate('<=', value, second)))
        return lower and upper if operator == 'between' else not (lower and upper)
    except (ExcelError, FormulaSyntaxError, FormulaError):
        return False


def _overlay(style: styles.Style, root: XmlElement) -> styles.Style:
    """Overlay only explicitly present DXF properties; absence always inherits."""
    font = root.find(NS + 'font')
    if font is not None:
        current = style.font
        for tag, name in (('b', 'bold'), ('i', 'italic'), ('strike', 'strike'), ('outline', 'outline'),
                          ('shadow', 'shadow'), ('condense', 'condense'), ('extend', 'extend')):
            item = font.find(NS + tag)
            if item is not None:
                current = replace(current, **{name: item.get('val', '1') in ('1', 'true')})
        for tag, name in (('name', 'name'), ('family', 'family'), ('charset', 'charset'), ('scheme', 'scheme')):
            item = font.find(NS + tag)
            if item is not None:
                current = replace(current, **{name: decode(item.get('val', ''))})
        size = font.find(NS + 'sz')
        if size is not None:
            current = replace(current, size=float(size.get('val', '11')))
        underline = font.find(NS + 'u')
        if underline is not None:
            value = underline.get('val', 'single')
            current = replace(current, underline='' if value == 'none' else value)
        vertical = font.find(NS + 'vertAlign')
        if vertical is not None:
            value = vertical.get('val', 'baseline')
            current = replace(current, vert_align='' if value == 'baseline' else value)
        color = read_color(font.find(NS + 'color'))
        if color is not None:
            current = replace(current, color=color)
        style = replace(style, font=current)
    fill = root.find(NS + 'fill')
    if fill is not None:
        if fill.find(NS + 'gradientFill') is not None:
            raise VBAUnsupportedError('Conditional gradient fill evaluation is not implemented')
        pattern = fill.find(NS + 'patternFill')
        if pattern is not None:
            current_fill = style.fill
            background = read_color(pattern.find(NS + 'bgColor'))
            foreground = read_color(pattern.find(NS + 'fgColor'))
            kind = pattern.get('patternType')
            if background is not None:
                current_fill = replace(current_fill, foreground=background, pattern='solid', gradient='', bare=False)
            if kind is not None:
                current_fill = replace(current_fill, pattern=kind, gradient='', bare=False)
            if foreground is not None:
                current_fill = replace(current_fill, foreground=foreground)
            if background is not None and current_fill.pattern not in ('none', 'solid'):
                current_fill = replace(current_fill, background=background)
            style = replace(style, fill=current_fill)
    border = root.find(NS + 'border')
    if border is not None:
        current_border = style.border
        for tag, name in (('left', 'left'), ('right', 'right'), ('top', 'top'), ('bottom', 'bottom'),
                          ('start', 'left'), ('end', 'right'), ('diagonal', 'diagonal')):
            entry = border.find(NS + tag)
            if entry is None:
                continue
            side = getattr(current_border, name)
            # A present empty side explicitly removes the line.
            color = read_color(entry.find(NS + 'color'))
            side = replace(side, style=entry.get('style', 'thin' if color is not None else ''))
            if color is not None:
                side = replace(side, color=color)
            current_border = replace(current_border, **{name: side})
        for attribute, name in (('diagonalUp', 'diagonal_up'), ('diagonalDown', 'diagonal_down')):
            if attribute in border.attrib:
                current_border = replace(current_border, **{name: border.get(attribute) in ('1', 'true')})
        if border.find(NS + 'vertical') is not None or border.find(NS + 'horizontal') is not None:
            raise VBAUnsupportedError('Conditional interior border evaluation is not implemented')
        style = replace(style, border=current_border)
    number = root.find(NS + 'numFmt')
    if number is not None:
        style = replace(style, number_format=decode(number.get('formatCode', 'General')))
    for tag in ('alignment', 'protection'):
        if root.find(NS + tag) is not None:
            raise VBAUnsupportedError(f'Conditional {tag} evaluation is not implemented')
    return style


def display_style(sheet: Worksheet, row: int, column: int) -> styles.Style:
    """The base style plus matching rules, with highest priority winning conflicts."""
    patches: list[XmlElement] = []
    for rule in sorted(sheet.conditional_formats, key=lambda entry: int(entry.root.get('priority', '0'))):
        if rule.deleted or not any(area.top <= row <= area.bottom and area.left <= column <= area.right
                                   for area in rule.areas):
            continue
        if rule.root.get('type') in ('dataBar', 'iconSet'):
            # A bar or an icon is drawn, and formats nothing DisplayFormat shows (see _data_bars, _icon_sets).
            continue
        if rule.root.get('type') == 'colorScale':
            from pyopenvba.apps.excel._colour_scales import scale_patch

            # A colour scale is a fill at its priority, for numbers alone (see _colour_scales).
            shown = scale_patch(sheet, rule, row, column)
            if shown is not None:
                patches.append(shown)
            continue
        if not _matches(sheet, rule, row, column):
            continue
        patch = rule.dxf
        identifier = rule.root.get('dxfId')
        if patch is None and identifier is not None:
            patch = sheet.book.stylesheet.dxf(int(identifier))
        if patch is not None:
            patches.append(patch)
        if rule.root.get('stopIfTrue', '0') in ('1', 'true'):
            break
    result = sheet.style_at(row, column)
    for patch in reversed(patches):
        result = _overlay(result, patch)
    return result


def value_hidden(sheet: Worksheet, row: int, column: int, value: object) -> bool:
    """Whether a data bar with ShowValue False or an icon set with ShowIconOnly True covers a cell holding the
    number ``value``, which then shows no text; text, errors and logicals show. A rule above it that matches and
    stops the rules after it leaves the value shown (tests/fixtures/data_bars_icons/)."""
    if not isinstance(value, float):
        return False
    covering = sorted((rule for rule in sheet.conditional_formats if not rule.deleted and any(
        area.top <= row <= area.bottom and area.left <= column <= area.right for area in rule.areas)),
        key=lambda entry: int(entry.root.get('priority', '0')))
    if not any(_hides(rule) for rule in covering):
        return False
    for rule in covering:
        if _hides(rule):
            return True
        if rule.root.get('type') not in ('dataBar', 'iconSet', 'colorScale') \
                and rule.root.get('stopIfTrue', '0') in ('1', 'true') and _matches(sheet, rule, row, column):
            return False
    return False


def _hides(rule: Rule) -> bool:
    """Whether a rule is a data bar or an icon set that hides the value it draws beside."""
    kind = rule.root.get('type')
    if kind not in ('dataBar', 'iconSet'):
        return False
    if rule.state is not None:
        return not rule.state.show_value
    if rule.extension_only:
        shape = rule.extension[0] if rule.extension is not None and len(rule.extension) else None
    else:
        shape = rule.root.find(NS + kind)
    return shape is not None and shape.get('showValue', '1') in ('0', 'false')


def conditional_number_format(sheet: Worksheet, row: int, column: int) -> str:
    """Resolve only number formats, without evaluating irrelevant visual rules."""
    base = sheet.style_at(row, column).number_format
    candidates: list[tuple[Rule, XmlElement | None]] = []
    for rule in sorted(sheet.conditional_formats, key=lambda entry: int(entry.root.get('priority', '0'))):
        if rule.deleted or not any(area.top <= row <= area.bottom and area.left <= column <= area.right
                                   for area in rule.areas):
            continue
        patch = rule.dxf
        identifier = rule.root.get('dxfId')
        if patch is None and identifier is not None:
            patch = sheet.book.stylesheet.dxf(int(identifier))
        candidates.append((rule, None if patch is None else patch.find(NS + 'numFmt')))
    if not any(number is not None for _, number in candidates):
        return base
    for rule, number in candidates:
        stopping = rule.root.get('stopIfTrue', '0') in ('1', 'true')
        if number is None and not stopping:
            continue
        if not _matches(sheet, rule, row, column):
            continue
        if number is not None:
            return decode(number.get('formatCode', 'General'))
        if stopping:
            return base
    return base
