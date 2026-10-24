"""Worksheet conditional rules and Excel's ordinary FormatConditions API.

The file model follows William's pyOfficeEditor: rules retain their original
XML, including extensions, and differential styles live in the workbook dxfs.
Runtime priorities and range deletion follow independent native Excel controls.
This module describes rules; it does not claim to render DisplayFormat.
"""
# ruff: noqa: N803 -- Excel's public named arguments use these exact names.
from __future__ import annotations

import re
from collections.abc import Sequence
from copy import deepcopy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_parser, xml_tree
from pyopenvba.apps.excel._conditional_predicates import (
    PREDICATE_TYPES,
    TEXT_OPERATORS,
    TEXT_TYPES,
    TIME_PERIODS,
    formula_for_rule,
    regenerate,
    saved_text,
)
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._conditional_extension import RuleState
    from pyopenvba.apps.excel._model import Worksheet

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
OPERATORS = ('between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan',
             'greaterThanOrEqual', 'lessThanOrEqual')
#: CT_CfRule's attributes in the schema's order, which Excel writes them in.
_RULE_ORDER = {name: index for index, name in enumerate((
    'type', 'dxfId', 'priority', 'stopIfTrue', 'aboveAverage', 'percent', 'bottom', 'operator', 'text',
    'timePeriod', 'rank', 'stdDev', 'equalAverage'))}
_AFTER = frozenset(('dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
                    'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches',
                    'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF',
                    'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'))


@dataclass(eq=False)
class Rule:
    """One original cfRule, its cells, and its lazily loaded differential style."""

    areas: list[Area]
    root: XmlElement
    container: XmlElement | None = field(default=None, repr=False)
    dxf: XmlElement | None = field(default=None, repr=False)
    deleted: bool = False
    imported: bool = False
    color_slot: int | None = None
    loaded: bool = False
    plain_load: bool | None = None
    child_order: int | None = None
    #: The rule's x14:cfRule in the worksheet's extension, as a file kept it (_conditional_extension).
    extension: XmlElement | None = field(default=None, repr=False)
    #: True for a rule the extension holds alone, with no cfRule in a conditionalFormatting element.
    extension_only: bool = False
    #: The id pairing the rule's cfRule with its x14:cfRule; '' until it has one.
    identifier: str = ''
    #: Where the file kept the x14:cfRule: its x14:conditionalFormatting's place and range, and its own place.
    extension_place: tuple[int, str, int] | None = None
    #: A data bar's or an icon set's settings, once read or made (_data_bars, _icon_sets).
    state: RuleState | None = field(default=None, repr=False)
    #: A colour scale's points VBA gave a value, by index: True for text, False for a number (_colour_scales).
    given: dict[int, bool] = field(default_factory=lambda: {}, repr=False)
    #: Whether a colour scale has coloured a cell since it was made or read, which its lowest and highest points
    #: need before they read other than 0 (_colour_scales).
    drawn: bool = False


def sqref(areas: list[Area]) -> str:
    """A range as a file's sqref spells it: each area by its corner cells, so a whole column is B1:B1048576 and a
    whole row A2:XFD2 (tests/fixtures/conditional_edits/, full_column_rows)."""
    def cells(area: Area) -> str:
        first = f'{column_letter(area.left)}{area.top}'
        if (area.top, area.left) == (area.bottom, area.right):
            return first
        return f'{first}:{column_letter(area.right)}{area.bottom}'

    return ' '.join(cells(area) for area in areas)


def formula_anchor(areas: list[Area]) -> Area:
    """Formula-backed conditions use the bounding corner, independently of API area order."""
    top, left = min(area.top for area in areas), min(area.left for area in areas)
    return Area(top, left, top, left)


def column_order(area: Area) -> tuple[int, int, int, int]:
    """Native group ordering compares each column interval before its row interval."""
    return area.left, area.right, area.top, area.bottom


def canonical_load(rule: Rule) -> bool:
    """The package's captured plain-load mode is independent of source FileFormat."""
    return rule.imported if rule.plain_load is None else rule.plain_load


#: Rules drawn from their whole range, whose areas a file's order keeps (tests/fixtures/conditional_edits/, order_*).
RANGE_RULES = frozenset(('top10', 'aboveAverage', 'uniqueValues', 'duplicateValues', 'colorScale', 'dataBar',
                         'iconSet'))


def projected_areas(rule: Rule) -> list[Area]:
    """Plain-loaded rules a formula decides expose row order, by each area's top row, then its bottom, then its left
    column (tests/fixtures/conditional_edits/, order_wide_first); rules drawn from their range retain input order."""
    if rule.imported and rule.root.get('type') not in RANGE_RULES:
        return sorted(rule.areas, key=lambda area: (area.top, area.bottom, area.left, area.right))
    return list(rule.areas)


def group_order(rule: Rule) -> tuple[tuple[int, int, int, int], ...]:
    """Plain-loaded groups compare the complete ordered family-projected population."""
    end = (MAX_COLUMNS + 1, MAX_COLUMNS + 1, MAX_ROWS + 1, MAX_ROWS + 1)
    return (*tuple(column_order(area) for area in projected_areas(rule)), end)


def _plain_container(rule: Rule) -> bool:
    container = rule.container
    return container is not None and container.tag == NS + 'conditionalFormatting' and all(
        key == 'sqref' or key == 'xmlns' or key.startswith('xmlns:') for key in container.attrib
    ) and all(child.tag == NS + 'cfRule' for child in container)


def rule_groups(rules: list[Rule]) -> list[list[Rule]]:
    """Stage native groups without changing stored populations, rules or metadata."""
    groups: list[list[Rule]] = []
    plain: dict[XmlElement, bool] = {}
    metadata_cache: dict[XmlElement, tuple[tuple[str, str], ...]] = {}

    def canonical_container(rule: Rule) -> bool:
        container = rule.container
        if not canonical_load(rule) or container is None:
            return False
        if container not in plain:
            plain[container] = _plain_container(rule)
        return plain[container]

    lookup: dict[tuple[object, ...], list[Rule]] = {}
    for rule in rules:
        if rule.deleted or not rule.areas:
            continue
        population = group_order(rule)
        if canonical_container(rule):
            assert rule.container is not None
            if rule.container not in metadata_cache:
                metadata_cache[rule.container] = tuple(sorted(
                    (key, value) for key, value in rule.container.attrib.items() if key != 'sqref'))
            metadata = metadata_cache[rule.container]
            key: tuple[object, ...] = (('captured', rule.color_slot, population, metadata)
                                      if rule.child_order is not None else ('canonical', population, metadata))
        else:
            key = ('raw', id(rule.container), population, rule.imported,
                   rule.color_slot if rule.child_order is not None else None)
            if rule.container is None and not rule.imported:
                # A rule made on a range a plain loaded group covers joins that group at its end, as Excel adds
                # it there (tests/fixtures/data_bars_icons/data_bars_icons_resaved.xlsx).
                key = next((held for held in lookup if (held[0] == 'canonical' and held[1] == population)
                            or (held[0] == 'captured' and held[2] == population)), key)
        group = lookup.get(key)
        if group is None:
            new_group: list[Rule] = []
            lookup[key] = new_group
            groups.append(new_group)
            group = new_group
        group.append(rule)
    canonical: list[list[Rule]] = []
    for group in groups:
        captured_children = all(rule.child_order is not None for rule in group)
        if captured_children or canonical_container(group[0]):
            try:
                # The file's rules in their places, then any made since in priority order.
                ordered = sorted(group, key=lambda rule: (0, rule.child_order) if rule.child_order is not None
                                 else (1, int(rule.root.get('priority', '0'))))
            except ValueError:
                ordered = list(group)
            group[:] = ordered
        if canonical_container(group[0]):
            canonical.append(group)
    canonical.sort(key=lambda group: group_order(group[0]))
    ordered_groups = iter(canonical)
    groups = [next(ordered_groups) if canonical_container(group[0]) else group for group in groups]
    captured = sorted((group for group in groups if group[0].child_order is not None),
                      key=lambda group: group[0].color_slot or 0)
    captured_groups = iter(captured)
    return [next(captured_groups) if group[0].child_order is not None else group for group in groups]


def read_rules(xml: str, *, imported: bool = True, plain_load: bool | None = None) -> list[Rule]:
    return read_sheet_rules(xml, imported=imported, plain_load=plain_load)[0]


def read_sheet_rules(xml: str, *, imported: bool = True, plain_load: bool | None = None,
                     sheet_index: int = 0) -> tuple[list[Rule], bool]:
    """A worksheet's rules, with those of its Excel 2010 extension; and False where that extension holds rules
    this model does not read, which then stay out (_conditional_extension). Each rule has the id Excel gives it,
    its x14:cfRule's or one made of ``sheet_index``, the sheet's place, and the rule's priority."""
    from pyopenvba.apps.excel._conditional_extension import loaded_identifier
    from pyopenvba.apps.excel._conditional_extension import read as read_extension

    root = xml_tree(xml)
    rules: list[Rule] = []
    for group in root.findall(NS + 'conditionalFormatting'):
        # A detached rule may use ancestor-only prefixes in MC/QName values.
        for key, value in root.attrib.items():
            if key == 'xmlns' or key.startswith('xmlns:'):
                group.attrib.setdefault(key, value)
        areas = [parse_area(value) for value in group.get('sqref', '').split()]
        if areas:
            rules.extend(Rule(list(areas), entry, group, imported=imported if plain_load is None else plain_load,
                              loaded=True, plain_load=plain_load) for entry in group.findall(NS + 'cfRule'))
    for rule in rules:
        try:
            rule.identifier = loaded_identifier(sheet_index, _priority(rule))
        except (ValueError, OverflowError):
            rule.identifier = ''
    understood = read_extension(root, rules, imported=imported if plain_load is None else plain_load,
                                plain_load=plain_load)
    return rules, understood


def with_rules(xml: str, rules: list[Rule], *, extension: bool = True) -> str:
    """Replace standard rules while retaining container and extension metadata; with extension, the worksheet's
    Excel 2010 extension of conditional formats too (_conditional_extension), else that stays as it is."""
    from pyopenvba.apps.excel._conditional_extension import RULES_URI, grouped, markup
    from pyopenvba.apps.excel._conditional_extension import parts as rule_parts

    groups: list[XmlElement] = []
    plain_containers: dict[XmlElement, bool] = {}
    extension_rules: list[tuple[Rule, XmlElement]] = []
    for entries in rule_groups(rules):
        first = entries[0]
        serialized_areas = projected_areas(first)
        placed = [(rule, *rule_parts(rule)) for rule in entries]
        extension_rules.extend((rule, part) for rule, _, part in placed if part is not None)
        main = [part for _, part, _ in placed if part is not None]
        if not main:
            continue
        container = first.container
        if container is None:
            group = XmlElement(NS + 'conditionalFormatting')
        else:
            if container not in plain_containers:
                plain_containers[container] = _plain_container(first)
            if plain_containers[container]:
                group = XmlElement(container.tag, dict(container.attrib))
                group.text, group.tail = container.text, container.tail
            else:
                group = deepcopy(container)
                for entry in group.findall(NS + 'cfRule'):
                    group.remove(entry)
        group.set('sqref', sqref(serialized_areas))
        # cfRule precedes extLst in CT_ConditionalFormatting.
        before = next((i for i, entry in enumerate(group) if entry.tag == NS + 'extLst'), len(group))
        for copied in main:
            # In the schema's order, as Excel writes a rule however it was made.
            ordered = sorted(copied.attrib.items(), key=lambda item: _RULE_ORDER.get(item[0], len(_RULE_ORDER)))
            copied.attrib.clear()
            copied.attrib.update(ordered)
            if copied.get('type') in TEXT_TYPES and copied.get('text') is not None:
                if any('\x00' in decode(entry.text or '') for entry in copied.findall(NS + 'formula')):
                    raise VBAUnsupportedError('Saving conditional formulas containing NUL is not supported by native Excel')
                copied.set('text', saved_text(copied.get('text', '')))
            group.insert(before, copied)
            before += 1
        groups.append(group)
    data = xml.encode('utf-8')
    parser = xml_parser()
    depth, active = 0, None
    root_start = 0
    spans: list[tuple[int, int]] = []
    insertion: int | None = None
    #: The worksheet's own extLst and its ext of x14 rules, each [start, stop], and how many exts the extLst holds.
    ext_list: list[int] = []
    rules_ext: list[int] = []
    exts = 0
    rules_start: int | None = None
    root_close = 0

    def tag_end(start: int) -> int:
        quote = 0
        for index in range(start, len(data)):
            char = data[index]
            if quote:
                if char == quote:
                    quote = 0
            elif char in (34, 39):
                quote = char
            elif char == 62:
                return index + 1
        raise ValueError('Unterminated Office XML tag')

    def element_stop(begun: int) -> int:
        first_end = tag_end(begun)
        return first_end if data[begun:first_end].rstrip().endswith(b'/>') else tag_end(parser.CurrentByteIndex)

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal depth, active, insertion, root_start, exts, rules_start
        depth += 1
        if depth == 1:
            root_start = parser.CurrentByteIndex
        if depth == 2:
            if name == NS[1:] + 'conditionalFormatting':
                active = parser.CurrentByteIndex
            elif name.startswith(NS[1:]) and name[len(NS) - 1:] in _AFTER and insertion is None:
                insertion = parser.CurrentByteIndex
            if name == NS[1:] + 'extLst':
                ext_list[:] = [parser.CurrentByteIndex]
        if depth == 3 and len(ext_list) == 1 and name == NS[1:] + 'ext':
            exts += 1
            if attrs.get('uri') == RULES_URI:
                rules_start = parser.CurrentByteIndex

    def end(name: str) -> None:
        nonlocal depth, active, insertion, rules_start, root_close
        if depth == 2 and active is not None:
            spans.append((active, element_stop(active)))
            active = None
        if depth == 3 and rules_start is not None:
            rules_ext[:] = [rules_start, element_stop(rules_start)]
            rules_start = None
        if depth == 2 and len(ext_list) == 1 and name == NS[1:] + 'extLst':
            ext_list.append(element_stop(ext_list[0]))
        if depth == 1:
            root_close = parser.CurrentByteIndex
            if insertion is None:
                insertion = parser.CurrentByteIndex
        depth -= 1

    parser.StartElementHandler, parser.EndElementHandler = start, end
    parser.Parse(data, True)
    assert insertion is not None
    root_end = tag_end(root_start)
    declared = {prefix.decode(): uri.decode() for prefix, uri in
                re.findall(rb'\sxmlns(?::([\w.-]+))?\s*=\s*"([^"]*)"', data[root_start:root_end])}
    groups_markup = b''.join(_group_markup(group, declared) for group in groups)
    order = {id(rule): index for index, rule in enumerate(rules)}
    extension_rules.sort(key=lambda placed: order[id(placed[0])])
    extended = markup(grouped(extension_rules)).encode('utf-8') if extension else b''
    if data[root_start:root_end].endswith(b'/>'):
        tail = b'<extLst>' + extended + b'</extLst>' if extended else b''
        if not groups_markup and not tail:
            return xml
        tag = data[root_start + 1:root_end].split(None, 1)[0].rstrip(b'/>')
        return (data[:root_end - 2] + b'>' + groups_markup + tail + b'</' + tag + b'>' + data[root_end:]).decode('utf-8')
    edits = [(first, last, b'') for first, last in spans]
    tail = b''
    if extension and rules_ext:
        if not extended and exts == 1 and len(ext_list) == 2:
            # The extLst held nothing else.
            edits.append((ext_list[0], ext_list[1], b''))
        else:
            edits.append((rules_ext[0], rules_ext[1], extended))
    elif extended and len(ext_list) == 2 and not data[ext_list[0]:ext_list[1]].rstrip().endswith(b'/>'):
        # A new ext of x14 rules goes first in an extLst, where Excel writes it.
        opening = tag_end(ext_list[0])
        edits.append((opening, opening, extended))
    elif extended and len(ext_list) == 2:
        edits.append((ext_list[0], ext_list[1], b'<extLst>' + extended + b'</extLst>'))
    elif extended:
        tail = b'<extLst>' + extended + b'</extLst>'
    if tail and insertion == root_close:
        edits.append((insertion, insertion, groups_markup + tail))
    else:
        edits.append((insertion, insertion, groups_markup))
        if tail:
            edits.append((root_close, root_close, tail))
    for first, last, replacement in sorted(edits, key=lambda edit: (edit[0], edit[1]), reverse=True):
        data = data[:first] + replacement + data[last:]
    return data.decode('utf-8')


def _group_markup(group: XmlElement, declared: dict[str, str]) -> bytes:
    """A conditionalFormatting element as Excel writes it: in the worksheet's own prefixes, declaring none of the
    namespaces the worksheet declares, and with ``/>`` closing an empty element (tests/fixtures/colour_scales/)."""
    for prefix, uri in declared.items():
        group.attrib.setdefault('xmlns:' + prefix if prefix else 'xmlns', uri)
    text = xml_bytes(group).split(b'?>', 1)[1].lstrip().replace(b' />', b'/>')
    # A cfRule's link to its x14:cfRule names its uri before declaring x14 (tests/fixtures/data_bars_icons/).
    text = re.sub(rb'<ext (xmlns:x14="[^"]*") (uri="[^"]*")', rb'<ext \2 \1', text)
    head = text.index(b'>')
    opening = text[:head]
    for prefix, uri in declared.items():
        name = ('xmlns:' + prefix if prefix else 'xmlns').encode()
        opening = opening.replace(b' ' + name + b'="' + uri.encode() + b'"', b'', 1)
    return opening + text[head:]


def _priority(rule: Rule) -> int:
    return int(rule.root.get('priority', '0'))


def changed(sheet: Worksheet) -> None:
    sheet.conditional_color_groups = None
    sheet.conditional_changed = True
    sheet.touched()


def guard(sheet: Worksheet) -> None:
    from pyopenvba.apps.excel._protection import APPLICATION_DEFINED

    protection = sheet.protection
    # Conditional formatting follows AllowFormattingCells even for UI-only
    # protection, where many ordinary macro writes would otherwise be allowed.
    if protection is not None and protection.contents and 'AllowFormattingCells' not in protection.allows:
        raise error(1004, APPLICATION_DEFINED)


def meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def without(area: Area, cut: Area) -> list[Area]:
    """Native CF split order: top, bottom, left, right (unlike validation)."""
    if not meets(area, cut):
        return [area]
    pieces: list[Area] = []
    if area.top < cut.top:
        pieces.append(Area(area.top, area.left, cut.top - 1, area.right))
    if cut.bottom < area.bottom:
        pieces.append(Area(cut.bottom + 1, area.left, area.bottom, area.right))
    top, bottom = max(area.top, cut.top), min(area.bottom, cut.bottom)
    if area.left < cut.left:
        pieces.append(Area(top, area.left, bottom, cut.left - 1))
    if cut.right < area.right:
        pieces.append(Area(top, cut.right + 1, bottom, area.right))
    return pieces


def remove(sheet: Worksheet, areas: list[Area]) -> bool:
    """Subtract cells from every rule; caller owns its operation's protection."""
    planned: list[tuple[Rule, list[Area], list[str]]] = []
    for rule in list(sheet.conditional_formats):
        remaining = list(rule.areas)
        for cut in areas:
            kept: list[Area] = []
            extra: list[Area] = []
            for area in remaining:
                pieces = without(area, cut)
                kept.extend(pieces[:1])
                extra.extend(pieces[1:])
            remaining = kept + extra
        if remaining != rule.areas:
            formulas = [child.text or '' for child in rule.root.findall(NS + 'formula')]
            if remaining:
                anchored = rule.root.get('type') in {'cellIs', 'expression'} | PREDICATE_TYPES
                original = formula_anchor(rule.areas) if anchored else rule.areas[0]
                first = formula_anchor(remaining) if anchored else remaining[0]
                if (original.top, original.left) != (first.top, first.left):
                    generated = formula_for_rule(rule.root, first)
                    if generated is not None:
                        formulas = [encode_text(generated)]
                        planned.append((rule, remaining, formulas))
                        continue
                    from pyopenvba.formula._parse import FormulaError, shift_text

                    try:
                        formulas = [encode_text(shift_text('=' + decode(text), first.top - original.top,
                                                           first.left - original.left)[1:]) if text else text
                                    for text in formulas]
                    except FormulaError:
                        raise VBAUnsupportedError('Rebasing this conditional formatting formula is not implemented') from None
            planned.append((rule, remaining, formulas))
    for rule, remaining, formulas in planned:
        rule.areas = remaining
        if remaining and rule.root.get('type') in PREDICATE_TYPES:
            regenerate(rule.root, formula_anchor(remaining))
        else:
            for child, text in zip(rule.root.findall(NS + 'formula'), formulas, strict=True):
                child.text = text
        if not remaining:
            rule.deleted = True
            sheet.conditional_formats.remove(rule)
    if planned:
        changed(sheet)
    return bool(planned)


def _formula(value: object, target: Range) -> str:
    from pyopenvba.apps.excel._model import spelled_formula

    if value is MISSING:
        raise error(449)
    if value is NULL:
        return ''
    text = to_text(value)
    if not text:
        raise error(5)
    # Formula arguments read back with '=' even for numeric constants.
    try:
        return spelled_formula(target.sheet, text if text.startswith('=') else '=' + text,
                               target.first.top, target.first.left)[1:]
    except VBARuntimeError:
        raise error(5) from None


def _made(target: Range, Type: object, Operator: object, Formula1: object, Formula2: object,
          String: object = MISSING, TextOperator: object = MISSING, DateOperator: object = MISSING) -> XmlElement:
    if Type is MISSING:
        raise error(449)
    kind = int(to_integer(Type, 'Long'))
    if kind in (9, 10, 11, 13, 16, 17):
        attributes = {'stopIfTrue': '1'}
        if kind == 9:
            if String is MISSING or TextOperator is MISSING:
                raise error(449)
            selected = _stat_long(TextOperator)
            if selected not in range(4):
                raise error(5)
            text = '' if String is NULL else to_text(String)
            units = text.encode('utf-16-le', 'surrogatepass')
            if len(units) > 510:
                raise error(5)
            if len(units) == 510:
                text = units[:508].decode('utf-16-le', 'surrogatepass')
            attributes.update(type=TEXT_TYPES[selected], operator=TEXT_OPERATORS[selected], text=text)
        elif kind == 11:
            if DateOperator is MISSING:
                raise error(449)
            selected = _stat_long(DateOperator)
            if selected not in range(len(TIME_PERIODS)):
                raise error(5)
            attributes.update(type='timePeriod', timePeriod=TIME_PERIODS[selected])
        else:
            attributes['type'] = {10: 'containsBlanks', 13: 'notContainsBlanks',
                                  16: 'containsErrors', 17: 'notContainsErrors'}[kind]
        root = XmlElement(NS + 'cfRule', attributes)
        regenerate(root, formula_anchor(target.areas))
        return root
    if kind in (5, 8, 12):
        attributes = {5: {'type': 'top10', 'rank': '10'}, 8: {'type': 'uniqueValues'},
                      12: {'type': 'aboveAverage', 'stopIfTrue': '1'}}[kind]
        return XmlElement(NS + 'cfRule', attributes)
    if kind not in (1, 2):
        if kind in (3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 16, 17):
            raise VBAUnsupportedError('This conditional formatting rule type is not implemented')
        raise error(5)
    attributes = {'type': 'cellIs' if kind == 1 else 'expression', 'stopIfTrue': '1'}
    count = 1
    if kind == 1:
        if Operator is MISSING:
            raise error(449)
        if Operator is NULL:
            raise error(5)
        operator = int(to_integer(Operator, 'Long'))
        if not 1 <= operator <= len(OPERATORS):
            raise error(5)
        attributes['operator'] = OPERATORS[operator - 1]
        count = 2 if operator in (1, 2) else 1
    formulas = [_formula(Formula1, target)]
    if count == 2:
        formulas.append(_formula(Formula2, target))
    anchor = formula_anchor(target.areas)
    if (anchor.top, anchor.left) != (target.first.top, target.first.left):
        from pyopenvba.formula._parse import shift_text

        formulas = [shift_text('=' + value, anchor.top - target.first.top,
                               anchor.left - target.first.left)[1:] if value else value for value in formulas]
    root = XmlElement(NS + 'cfRule', attributes)
    for value in formulas:
        child = XmlElement(NS + 'formula')
        child.text = encode_text(value)
        root.append(child)
    return root


class FormatConditions(VBACollection, ExcelObject):
    vba_type_name = 'FormatConditions'

    def __init__(self, target: Range) -> None:
        self.target, self.sheet = target, target.sheet

    def vba_items(self) -> list[object]:
        rules = [rule for rule in self.sheet.conditional_formats if not rule.deleted
                 and any(meets(one, area) for one in rule.areas for area in self.target.areas)]
        return [_wrapper(self.sheet, rule) for rule in sorted(rules, key=_priority)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(1004)
        return items[position - 1]

    @member
    def Parent(self) -> object:
        return self.target

    @method
    def Add(self, Type: object = MISSING, Operator: object = MISSING, Formula1: object = MISSING,
            Formula2: object = MISSING, String: object = MISSING, TextOperator: object = MISSING,
            DateOperator: object = MISSING, ScopeType: object = MISSING) -> object:
        guard(self.sheet)
        if ScopeType is not MISSING and ScopeType is not NULL:
            raise error(5)
        root = _made(self.target, Type, Operator, Formula1, Formula2, String, TextOperator, DateOperator)
        return self._append(root)

    def _append(self, root: XmlElement, state: RuleState | None = None) -> object:
        from pyopenvba.apps.excel._conditional_border_context import register

        root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        rule = Rule([Area(area.top, area.left, area.bottom, area.right) for area in self.target.areas], root,
                    state=state)
        register(self.sheet, rule)
        self.sheet.conditional_formats.append(rule)
        changed(self.sheet)
        return _wrapper(self.sheet, rule)

    @method
    def AddTop10(self) -> object:
        guard(self.sheet)
        return self._append(_made(self.target, 5, MISSING, MISSING, MISSING))

    @method
    def AddAboveAverage(self) -> object:
        guard(self.sheet)
        return self._append(_made(self.target, 12, MISSING, MISSING, MISSING))

    @method
    def AddUniqueValues(self) -> object:
        guard(self.sheet)
        return self._append(_made(self.target, 8, MISSING, MISSING, MISSING))

    @method
    def AddColorScale(self, ColorScaleType: object = MISSING) -> object:
        from pyopenvba.apps.excel._colour_scales import made

        guard(self.sheet)
        if ColorScaleType is MISSING:
            raise error(450)
        return self._append(made(int(to_integer(ColorScaleType, 'Long'))))

    @method
    def AddDatabar(self) -> object:
        from pyopenvba.apps.excel._conditional_extension import foreign
        from pyopenvba.apps.excel._data_bars import made

        guard(self.sheet)
        foreign(self.sheet)
        return self._append(*made())

    @method
    def AddIconSetCondition(self) -> object:
        from pyopenvba.apps.excel._conditional_extension import foreign
        from pyopenvba.apps.excel._icon_sets import made

        guard(self.sheet)
        foreign(self.sheet)
        return self._append(*made())

    @method
    def Delete(self) -> object:
        guard(self.sheet)
        remove(self.sheet, self.target.areas)
        return EMPTY


class RuleObject(ExcelObject):
    """What every kind of rule answers: its range, priority and life; a rule with a differential style adds that
    (see _ConditionalRule)."""

    vba_type_name = ''

    def __init__(self, sheet: Worksheet, rule: Rule) -> None:
        self.sheet, self.rule = sheet, rule

    @member
    def Parent(self) -> object:
        return self.AppliesTo()


    def touched(self) -> None:
        """Note a change to the rule, for its sheet's file to keep."""
        changed(self.sheet)

    @member
    def AppliesTo(self) -> object:
        areas = projected_areas(self.rule)
        return Range(self.sheet, [Area(a.top, a.left, a.bottom, a.right, self.sheet.name) for a in areas])

    @member
    def Priority(self) -> object:
        return VBAInt(_priority(self.rule), 'Long')

    @setter('Priority')
    def set_priority(self, value: object) -> None:
        guard(self.sheet)
        wanted = int(to_integer(value, 'Long'))
        if wanted < 1:
            raise error(5)
        for other in self.sheet.conditional_formats:
            priority = _priority(other)
            if other is not self.rule and wanted <= priority:
                other.root.set('priority', str(priority + 1))
        self.rule.root.set('priority', str(wanted))
        changed(self.sheet)

    @member
    def StopIfTrue(self) -> object:
        return self.rule.root.get('stopIfTrue', '0') in ('1', 'true')

    @setter('StopIfTrue')
    def set_stop(self, value: object) -> None:
        guard(self.sheet)
        # Excel writes stopIfTrue="1" alone, leaving a rule that goes on unmarked (tests/fixtures/conditional_joins/).
        if to_bool(value):
            self.rule.root.set('stopIfTrue', '1')
        else:
            self.rule.root.attrib.pop('stopIfTrue', None)
        changed(self.sheet)

    @method
    def SetFirstPriority(self) -> object:
        guard(self.sheet)
        for other in self.sheet.conditional_formats:
            if other is not self.rule:
                other.root.set('priority', str(_priority(other) + 1))
        self.rule.root.set('priority', '1')
        changed(self.sheet)
        return EMPTY

    @method
    def SetLastPriority(self) -> object:
        guard(self.sheet)
        self.rule.root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        changed(self.sheet)
        return EMPTY

    @method
    def Delete(self) -> object:
        guard(self.sheet)
        if self.rule not in self.sheet.conditional_formats:
            raise error(1004)
        from pyopenvba.apps.excel._conditional_border_context import delete_member

        delete_member(self.sheet, self.rule)
        self.sheet.conditional_formats.remove(self.rule)
        self.rule.deleted = True
        changed(self.sheet)
        return EMPTY


    @method
    def ModifyAppliesToRange(self, Range: object = MISSING) -> object:
        from pyopenvba.apps.excel._model import Range as ExcelRange

        guard(self.sheet)
        if not isinstance(Range, ExcelRange) or Range.sheet is not self.sheet:
            raise error(1004)
        areas = [Area(a.top, a.left, a.bottom, a.right) for a in Range.areas]
        prepared = deepcopy(self.rule.root)
        if prepared.get('type') in ('cellIs', 'expression'):
            from pyopenvba.formula._parse import FormulaError, shift_text

            old, new = formula_anchor(self.rule.areas), formula_anchor(areas)
            try:
                for child in prepared.findall(NS + 'formula'):
                    if child.text:
                        child.text = encode_text(shift_text('=' + decode(child.text), new.top - old.top,
                                                            new.left - old.left)[1:])
            except FormulaError:
                raise VBAUnsupportedError('Rebasing this conditional formatting formula is not implemented') from None
        else:
            regenerate(prepared, formula_anchor(areas))
        from pyopenvba.apps.excel._conditional_border_context import move_member

        move_member(self.sheet, self.rule)
        self.rule.areas = areas
        self.rule.imported = False
        self.rule.root[:] = list(prepared)
        changed(self.sheet)
        return EMPTY


class PointRule(RuleObject):
    """A colour scale, a data bar or an icon set: a rule drawn from its points rather than a format for what it
    matches. StopIfTrue reads False and is error 450 to set; Formula is error 1004 until one is set, then a
    formula the cfRule holds, and setting it empty or past reading is error 1004; ScopeType is error 5
    (tests/fixtures/data_bars_icons/)."""

    vba_type_name = ''

    @member
    def StopIfTrue(self) -> object:
        return False

    @setter('StopIfTrue')
    def set_stop(self, value: object) -> None:
        raise error(450)

    @member
    def Formula(self) -> object:
        found = self.rule.root.find(NS + 'formula')
        if found is None:
            raise error(1004)
        return '=' + decode(found.text or '')

    @setter('Formula')
    def set_formula(self, value: object) -> None:
        from pyopenvba.formula._parse import FormulaError, parse

        guard(self.sheet)
        text = to_text(value).removeprefix('=')
        try:
            parse('=' + text)
        except FormulaError:
            raise error(1004) from None
        found = self.rule.root.find(NS + 'formula')
        if found is None:
            found = XmlElement(NS + 'formula')
            self.rule.root.insert(0, found)
        found.text = encode_text(text)
        self.touched()

    @member
    def ScopeType(self) -> object:
        raise error(5)


class _ConditionalRule(RuleObject):
    """A rule that formats the cells it matches with a differential style."""

    def _dxf(self) -> XmlElement:
        if self.rule.dxf is None:
            identifier = self.rule.root.get('dxfId')
            self.rule.dxf = (self.sheet.book.stylesheet.dxf(int(identifier)) if identifier is not None
                             else XmlElement(NS + 'dxf'))
        return self.rule.dxf

    def _style_changed(self) -> None:
        self.rule.root.set('dxfId', str(self.sheet.book.stylesheet.ensure_dxf(self._dxf())))
        changed(self.sheet)

    @member
    def Font(self) -> object:
        from pyopenvba.apps.excel._differential import ConditionalFont

        return ConditionalFont(self._dxf(), self._style_changed, self.sheet)

    @member
    def Interior(self) -> object:
        from pyopenvba.apps.excel._differential import ConditionalInterior

        return ConditionalInterior(self._dxf(), self._style_changed, self.sheet)

    @member
    def Borders(self, Index: object = MISSING) -> object:
        from pyopenvba.apps.excel._conditional_border_context import resolve
        from pyopenvba.apps.excel._differential import ConditionalBorders

        collection = ConditionalBorders(self._dxf(), self._style_changed, self.sheet,
                                        color_source=lambda: resolve(self.sheet, self.rule))
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def NumberFormat(self) -> object:
        from pyopenvba.apps.excel._differential import number_format

        return number_format(self._dxf())

    @setter('NumberFormat')
    def set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._differential import set_number_format

        guard(self.sheet)
        set_number_format(self._dxf(), value, self.sheet.book.stylesheet)
        self._style_changed()


class FormatCondition(_ConditionalRule):
    vba_type_name = 'FormatCondition'

    def vba_get(self, name: str, args: Sequence[object] = (), named: dict[str, object] | None = None) -> object:
        self._check_alive()
        if name.casefold() == 'modify' and named:
            allowed = {'type', 'operator', 'formula1', 'formula2', 'string', 'operator2'}
            if any(key.casefold() not in allowed for key in named):
                raise error(1004)
        return super().vba_get(name, args, named)

    @member
    def Type(self) -> object:
        kind = self.rule.root.get('type', '')
        if kind in PREDICATE_TYPES:
            number = 9 if kind in TEXT_TYPES else {'containsBlanks': 10, 'timePeriod': 11,
                                                  'notContainsBlanks': 13, 'containsErrors': 16, 'notContainsErrors': 17}[kind]
            return VBAInt(number, 'Long')
        if kind not in ('cellIs', 'expression'):
            raise VBAUnsupportedError('This conditional formatting rule type is not implemented')
        return VBAInt(1 if kind == 'cellIs' else 2, 'Long')


    @member
    def Operator(self) -> object:
        if self.rule.root.get('type') != 'cellIs':
            raise error(1004)
        return VBAInt(OPERATORS.index(self.rule.root.get('operator', 'between')) + 1, 'Long')


    def _formula(self, number: int) -> str:
        formulas = self.rule.root.findall(NS + 'formula')
        if len(formulas) < number or formulas[number - 1].text is None:
            raise error(1004)
        return ('=' + decode(formulas[number - 1].text or '')).split('\x00', 1)[0]


    @member
    def Formula1(self) -> object:
        return self._formula(1)


    @member
    def Formula2(self) -> object:
        return self._formula(2)

    @member
    def Text(self) -> object:
        if self.rule.root.get('type') not in TEXT_TYPES:
            raise error(1004)
        return self.rule.root.get('text', '')

    @setter('Text')
    def set_text(self, value: object) -> None:
        guard(self.sheet)
        if value is NULL:
            raise error(13)
        if self.rule.root.get('type') not in TEXT_TYPES:
            raise error(1004)
        text = to_text(value)
        self.rule.root.set('text', text)
        regenerate(self.rule.root, formula_anchor(self.rule.areas))
        changed(self.sheet)

    @member
    def TextOperator(self) -> object:
        kind = self.rule.root.get('type', '')
        if kind not in TEXT_TYPES:
            raise error(1004)
        return VBAInt(TEXT_TYPES.index(kind), 'Long')

    @setter('TextOperator')
    def set_text_operator(self, value: object) -> None:
        guard(self.sheet)
        selected = _stat_long(value)
        if selected not in range(4):
            raise error(5)
        if self.rule.root.get('type') not in TEXT_TYPES:
            # Excel changes the rule before this unsupported transition fails.
            self.rule.root.set('type', TEXT_TYPES[selected])
            self.rule.root.set('operator', TEXT_OPERATORS[selected])
            self.rule.root.set('text', '')
            self.rule.root.attrib.pop('timePeriod', None)
            for child in self.rule.root.findall(NS + 'formula'):
                self.rule.root.remove(child)
            changed(self.sheet)
            raise error(1004)
        self.rule.root.set('type', TEXT_TYPES[selected])
        self.rule.root.set('operator', TEXT_OPERATORS[selected])
        regenerate(self.rule.root, formula_anchor(self.rule.areas))
        changed(self.sheet)

    @member
    def DateOperator(self) -> object:
        if self.rule.root.get('type') != 'timePeriod':
            raise error(5)
        period = self.rule.root.get('timePeriod', 'today')
        if period not in TIME_PERIODS:
            raise error(5)
        return VBAInt(TIME_PERIODS.index(period), 'Long')

    @setter('DateOperator')
    def set_date_operator(self, value: object) -> None:
        guard(self.sheet)
        selected = _stat_long(value)
        if selected not in range(len(TIME_PERIODS)):
            raise error(5)
        if self.rule.root.get('type') != 'timePeriod':
            return
        self.rule.root.set('timePeriod', TIME_PERIODS[selected])
        regenerate(self.rule.root, formula_anchor(self.rule.areas))
        changed(self.sheet)

    @member
    def ScopeType(self) -> object:
        raise error(5)

    @setter('ScopeType')
    def set_scope_type(self, value: object) -> None:
        guard(self.sheet)
        _stat_long(value)
        raise error(1004)


    @method
    def Modify(self, Type: object = MISSING, Operator: object = MISSING, Formula1: object = MISSING,
               Formula2: object = MISSING, String: object = MISSING, Operator2: object = MISSING) -> object:
        guard(self.sheet)
        if Type is MISSING:
            raise error(450)
        kind = int(to_integer(Type, 'Long'))
        if kind == 9:
            if self.rule.root.get('type') not in TEXT_TYPES:
                raise error(1004)
            if String is not MISSING:
                # Modify accepts Null as an empty term, unlike the Text setter.
                self.rule.root.set('text', '' if String is NULL else to_text(String))
                regenerate(self.rule.root, formula_anchor(self.rule.areas))
                changed(self.sheet)
            if Operator2 is not MISSING:
                self.set_text_operator(Operator2)
            return EMPTY
        if kind not in (1, 2):
            return EMPTY
        target = self.AppliesTo()
        assert isinstance(target, Range)
        made = _made(target, Type, Operator, Formula1, Formula2)
        for attribute in ('operator', 'text', 'timePeriod', 'rank', 'percent', 'bottom',
                          'aboveAverage', 'equalAverage', 'stdDev'):
            self.rule.root.attrib.pop(attribute, None)
        self.rule.root.attrib.update(made.attrib)
        self.rule.root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        self.rule.root.set('stopIfTrue', '1')
        for child in self.rule.root.findall(NS + 'formula'):
            self.rule.root.remove(child)
        for index, child in enumerate(made):
            self.rule.root.insert(index, child)
        changed(self.sheet)
        return EMPTY


def _stat_long(value: object) -> int:
    # These specialized COM properties reject Null as a type mismatch.
    if value is NULL:
        raise error(13)
    return int(to_integer(value, 'Long'))


class _StatisticalRule(_ConditionalRule):
    vba_type_name = ''

    @member
    def ScopeType(self) -> object:
        # A range-based rule has no PivotTable selection scope.
        raise error(5)

    @setter('ScopeType')
    def set_scope_type(self, value: object) -> None:
        guard(self.sheet)
        _stat_long(value)
        raise error(1004)


class Top10(_StatisticalRule):
    vba_type_name = 'Top10'

    @member
    def Type(self) -> object:
        return VBAInt(5, 'Long')

    @member
    def TopBottom(self) -> object:
        return VBAInt(0 if self.rule.root.get('bottom', '0') in ('1', 'true') else 1, 'Long')

    @setter('TopBottom')
    def set_top_bottom(self, value: object) -> None:
        guard(self.sheet)
        chosen = _stat_long(value)
        if chosen not in (0, 1):
            raise error(5)
        self.rule.root.attrib.pop('bottom', None)
        if chosen == 0:
            self.rule.root.set('bottom', '1')
        changed(self.sheet)

    @member
    def Rank(self) -> object:
        return VBAInt(int(self.rule.root.get('rank', '10')), 'Long')

    @setter('Rank')
    def set_rank(self, value: object) -> None:
        guard(self.sheet)
        chosen = _stat_long(value)
        if not 1 <= chosen <= (100 if self.Percent() else 1000):
            raise error(5)
        self.rule.root.set('rank', str(chosen))
        changed(self.sheet)

    @member
    def Percent(self) -> object:
        return self.rule.root.get('percent', '0') in ('1', 'true')

    @setter('Percent')
    def set_percent(self, value: object) -> None:
        guard(self.sheet)
        if value is NULL:
            raise error(13)
        selected = to_bool(value)
        self.rule.root.attrib.pop('percent', None)
        if selected:
            self.rule.root.set('percent', '1')
            self.rule.root.set('rank', str(min(100, int(self.Rank()))))
        changed(self.sheet)


class AboveAverage(_StatisticalRule):
    vba_type_name = 'AboveAverage'

    @member
    def Type(self) -> object:
        return VBAInt(12, 'Long')

    @member
    def AboveBelow(self) -> object:
        below = self.rule.root.get('aboveAverage', '1') in ('0', 'false')
        standard = int(self.rule.root.get('stdDev', '0')) > 0
        equal = self.rule.root.get('equalAverage', '0') in ('1', 'true')
        return VBAInt((4 if standard else 2 if equal else 0) + int(below), 'Long')

    @setter('AboveBelow')
    def set_above_below(self, value: object) -> None:
        guard(self.sheet)
        chosen = _stat_long(value)
        if chosen not in range(6):
            raise error(5)
        if chosen == int(self.AboveBelow()):
            return
        for attribute in ('aboveAverage', 'equalAverage', 'stdDev'):
            self.rule.root.attrib.pop(attribute, None)
        if chosen % 2:
            self.rule.root.set('aboveAverage', '0')
        if chosen in (2, 3):
            self.rule.root.set('equalAverage', '1')
        elif chosen in (4, 5):
            self.rule.root.set('stdDev', '1')
        changed(self.sheet)

    @member
    def NumStdDev(self) -> object:
        return VBAInt(int(self.rule.root.get('stdDev', '1')), 'Long')

    @setter('NumStdDev')
    def set_num_std_dev(self, value: object) -> None:
        guard(self.sheet)
        chosen = _stat_long(value)
        if int(self.AboveBelow()) not in (4, 5) or chosen not in (1, 2, 3):
            raise error(5)
        self.rule.root.set('stdDev', str(chosen))
        changed(self.sheet)


class UniqueValues(_StatisticalRule):
    vba_type_name = 'UniqueValues'

    @member
    def Type(self) -> object:
        return VBAInt(8, 'Long')

    @member
    def DupeUnique(self) -> object:
        from pyopenvba.apps.excel._conditional_border_context import bound_rule

        _, bound = bound_rule(self.sheet, self.rule)
        return VBAInt(int(bound.root.get('type') != 'uniqueValues'), 'Long')

    @setter('DupeUnique')
    def set_dupe_unique(self, value: object) -> None:
        guard(self.sheet)
        chosen = _stat_long(value)
        if chosen not in (0, 1):
            raise error(5)
        self.rule.root.set('type', 'duplicateValues' if chosen else 'uniqueValues')
        changed(self.sheet)


def _wrapper(sheet: Worksheet, rule: Rule) -> RuleObject:
    from pyopenvba.apps.excel._colour_scales import ColorScale
    from pyopenvba.apps.excel._data_bars import Databar
    from pyopenvba.apps.excel._icon_sets import IconSetCondition

    kind = rule.root.get('type', '')
    factory: type[RuleObject] = {'top10': Top10, 'aboveAverage': AboveAverage, 'uniqueValues': UniqueValues,
                                 'duplicateValues': UniqueValues, 'colorScale': ColorScale, 'dataBar': Databar,
                                 'iconSet': IconSetCondition}.get(kind, FormatCondition)
    return factory(sheet, rule)
