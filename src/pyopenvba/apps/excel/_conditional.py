"""Worksheet conditional rules and Excel's ordinary FormatConditions API.

The file model follows William's pyOfficeEditor: rules retain their original
XML, including extensions, and differential styles live in the workbook dxfs.
Runtime priorities and range deletion follow independent native Excel controls.
This module describes rules; it does not claim to render DisplayFormat.
"""
# ruff: noqa: N803 -- Excel's public named arguments use these exact names.
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_parser, xml_tree
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.apps.excel._xstring import decode, encode_text
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
OPERATORS = ('between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan',
             'greaterThanOrEqual', 'lessThanOrEqual')
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


def read_rules(xml: str) -> list[Rule]:
    root = xml_tree(xml)
    rules: list[Rule] = []
    for group in root.findall(NS + 'conditionalFormatting'):
        # A detached rule may use ancestor-only prefixes in MC/QName values.
        for key, value in root.attrib.items():
            if key == 'xmlns' or key.startswith('xmlns:'):
                group.attrib.setdefault(key, value)
        areas = [parse_area(value) for value in group.get('sqref', '').split()]
        if areas:
            rules.extend(Rule(list(areas), entry, group) for entry in group.findall(NS + 'cfRule'))
    return rules


def with_rules(xml: str, rules: list[Rule]) -> str:
    """Replace standard rules while retaining container and extension metadata."""
    groups: list[tuple[XmlElement | None, list[Area], XmlElement]] = []
    for rule in rules:
        if rule.deleted or not rule.areas:
            continue
        group = next((output for original, areas, output in groups
                      if original is rule.container and areas == rule.areas), None)
        if group is None:
            group = deepcopy(rule.container) if rule.container is not None else XmlElement(NS + 'conditionalFormatting')
            for entry in group.findall(NS + 'cfRule'):
                group.remove(entry)
            group.set('sqref', ' '.join(area.address(absolute=False) for area in rule.areas))
            groups.append((rule.container, list(rule.areas), group))
        # cfRule precedes extLst in CT_ConditionalFormatting.
        before = next((i for i, entry in enumerate(group) if entry.tag == NS + 'extLst'), len(group))
        group.insert(before, deepcopy(rule.root))
    data = xml.encode('utf-8')
    parser = xml_parser()
    depth, active = 0, None
    root_start = 0
    spans: list[tuple[int, int]] = []
    insertion: int | None = None

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

    def start(name: str, _attrs: dict[str, str]) -> None:
        nonlocal depth, active, insertion, root_start
        depth += 1
        if depth == 1:
            root_start = parser.CurrentByteIndex
        if depth == 2:
            if name == NS[1:] + 'conditionalFormatting':
                active = parser.CurrentByteIndex
            elif name.startswith(NS[1:]) and name[len(NS) - 1:] in _AFTER and insertion is None:
                insertion = parser.CurrentByteIndex

    def end(_name: str) -> None:
        nonlocal depth, active, insertion
        if depth == 2 and active is not None:
            first_end = tag_end(active)
            stop = first_end if data[active:first_end].rstrip().endswith(b'/>') else tag_end(parser.CurrentByteIndex)
            spans.append((active, stop))
            active = None
        if depth == 1 and insertion is None:
            insertion = parser.CurrentByteIndex
        depth -= 1

    parser.StartElementHandler, parser.EndElementHandler = start, end
    parser.Parse(data, True)
    assert insertion is not None
    markup = b''.join(xml_bytes(group).split(b'?>', 1)[1].lstrip() for _, _, group in groups)
    root_end = tag_end(root_start)
    if data[root_start:root_end].endswith(b'/>'):
        if not markup:
            return xml
        tag = data[root_start + 1:root_end].split(None, 1)[0].rstrip(b'/>')
        return (data[:root_end - 2] + b'>' + markup + b'</' + tag + b'>' + data[root_end:]).decode('utf-8')
    edits = [(first, last, b'') for first, last in spans]
    edits.append((insertion, insertion, markup))
    for first, last, replacement in sorted(edits, reverse=True):
        data = data[:first] + replacement + data[last:]
    return data.decode('utf-8')


def _priority(rule: Rule) -> int:
    return int(rule.root.get('priority', '0'))


def _changed(sheet: Worksheet) -> None:
    sheet.conditional_changed = True
    sheet.touched()


def _guard(sheet: Worksheet) -> None:
    from pyopenvba.apps.excel._protection import APPLICATION_DEFINED

    protection = sheet.protection
    # Conditional formatting follows AllowFormattingCells even for UI-only
    # protection, where many ordinary macro writes would otherwise be allowed.
    if protection is not None and protection.contents and 'AllowFormattingCells' not in protection.allows:
        raise error(1004, APPLICATION_DEFINED)


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def _without(area: Area, cut: Area) -> list[Area]:
    """Native CF split order: top, bottom, left, right (unlike validation)."""
    if not _meets(area, cut):
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
                pieces = _without(area, cut)
                kept.extend(pieces[:1])
                extra.extend(pieces[1:])
            remaining = kept + extra
        if remaining != rule.areas:
            formulas = [child.text or '' for child in rule.root.findall(NS + 'formula')]
            if remaining:
                original, first = rule.areas[0], remaining[0]
                if (original.top, original.left) != (first.top, first.left):
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
        for child, text in zip(rule.root.findall(NS + 'formula'), formulas, strict=True):
            child.text = text
        if not remaining:
            rule.deleted = True
            sheet.conditional_formats.remove(rule)
    if planned:
        _changed(sheet)
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


def _made(target: Range, Type: object, Operator: object, Formula1: object, Formula2: object) -> XmlElement:
    if Type is MISSING:
        raise error(449)
    kind = int(to_integer(Type, 'Long'))
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
                 and any(_meets(one, area) for one in rule.areas for area in self.target.areas)]
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
            Formula2: object = MISSING) -> object:
        _guard(self.sheet)
        root = _made(self.target, Type, Operator, Formula1, Formula2)
        return self._append(root)

    def _append(self, root: XmlElement) -> object:
        root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        rule = Rule([Area(area.top, area.left, area.bottom, area.right) for area in self.target.areas], root)
        self.sheet.conditional_formats.append(rule)
        _changed(self.sheet)
        return _wrapper(self.sheet, rule)

    @method
    def AddTop10(self) -> object:
        _guard(self.sheet)
        return self._append(_made(self.target, 5, MISSING, MISSING, MISSING))

    @method
    def AddAboveAverage(self) -> object:
        _guard(self.sheet)
        return self._append(_made(self.target, 12, MISSING, MISSING, MISSING))

    @method
    def AddUniqueValues(self) -> object:
        _guard(self.sheet)
        return self._append(_made(self.target, 8, MISSING, MISSING, MISSING))

    @method
    def Delete(self) -> object:
        _guard(self.sheet)
        remove(self.sheet, self.target.areas)
        return EMPTY


class _ConditionalRule(ExcelObject):
    vba_type_name = ''

    def __init__(self, sheet: Worksheet, rule: Rule) -> None:
        self.sheet, self.rule = sheet, rule

    @member
    def Parent(self) -> object:
        return self.AppliesTo()


    @member
    def AppliesTo(self) -> object:
        return Range(self.sheet, [Area(a.top, a.left, a.bottom, a.right, self.sheet.name) for a in self.rule.areas])

    @member
    def Priority(self) -> object:
        return VBAInt(_priority(self.rule), 'Long')

    @setter('Priority')
    def set_priority(self, value: object) -> None:
        _guard(self.sheet)
        wanted = int(to_integer(value, 'Long'))
        if wanted < 1:
            raise error(5)
        for other in self.sheet.conditional_formats:
            priority = _priority(other)
            if other is not self.rule and wanted <= priority:
                other.root.set('priority', str(priority + 1))
        self.rule.root.set('priority', str(wanted))
        _changed(self.sheet)

    @member
    def StopIfTrue(self) -> object:
        return self.rule.root.get('stopIfTrue', '0') in ('1', 'true')

    @setter('StopIfTrue')
    def set_stop(self, value: object) -> None:
        _guard(self.sheet)
        self.rule.root.set('stopIfTrue', str(int(to_bool(value))))
        _changed(self.sheet)

    @method
    def SetFirstPriority(self) -> object:
        _guard(self.sheet)
        for other in self.sheet.conditional_formats:
            if other is not self.rule:
                other.root.set('priority', str(_priority(other) + 1))
        self.rule.root.set('priority', '1')
        _changed(self.sheet)
        return EMPTY

    @method
    def SetLastPriority(self) -> object:
        _guard(self.sheet)
        self.rule.root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        _changed(self.sheet)
        return EMPTY

    @method
    def Delete(self) -> object:
        _guard(self.sheet)
        if self.rule not in self.sheet.conditional_formats:
            raise error(1004)
        self.sheet.conditional_formats.remove(self.rule)
        self.rule.deleted = True
        _changed(self.sheet)
        return EMPTY


    @method
    def ModifyAppliesToRange(self, Range: object = MISSING) -> object:
        from pyopenvba.apps.excel._model import Range as ExcelRange

        _guard(self.sheet)
        if not isinstance(Range, ExcelRange) or Range.sheet is not self.sheet:
            raise error(1004)
        self.rule.areas = [Area(a.top, a.left, a.bottom, a.right) for a in Range.areas]
        _changed(self.sheet)
        return EMPTY

    def _dxf(self) -> XmlElement:
        if self.rule.dxf is None:
            identifier = self.rule.root.get('dxfId')
            self.rule.dxf = (self.sheet.book.stylesheet.dxf(int(identifier)) if identifier is not None
                             else XmlElement(NS + 'dxf'))
        return self.rule.dxf

    def _style_changed(self) -> None:
        self.rule.root.set('dxfId', str(self.sheet.book.stylesheet.ensure_dxf(self._dxf())))
        _changed(self.sheet)

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
        from pyopenvba.apps.excel._differential import ConditionalBorders

        collection = ConditionalBorders(self._dxf(), self._style_changed, self.sheet)
        return collection if Index is MISSING else collection.Item(Index)

    @member
    def NumberFormat(self) -> object:
        from pyopenvba.apps.excel._differential import number_format

        return number_format(self._dxf(), self.sheet.book.stylesheet)

    @setter('NumberFormat')
    def set_number_format(self, value: object) -> None:
        from pyopenvba.apps.excel._differential import set_number_format

        _guard(self.sheet)
        set_number_format(self._dxf(), value, self.sheet.book.stylesheet)
        self._style_changed()


class FormatCondition(_ConditionalRule):
    vba_type_name = 'FormatCondition'

    @member
    def Type(self) -> object:
        kind = self.rule.root.get('type', '')
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
        return '=' + decode(formulas[number - 1].text or '')


    @member
    def Formula1(self) -> object:
        return self._formula(1)


    @member
    def Formula2(self) -> object:
        return self._formula(2)


    @method
    def Modify(self, Type: object = MISSING, Operator: object = MISSING, Formula1: object = MISSING,
               Formula2: object = MISSING) -> object:
        _guard(self.sheet)
        if Type is MISSING:
            raise error(450)
        if Type is not MISSING and int(to_integer(Type, 'Long')) not in (1, 2):
            return EMPTY
        target = self.AppliesTo()
        assert isinstance(target, Range)
        made = _made(target, Type, Operator, Formula1, Formula2)
        self.rule.root.attrib.pop('operator', None)
        self.rule.root.attrib.update(made.attrib)
        self.rule.root.set('priority', str(max((_priority(rule) for rule in self.sheet.conditional_formats), default=0) + 1))
        self.rule.root.set('stopIfTrue', '1')
        for child in self.rule.root.findall(NS + 'formula'):
            self.rule.root.remove(child)
        for index, child in enumerate(made):
            self.rule.root.insert(index, child)
        _changed(self.sheet)
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
        _guard(self.sheet)
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
        _guard(self.sheet)
        chosen = _stat_long(value)
        if chosen not in (0, 1):
            raise error(5)
        self.rule.root.attrib.pop('bottom', None)
        if chosen == 0:
            self.rule.root.set('bottom', '1')
        _changed(self.sheet)

    @member
    def Rank(self) -> object:
        return VBAInt(int(self.rule.root.get('rank', '10')), 'Long')

    @setter('Rank')
    def set_rank(self, value: object) -> None:
        _guard(self.sheet)
        chosen = _stat_long(value)
        if not 1 <= chosen <= (100 if self.Percent() else 1000):
            raise error(5)
        self.rule.root.set('rank', str(chosen))
        _changed(self.sheet)

    @member
    def Percent(self) -> object:
        return self.rule.root.get('percent', '0') in ('1', 'true')

    @setter('Percent')
    def set_percent(self, value: object) -> None:
        _guard(self.sheet)
        if value is NULL:
            raise error(13)
        selected = to_bool(value)
        self.rule.root.attrib.pop('percent', None)
        if selected:
            self.rule.root.set('percent', '1')
            self.rule.root.set('rank', str(min(100, int(self.Rank()))))
        _changed(self.sheet)


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
        _guard(self.sheet)
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
        _changed(self.sheet)

    @member
    def NumStdDev(self) -> object:
        return VBAInt(int(self.rule.root.get('stdDev', '1')), 'Long')

    @setter('NumStdDev')
    def set_num_std_dev(self, value: object) -> None:
        _guard(self.sheet)
        chosen = _stat_long(value)
        if int(self.AboveBelow()) not in (4, 5) or chosen not in (1, 2, 3):
            raise error(5)
        self.rule.root.set('stdDev', str(chosen))
        _changed(self.sheet)


class UniqueValues(_StatisticalRule):
    vba_type_name = 'UniqueValues'

    @member
    def Type(self) -> object:
        return VBAInt(8, 'Long')

    @member
    def DupeUnique(self) -> object:
        return VBAInt(int(self.rule.root.get('type') == 'duplicateValues'), 'Long')

    @setter('DupeUnique')
    def set_dupe_unique(self, value: object) -> None:
        _guard(self.sheet)
        chosen = _stat_long(value)
        if chosen not in (0, 1):
            raise error(5)
        self.rule.root.set('type', 'duplicateValues' if chosen else 'uniqueValues')
        _changed(self.sheet)


def _wrapper(sheet: Worksheet, rule: Rule) -> _ConditionalRule:
    kind = rule.root.get('type', '')
    factory = {'top10': Top10, 'aboveAverage': AboveAverage, 'uniqueValues': UniqueValues,
               'duplicateValues': UniqueValues}.get(kind, FormatCondition)
    return factory(sheet, rule)
