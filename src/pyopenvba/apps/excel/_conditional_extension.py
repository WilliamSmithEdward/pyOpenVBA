"""Conditional formats in Excel 2010's worksheet extension, which data bars, icon sets and colour scales need.

Measured in live Excel (scripts/measure_data_bars_icons.py, tests/fixtures/data_bars_icons/):

- A rule Excel 2007 could read is a cfRule in a conditionalFormatting element. A data bar that needs more keeps
  that cfRule, with an x14:id in its extLst naming an x14:cfRule in the worksheet's extLst that holds the rest.
  A rule Excel 2007 could not read at all -- an icon set of stars, triangles or boxes, custom icons, a point whose
  formula names another sheet -- is an x14:cfRule alone, which then carries the rule's priority itself.
- The extension groups its rules by range as the conditionalFormatting elements group theirs: a range's x14
  rules sit together whether or not they keep a cfRule too, in the same order, with the range last as xm:sqref.
  Rules the extension holds alone take their places among the rest, a new rule after those already there.
- A rule keeps its id through a save, and a new one takes a new GUID; one read without an id takes the id Excel
  makes of its sheet's place and its priority (loaded_identifier).
- A point's formula naming its own sheet loses the sheet's name as the point takes it.

A worksheet whose extension holds rules this module does not read keeps its extension as the file had it, and
those rules stay out of FormatConditions as before; a data bar or icon set there reports itself unsupported
(see foreign).
"""
from __future__ import annotations

import re
import uuid
from abc import ABC, abstractmethod
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba._xml import escape, escape_text
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._conditional import NS, Rule
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, split_sheet, tokenize

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

X14_URI = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM_URI = 'http://schemas.microsoft.com/office/excel/2006/main'
X14 = '{' + X14_URI + '}'
XM = '{' + XM_URI + '}'
#: The worksheet ext holding x14 conditional formats, and the cfRule ext naming a rule's x14:cfRule.
RULES_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
LINK_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
#: The rule types this module reads from the extension.
KNOWN = ('dataBar', 'iconSet', 'colorScale')
_PREFIXES = {X14_URI: 'x14:', XM_URI: 'xm:', NS[1:-1]: ''}


class RuleState(ABC):
    """A data bar's or an icon set's settings, read from a file or made by VBA; dirty once VBA changes one, when
    the file keeps the rule as Excel writes it rather than as it was read."""

    dirty: bool = False
    #: Whether the rule shows a cell's value beside what it draws.
    show_value: bool = True

    @abstractmethod
    def parts(self, rule: Rule) -> tuple[XmlElement | None, XmlElement | None]:
        """The rule's cfRule and its x14:cfRule as a file keeps them, either None where it keeps none."""


def new_identifier() -> str:
    """A GUID for a rule's x14:cfRule, as Excel writes one."""
    return '{' + str(uuid.uuid4()).upper() + '}'


def loaded_identifier(sheet: int, priority: int) -> str:
    """The id Excel gives a rule it reads without one, and writes once the rule needs the extension: zeros, 000E,
    then the sheet's place among the workbook's sheets from 0 and the rule's priority as it was read, each four
    bytes low first (tests/fixtures/data_bars_icons/data_bars_icons_resaved.xlsx)."""
    data = sheet.to_bytes(4, 'little').hex().upper() + priority.to_bytes(4, 'little').hex().upper()
    return '{00000000-000E-0000-' + data[:4] + '-' + data[4:] + '}'


def identified(rule: Rule) -> str:
    """The rule's id, made the first time it needs one."""
    if not rule.identifier:
        rule.identifier = new_identifier()
    return rule.identifier


def linked(root: XmlElement) -> str:
    """The id a cfRule's extLst names, or ''."""
    for ext in root.findall(NS + 'extLst/' + NS + 'ext'):
        if ext.get('uri') == LINK_URI:
            found = ext.find(X14 + 'id')
            return (found.text or '') if found is not None else ''
    return ''


def link(root: XmlElement, identifier: str) -> None:
    """Make a cfRule's extLst name its x14:cfRule."""
    ext_list = XmlElement(NS + 'extLst')
    ext = XmlElement(NS + 'ext', {'uri': LINK_URI, 'xmlns:x14': X14_URI})
    named = XmlElement(X14 + 'id')
    named.text = identifier
    ext.append(named)
    ext_list.append(ext)
    root.append(ext_list)


# --- a point's formula --------------------------------------------------------------------------------------

#: A number as a point's value spells one.
NUMBER = re.compile(r'-?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?')


def given_text(value: str, sheet: str) -> str:
    """The formula a point keeps of text VBA gives it, as Excel keeps it: after an "=", the formula without its
    own sheet's name; else a number as it stands, and other text as a string."""
    if value.startswith('='):
        return without_own_sheet(value[1:], sheet)
    if NUMBER.fullmatch(value):
        return value
    return '"' + value.replace('"', '""') + '"'


def without_own_sheet(text: str, sheet: str) -> str:
    """A point's formula with its own sheet's name taken from its references, as Excel takes it."""
    try:
        tokens = tokenize(text, spaces=True)
    except FormulaError:
        return text
    out: list[str] = []
    for token in tokens:
        if token.kind in ('ref', 'name', 'error') and '!' in token.text:
            head, rest = split_sheet(token.text)
            if head.casefold() == sheet.casefold():
                out.append(rest)
                continue
        out.append(token.text)
    return ''.join(out)


def point_formulas(rule: Rule) -> list[str]:
    """Every formula a rule holds as text: its own, and its points' where it is a colour scale, a data bar or an
    icon set, for what renaming a sheet or moving cells must keep right."""
    from pyopenvba.apps.excel._xstring import decode

    texts = [decode(one.text or '') for one in rule.root.findall(NS + 'formula')]
    kind = rule.root.get('type')
    if kind == 'colorScale':
        texts += [point.get('val', '') for point in rule.root.findall(NS + 'colorScale/' + NS + 'cfvo')
                  if point.get('type') not in ('min', 'max') and not NUMBER.fullmatch(point.get('val', ''))]
    elif kind in ('dataBar', 'iconSet'):
        from pyopenvba.apps.excel._data_bars import Bar
        from pyopenvba.apps.excel._data_bars import parse as parse_bar
        from pyopenvba.apps.excel._icon_sets import Icons
        from pyopenvba.apps.excel._icon_sets import parse as parse_icons

        if rule.state is None:
            rule.state = parse_bar(rule) if kind == 'dataBar' else parse_icons(rule)
        state = rule.state
        values: list[float | str | None] = []
        if isinstance(state, Bar):
            values = [point.value for point in state.points]
        elif isinstance(state, Icons):
            values = [criterion.value for criterion in state.criteria]
        texts += [value for value in values if isinstance(value, str) and value]
    return texts


def names_a_sheet(text: str) -> bool:
    """Whether a point's formula names a sheet, which Excel 2007 could not read."""
    try:
        tokens = tokenize(text)
    except FormulaError:
        return False
    return any(token.kind in ('ref', 'name', 'error') and '!' in token.text for token in tokens)


# --- reading ------------------------------------------------------------------------------------------------


def _rules_ext(root: XmlElement) -> XmlElement | None:
    ext_list = root.find(NS + 'extLst')
    if ext_list is None:
        return None
    return next((ext for ext in ext_list.findall(NS + 'ext') if ext.get('uri') == RULES_URI), None)


def read(root: XmlElement, rules: list[Rule], *, imported: bool, plain_load: bool | None) -> bool:
    """Pair each rule with its x14:cfRule and add the rules the extension holds alone; False, changing nothing,
    when the extension holds a rule this module does not read."""
    from pyopenvba.apps.excel._colour_scales import extension_known as scale_known
    from pyopenvba.apps.excel._data_bars import extension_known as bar_known
    from pyopenvba.apps.excel._icon_sets import extension_known as icons_known

    ext = _rules_ext(root)
    if ext is None:
        return True
    by_id = {identifier: rule for rule in rules if (identifier := linked(rule.root))}
    found: list[tuple[str, XmlElement, tuple[int, str, int]]] = []
    for holder in ext.findall(X14 + 'conditionalFormattings'):
        for number, group in enumerate(holder.findall(X14 + 'conditionalFormatting')):
            sqref = group.find(XM + 'sqref')
            if sqref is None or not (sqref.text or '').split() or any(
                    child.tag not in (X14 + 'cfRule', XM + 'sqref') for child in group):
                return False
            found.extend((sqref.text or '', entry, (number, sqref.text or '', place))
                         for place, entry in enumerate(group.findall(X14 + 'cfRule')))
    checks = {'dataBar': bar_known, 'iconSet': icons_known, 'colorScale': scale_known}
    for _, entry, _ in found:
        kind = entry.get('type', '')
        paired = entry.get('id', '') in by_id
        if kind not in checks or not checks[kind](entry, paired=paired) or (paired and kind != 'dataBar'):
            return False
    for sqref, entry, place in found:
        identifier = entry.get('id', '')
        owner = by_id.get(identifier)
        if owner is None:
            try:
                areas = [parse_area(value) for value in sqref.split()]
            except ValueError:
                return False
            owner = alone(areas, entry, imported=imported, plain_load=plain_load)
            rules.append(owner)
        owner.extension, owner.identifier, owner.extension_place = entry, identifier, place
    return True


def alone(areas: list[Area], entry: XmlElement, *, imported: bool, plain_load: bool | None) -> Rule:
    """A rule the extension holds alone, its cfRule a stand-in carrying its type and priority; a colour scale's
    stand-in holds its points and colours, as the colour-scale code reads them (see _colour_scales)."""
    root = XmlElement(NS + 'cfRule', {'type': entry.get('type', ''), 'priority': entry.get('priority', '1')})
    if entry.get('type') == 'colorScale':
        from pyopenvba.apps.excel._colour_scales import from_extension

        root.append(from_extension(entry))
    return Rule(list(areas), root, None, imported=imported, loaded=True, plain_load=plain_load, extension=entry,
                extension_only=True, identifier=entry.get('id', ''))


def foreign(sheet: Worksheet) -> None:
    """Refuse a rule that needs the extension on a sheet whose extension holds rules this module does not read."""
    if sheet.conditional_foreign:
        raise VBAUnsupportedError('Data bars, icon sets and colour scales on a worksheet whose Excel 2010 '
                                  'conditional-format extension holds other rules are not implemented')


# --- writing ------------------------------------------------------------------------------------------------


def parts(rule: Rule) -> tuple[XmlElement | None, XmlElement | None]:
    """The rule's cfRule and x14:cfRule as a file keeps them, either None where it keeps none: as Excel writes
    them for a rule VBA made or changed, else as they were read, with the rule's priority."""
    if rule.state is not None and rule.state.dirty:
        return rule.state.parts(rule)
    if rule.root.get('type') == 'colorScale':
        from pyopenvba.apps.excel._colour_scales import placed

        return placed(rule)
    if rule.extension_only and rule.extension is not None:
        entry = deepcopy(rule.extension)
        entry.set('priority', rule.root.get('priority', '1'))
        return None, entry
    return deepcopy(rule.root), deepcopy(rule.extension) if rule.extension is not None else None


def grouped(placed: list[tuple[Rule, XmlElement]]) -> list[tuple[str, list[XmlElement]]]:
    """The extension's x14:cfRules by range, as Excel groups them: the groups a file held in its order, each rule
    in its place, a rule made since or moved joining the group of its range at its end, and a group for each
    other range after them (tests/fixtures/data_bars_icons/data_bars_icons_resaved.xlsx)."""
    from pyopenvba.apps.excel._conditional import projected_areas

    held: dict[int, tuple[str, list[tuple[int, XmlElement]]]] = {}
    later: list[tuple[Rule, str, XmlElement]] = []
    for rule, entry in placed:
        sqref = ' '.join(area.address(absolute=False) for area in projected_areas(rule))
        place = rule.extension_place
        if place is not None and place[1] == sqref:
            held.setdefault(place[0], (sqref, []))[1].append((place[2], entry))
        else:
            later.append((rule, sqref, entry))
    groups = [(sqref, [entry for _, entry in sorted(members, key=lambda member: member[0])])
              for _, (sqref, members) in sorted(held.items())]
    for _, sqref, entry in later:
        found = next((members for range_, members in groups if range_ == sqref), None)
        if found is None:
            groups.append((sqref, [entry]))
        else:
            found.append(entry)
    return groups


def _markup(element: XmlElement) -> str:
    """An x14 element as Excel writes it, in the prefixes the extension declares."""
    def name(expanded: str) -> str:
        if not expanded.startswith('{'):
            return expanded
        uri, local = expanded[1:].split('}', 1)
        return _PREFIXES[uri] + local

    tag = name(element.tag)
    attributes = ''.join(f' {name(key)}="{escape(value)}"' for key, value in element.attrib.items()
                         if key != 'xmlns' and not key.startswith('xmlns:'))
    inner = escape_text(element.text or '') + ''.join(_markup(child) for child in element)
    return f'<{tag}{attributes}>{inner}</{tag}>' if inner else f'<{tag}{attributes}/>'


def markup(groups: list[tuple[str, list[XmlElement]]]) -> str:
    """The worksheet's ext of x14 rules, a range's rules to each conditionalFormatting; '' when there are none."""
    if not groups:
        return ''
    body = ''.join(f'<x14:conditionalFormatting xmlns:xm="{XM_URI}">' + ''.join(_markup(one) for one in entries)
                   + f'<xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting>' for sqref, entries in groups)
    return (f'<ext uri="{RULES_URI}" xmlns:x14="{X14_URI}"><x14:conditionalFormattings>{body}'
            '</x14:conditionalFormattings></ext>')


def x14_color(element: XmlElement, local: str) -> XmlElement:
    """A colour as an x14 element of that name."""
    return XmlElement(X14 + local, dict(element.attrib))


def x14_point(kind: str, value: str | None, *, greater: bool = False) -> XmlElement:
    """An x14:cfvo: its type, gte="0" for a greater-than threshold, and its value as an xm:f."""
    point = XmlElement(X14 + 'cfvo', {'type': kind, **({'gte': '0'} if greater else {})})
    if value is not None:
        formula = XmlElement(XM + 'f')
        formula.text = value
        point.append(formula)
    return point
