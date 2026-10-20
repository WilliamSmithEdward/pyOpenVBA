"""Plan the measured, disjoint single-block conditional-format replacement copy.

Admission is intentionally narrower than ordinary cell copying. Planning never
changes cells, rule identities, names, differential-style IDs or dirty flags.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule, changed, formula_anchor, guard, meets, without
from pyopenvba.apps.excel._conditional_predicates import PREDICATE_TYPES, regenerate
from pyopenvba.apps.excel._xstring import decode, escape
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, shift_text
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Worksheet

_KINDS = {'cellIs', 'expression', 'top10', 'aboveAverage', 'uniqueValues', 'duplicateValues'} | PREDICATE_TYPES
_ATTRIBUTES = {'type', 'priority', 'dxfId', 'stopIfTrue', 'operator', 'text', 'timePeriod',
               'rank', 'percent', 'bottom', 'aboveAverage', 'equalAverage', 'stdDev'}
_DXF_TAGS = {NS + tag for tag in ('dxf', 'font', 'b', 'i', 'strike', 'color', 'sz', 'u', 'vertAlign', 'name',
                 'family', 'charset', 'scheme', 'outline', 'shadow', 'condense', 'extend',
                 'fill', 'patternFill', 'fgColor', 'bgColor', 'border', 'left', 'right', 'top',
                 'bottom', 'start', 'end', 'diagonal', 'vertical', 'horizontal', 'numFmt')}
_DXF_ATTRIBUTES = {'val', 'rgb', 'theme', 'indexed', 'auto', 'tint', 'numFmtId', 'formatCode',
                   'patternType', 'style', 'diagonalUp', 'diagonalDown', 'outline'}


def _unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Conditional-format copying ' + detail + ' is not implemented')


def _touches(rule: Rule, area: Area) -> bool:
    return not rule.deleted and any(meets(part, area) for part in rule.areas)


def _attributes_known(root: XmlElement, allowed: set[str]) -> bool:
    return all(key in allowed or key == 'xmlns' or key.startswith('xmlns:') for key in root.attrib)


def _validate(rule: Rule, dxf_count: int) -> None:
    if rule.root.get('type') not in _KINDS or not _attributes_known(rule.root, _ATTRIBUTES):
        raise _unsupported('with unsupported rule attributes or families')
    if any(child.tag != NS + 'formula' or list(child) or not _attributes_known(child, set())
           for child in rule.root):
        raise _unsupported('with rule extensions')
    if rule.container is not None and (
            not _attributes_known(rule.container, {'sqref'})
            or any(child.tag != NS + 'cfRule' for child in rule.container)):
        raise _unsupported('with container extensions or pivot scopes')
    try:
        if int(rule.root.get('priority', '0')) <= 0:
            raise _unsupported('with invalid priorities')
        identifier = rule.root.get('dxfId')
        if identifier is not None and not 0 <= int(identifier) < dxf_count:
            raise _unsupported('with invalid differential-style identifiers')
    except ValueError:
        raise _unsupported('with invalid priorities or style identifiers') from None


def guard_extensions(sheet: Worksheet) -> None:
    package = sheet.book.package
    if package is None or not sheet.part_name or not package.has(sheet.part_name):
        return
    entry = package.entry(sheet.part_name)
    if sheet.conditional_extension_entry is not entry or sheet.conditional_extensions is None:
        root = xml_tree(entry.read())
        # read_rules models only direct worksheet children. Main-namespace
        # rules inside MC wrappers are just as unmodeled as foreign rules.
        direct = {id(node) for node in root if node.tag == NS + 'conditionalFormatting'}
        sheet.conditional_extensions = any(
            node.tag.rsplit('}', 1)[-1] == 'conditionalFormattings'
            or (node.tag.rsplit('}', 1)[-1] == 'conditionalFormatting' and id(node) not in direct)
            for node in root.iter())
        sheet.conditional_extension_entry = entry
    if sheet.conditional_extensions:
        raise _unsupported('with worksheet conditional extensions')


def _layers(sheet: Worksheet, area: Area, *, content: bool) -> None:
    if any(meets(area, merged) for merged in sheet.merged_areas):
        raise _unsupported('involving merged cells')
    if any(meets(area, table.area) for table in sheet.tables):
        raise _unsupported('involving tables')
    if any(meets(area, array) for array in sheet.array_formulas.values()) or any(
            meets(area, spill.area) for spill in sheet.spills.values()):
        raise _unsupported('involving arrays or spills')
    if not content:
        return
    from pyopenvba.apps.excel._hyperlinks import touches

    if (any(meets(area, part) for rule in sheet.validations for part in rule.areas)
            or any(area.contains(*position) for position in sheet.notes) or touches(sheet, [area])):
        raise _unsupported('with affected validation, notes or hyperlinks')


def _root_at(rule: Rule, areas: list[Area]) -> XmlElement:
    root = deepcopy(rule.root)
    if root.get('type') in PREDICATE_TYPES:
        regenerate(root, formula_anchor(areas))
    elif root.get('type') in ('cellIs', 'expression'):
        old, new = formula_anchor(rule.areas), formula_anchor(areas)
        for child in root.findall(NS + 'formula'):
            if child.text:
                child.text = escape(shift_text('=' + decode(child.text), new.top - old.top,
                                               new.left - old.left)[1:])
    return root


@dataclass
class ConditionalPlan:
    sheet: Worksheet
    retained: list[tuple[Rule, list[Area], XmlElement]]
    incoming: Rule | None

    def commit(self) -> None:
        """Apply a fully admitted plan, retaining existing rule/proxy identities."""
        from pyopenvba.apps.excel._conditional_border_context import register

        for rule, areas, root in self.retained:
            rule.areas = areas
            rule.root.attrib.clear()
            rule.root.attrib.update(root.attrib)
            rule.root[:] = list(root)
            if not areas:
                rule.deleted = True
                self.sheet.conditional_formats.remove(rule)
        if self.incoming is not None:
            register(self.sheet, self.incoming)
            self.sheet.conditional_formats.append(self.incoming)
        changed(self.sheet)


def plan(source: Worksheet, areas: list[Area], destination: Worksheet, landed: list[Area], *,
         kind: int = -4104, transpose: bool = False, operation: str = '') -> ConditionalPlan | None:
    """Stage replacement; unsupported mappings fail before any ordinary copy layer."""
    incoming = [rule for rule in source.conditional_formats
                if any(_touches(rule, area) for area in areas)]
    affected = [rule for rule in destination.conditional_formats
                if any(_touches(rule, area) for area in landed)]
    if (incoming or affected) and destination.book.application.udf_callers:
        raise error(1004, 'a worksheet function cannot change cells')
    guard_extensions(source)
    if destination is not source:
        guard_extensions(destination)
    if not incoming and not affected:
        return None
    if (source.book is not destination.book or len(areas) != 1 or len(landed) != 1
            or kind not in (-4104, -4122) or transpose or operation):
        raise _unsupported('with this paste mode, mapping or workbook boundary')
    first, last = areas[0], landed[0]
    if (first.rows, first.columns) != (last.rows, last.columns):
        raise _unsupported('with tiled or clipped worksheet-edge mappings')
    if source is destination and (meets(first, last) or any(_touches(rule, last) for rule in incoming)):
        raise _unsupported('with overlapping cells or shared source/destination rules')
    if len(incoming) > 1:
        raise _unsupported('with multiple incoming rules')
    guard(destination)
    for owner, area in ((source, first), (destination, last)):
        _layers(owner, area, content=kind == -4104)
    for rule in incoming + (list(destination.conditional_formats) if incoming else affected):
        if not rule.deleted:
            _validate(rule, source.book.stylesheet.dxf_count)
    try:
        # Copy formula spelling can fail late in ordinary copy paths. Validate it
        # before either those paths or a differential-style mutation can begin.
        if kind == -4104:
            for (row, column), cell in source.cells_.items():
                if first.contains(row, column) and cell.formula:
                    shift_text(cell.formula, last.top - first.top, last.left - first.left)
        copied: Rule | None = None
        if incoming:
            original = incoming[0]
            mapped = [Area(max(part.top, first.top) + last.top - first.top,
                           max(part.left, first.left) + last.left - first.left,
                           min(part.bottom, first.bottom) + last.top - first.top,
                           min(part.right, first.right) + last.left - first.left)
                      for part in original.areas if meets(part, first)]
            copied = Rule(mapped, _root_at(original, mapped), deepcopy(original.container))
            copied.root.set('priority', '1')
            identifier = original.root.get('dxfId')
            try:
                copied.dxf = (deepcopy(original.dxf) if original.dxf is not None
                              else source.book.stylesheet.dxf(int(identifier)) if identifier is not None else None)
            except IndexError:
                raise _unsupported('with invalid differential-style identifiers') from None
            if copied.dxf is not None and any(node.tag not in _DXF_TAGS
                                              or not _attributes_known(node, _DXF_ATTRIBUTES)
                                              for node in copied.dxf.iter()):
                raise _unsupported('with differential-style extensions')
        retained: list[tuple[Rule, list[Area], XmlElement]] = []
        for rule in destination.conditional_formats:
            if rule.deleted:
                continue
            kept: list[Area] = []
            extra: list[Area] = []
            for part in rule.areas:
                pieces = without(part, last)
                kept.extend(pieces[:1])
                extra.extend(pieces[1:])
            remaining = kept + extra
            if copied is None and remaining == rule.areas:
                continue
            root = _root_at(rule, remaining) if remaining else deepcopy(rule.root)
            if copied is not None and remaining:
                root.set('priority', str(int(root.get('priority', '0')) + 1))
            if copied is not None or remaining != rule.areas:
                retained.append((rule, remaining, root))
        return ConditionalPlan(destination, retained, copied)
    except (FormulaError, ValueError):
        raise _unsupported('with formulas or style identifiers that cannot be translated') from None
