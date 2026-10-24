"""Read-only admission of known secondary references during worksheet rename."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement, xml_tree
from pyopenvba.apps.excel._conditional_predicates import PREDICATE_TYPES
from pyopenvba.apps.excel._xstring import decode

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._worksheet_rename import Budget

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CONDITIONAL_KINDS = {'cellIs', 'expression', 'top10', 'aboveAverage', 'uniqueValues', 'duplicateValues'} | PREDICATE_TYPES
CONDITIONAL_ATTRIBUTES = {'type', 'priority', 'dxfId', 'stopIfTrue', 'operator', 'text', 'timePeriod',
                          'rank', 'percent', 'bottom', 'aboveAverage', 'equalAverage', 'stdDev'}
VALIDATION_ATTRIBUTES = {'type', 'errorStyle', 'imeMode', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage',
                         'showErrorMessage', 'errorTitle', 'error', 'promptTitle', 'prompt', 'sqref'}


def attributes_known(root: XmlElement, allowed: set[str]) -> bool:
    return all(key in allowed or key == 'xmlns' or key.startswith('xmlns:') for key in root.attrib)


def container(root: XmlElement) -> None:
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    if (root.tag != NS + 'conditionalFormatting' or not attributes_known(root, {'sqref'})
            or any(child.tag != NS + 'cfRule' for child in root)):
        raise unsupported('with conditional reference extensions')


def reference(text: str, old: str, wanted: str, budget: Budget, *, encoded: bool = False) -> None:
    from pyopenvba.apps.excel._worksheet_rename import formula, unsupported

    if not text:
        return
    budget.text(text)
    value = decode(text) if encoded else text
    if formula(value, old, wanted, budget) != value:
        raise unsupported('with affected secondary references')


def conditional(root: XmlElement, old: str, wanted: str, budget: Budget) -> None:
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    for key, value in root.attrib.items():
        budget.text(key)
        budget.text(value)
    if root.tag == NS + 'cfRule' and root.get('type') == 'colorScale':
        colour_scale(root, old, wanted, budget)
        return
    if root.tag == NS + 'cfRule' and root.get('type') in ('dataBar', 'iconSet'):
        drawn(root, old, wanted, budget)
        return
    if root.tag != NS + 'cfRule' or root.get('type') not in CONDITIONAL_KINDS or not attributes_known(root, CONDITIONAL_ATTRIBUTES):
        raise unsupported('with unknown conditional reference metadata')
    for child in root:
        if child.tag != NS + 'formula' or list(child) or not attributes_known(child, set()):
            raise unsupported('with conditional reference extensions')
        reference(child.text or '', old, wanted, budget, encoded=True)


def colour_scale(root: XmlElement, old: str, wanted: str, budget: Budget) -> None:
    """A colour scale's points and colours, of which only a formula point can name a sheet."""
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    scale = root.find(NS + 'colorScale')
    if not attributes_known(root, {'type', 'priority'}) or len(root) != 1 or scale is None or scale.attrib:
        raise unsupported('with unknown conditional reference metadata')
    for child in scale:
        for key, value in child.attrib.items():
            budget.text(key)
            budget.text(value)
        if list(child):
            raise unsupported('with conditional reference extensions')
        if child.tag == NS + 'cfvo' and attributes_known(child, {'type', 'val', 'gte'}):
            _point(child, old, wanted, budget)
        elif child.tag != NS + 'color' or not attributes_known(child, {'rgb', 'theme', 'tint', 'indexed', 'auto'}):
            raise unsupported('with conditional reference extensions')


def _point(point: XmlElement, old: str, wanted: str, budget: Budget) -> None:
    """A point's value, which any but a lowest or highest can hold as a formula."""
    from pyopenvba.apps.excel._conditional_extension import NUMBER

    value = point.get('val', '')
    if point.get('type') not in ('min', 'max') and not NUMBER.fullmatch(value):
        reference(value, old, wanted, budget)


def drawn(root: XmlElement, old: str, wanted: str, budget: Budget) -> None:
    """A data bar's or icon set's cfRule: its formula, its points and colours, and the link to its x14:cfRule,
    whose formulas admit checks through the rule's settings."""
    from pyopenvba.apps.excel._conditional_extension import LINK_URI, X14
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    shape = NS + ('dataBar' if root.get('type') == 'dataBar' else 'iconSet')
    if not attributes_known(root, {'type', 'priority'}):
        raise unsupported('with unknown conditional reference metadata')
    for child in root:
        for key, value in child.attrib.items():
            budget.text(key)
            budget.text(value)
        if child.tag == NS + 'formula' and not list(child):
            reference(child.text or '', old, wanted, budget, encoded=True)
        elif child.tag == shape:
            for part in child:
                if part.tag == NS + 'cfvo' and attributes_known(part, {'type', 'val', 'gte'}) and not list(part):
                    _point(part, old, wanted, budget)
                elif part.tag != NS + 'color' or list(part):
                    raise unsupported('with conditional reference extensions')
        elif not (child.tag == NS + 'extLst' and [(ext.tag, ext.get('uri')) for ext in child] == [(NS + 'ext', LINK_URI)]
                  and all(named.tag == X14 + 'id' for named in child[0])):
            raise unsupported('with conditional reference extensions')


def worksheet(raw: bytes, old: str, wanted: str, budget: Budget) -> None:
    """Inspect full original XML so wrapped/foreign carriers cannot disappear."""
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    root = xml_tree(raw)
    direct = set(root)
    parents = {child: parent for parent in root.iter() for child in parent}
    for node in root.iter():
        local = node.tag.rsplit('}', 1)[-1]
        if local in ('cfRule', 'formula', 'dataValidation', 'formula1', 'formula2'):
            parent = parents.get(node)
            grandparent = parents.get(parent) if parent is not None else None
            if local == 'cfRule':
                owned = parent is not None and parent in direct and parent.tag == NS + 'conditionalFormatting'
            elif local == 'dataValidation':
                owned = parent is not None and parent in direct and parent.tag == NS + 'dataValidations'
            else:
                expected = 'cfRule' if local == 'formula' else 'dataValidation'
                group = 'conditionalFormatting' if local == 'formula' else 'dataValidations'
                owned = (parent is not None and parent.tag == NS + expected and grandparent is not None
                         and grandparent in direct and grandparent.tag == NS + group)
            if node.tag != NS + local or not owned:
                raise unsupported('with unowned secondary formula metadata')
        if local in ('conditionalFormatting', 'conditionalFormattings'):
            if node not in direct or node.tag != NS + 'conditionalFormatting' or not attributes_known(node, {'sqref'}):
                raise unsupported('with wrapped or unknown conditional references')
            for rule in node:
                conditional(rule, old, wanted, budget)
        elif local == 'dataValidations':
            # Extended validation has additional formula carriers and remains deferred.
            if (node not in direct or node.tag != NS + local
                    or not attributes_known(node, {'count', 'disablePrompts', 'xWindow', 'yWindow'})):
                raise unsupported('with extended validation references')
            for rule in node:
                if rule.tag != NS + 'dataValidation' or not attributes_known(rule, VALIDATION_ATTRIBUTES):
                    raise unsupported('with unknown validation references')
                for child in rule:
                    if child.tag not in (NS + 'formula1', NS + 'formula2') or list(child):
                        raise unsupported('with unknown validation formulas')
                    reference(child.text or '', old, wanted, budget)


def admit(book: Workbook, old: str, wanted: str, budget: Budget) -> None:
    """Prove supported carriers unaffected; never replace roots or retained proxies."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._worksheet_rename import unsupported

    containers: set[XmlElement] = set()
    count = sum(len(owner.conditional_formats) + len(owner.validations) + len(owner.shapes_)
                + len(owner.hyperlinks) + len(owner.pivot_tables) for owner in book.sheets_) + len(book.pivot_caches)
    if count > 100_000:
        raise unsupported('beyond the supported secondary-reference count')
    for owner in book.sheets_:
        for rule in owner.conditional_formats:
            if rule.container is not None and rule.container not in containers:
                for key, value in rule.container.attrib.items():
                    budget.text(key)
                    budget.text(value)
                container(rule.container)
                containers.add(rule.container)
            conditional(rule.root, old, wanted, budget)
            if rule.root.get('type') in ('dataBar', 'iconSet'):
                from pyopenvba.apps.excel._conditional_extension import point_formulas

                for text in point_formulas(rule):
                    reference(text, old, wanted, budget)
        for validation in owner.validations:
            if validation.extension:
                raise unsupported('with extended validation references')
            reference(validation.formula1, old, wanted, budget)
            reference(validation.formula2, old, wanted, budget)
        for link in owner.hyperlinks:
            if link.address:
                raise unsupported('with external hyperlinks')
            reference(link.sub, old, wanted, budget)
        for shape in owner.shapes_:
            if shape.kind not in ('shape', 'textBox', 'line', 'formControl') or shape.macro or shape.source or shape.children:
                raise unsupported('with unmeasured drawing references')
            if shape.control is not None:
                if shape.control.macro_text:
                    raise unsupported('with unmeasured control macro references')
                reference(shape.control.linked_cell, old, wanted, budget)
                reference(shape.control.list_range, old, wanted, budget)
    from pyopenvba.apps.excel._pivot_refresh import admit_cache, admit_report, integer

    caches = set(book.pivot_caches)
    roots = [cache.root for cache in book.pivot_caches]
    roots.extend(table.root for owner in book.sheets_ for table in owner.pivot_tables)
    pivot_nodes = 0
    for root in roots:
        for node in root.iter():
            pivot_nodes += 1
            if pivot_nodes > 100_000:
                raise unsupported('beyond the supported Pivot metadata count')
            budget.text(node.tag)
            budget.text(node.text or '')
            budget.text(node.tail or '')
            if (node.text or '').strip() or (node.tail or '').strip():
                raise unsupported('with opaque Pivot metadata text')
            for key, value in node.attrib.items():
                budget.text(key)
                budget.text(value)
    for cache in book.pivot_caches:
        admit_cache(cache)
    for owner in book.sheets_:
        for table in owner.pivot_tables:
            if table.cache not in caches or table.sheet is not owner:
                raise unsupported('with unregistered Pivot report ownership')
            fields = table.cache.root.find(NS + 'cacheFields')
            assert fields is not None
            admit_report(table, len(fields), integer(table.cache.root, 'recordCount'))
    for cache in book.pivot_caches:
        origin = cache.root.find(NS + 'cacheSource/' + NS + 'worksheetSource')
        if origin is not None:
            for key, value in origin.attrib.items():
                budget.text(key)
                budget.text(value)
        if origin is None or origin.get('name') or origin.get('sheet', '').casefold() == old.casefold():
            raise unsupported('with affected or unmeasured Pivot cache references')
        # The explicit source is authoritative; no name-resolution side effects.
        from pyopenvba._a1 import parse_area

        if not any(owner.name.casefold() == origin.get('sheet', '').casefold() for owner in book.sheets_):
            raise unsupported('with an unregistered Pivot source owner')
        parse_area(origin.get('ref', ''))
        source = cache.pending_source
        if isinstance(source, Range):
            if source.sheet.name.casefold() == old.casefold():
                raise unsupported('with affected Pivot source ranges')
        elif source is not None:
            from pyopenvba.apps.excel._pivot_create import a1_text

            reference(a1_text(str(source)), old, wanted, budget)
