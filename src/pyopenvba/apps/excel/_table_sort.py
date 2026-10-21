"""Measured, staged sorting of an ordinary local table's data rows."""
# ruff: noqa: N803 -- Excel exposes these named VBA arguments.
from __future__ import annotations

import re
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba._a1 import parse_area
from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._sort import (
    ASCENDING,
    DESCENDING,
    NO,
    YES,
    Key,
    SortField,
    SortFields,
    SortObject,
    SortRange,
    SortState,
    apply_sort,
    plan_sort_area,
    settings_signature,
)
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._tables import Table

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
MAX_CELLS = 1_000_000
MAX_WORK = 4_000_000
MAX_TEXT = 4_000_000
MAX_FIELDS = 64


def unsupported(detail: str) -> VBAUnsupportedError:
    return VBAUnsupportedError('Table sorting ' + detail + ' is not implemented')


def alive(table: Table) -> None:
    if table not in table.sheet.tables:
        raise unsupported('through a removed table proxy')


class _TableState(SortState):
    def __init__(self, table: Table) -> None:
        super().__init__()
        self.header = YES if table.headers else NO
        self.bound_header = False
        self.applied: XmlElement | None = None
        self.dirty = False
        self.cleared = False


def state_for(table: Table) -> _TableState:
    alive(table)
    if table.sort_state is not None:
        if not isinstance(table.sort_state, _TableState):
            raise unsupported('with an unknown settings owner')
        return table.sort_state
    state = _TableState(table)
    if table.xml:
        if len(table.xml) > MAX_TEXT:
            raise unsupported('beyond the supported metadata budget')
        try:
            root = xml_tree(table.xml)
        except ValueError:
            raise unsupported('with invalid table XML') from None
        groups = root.findall(NS + 'sortState')
        if len(groups) > 1:
            raise unsupported('with repeated sort metadata')
        if groups:
            group = groups[0]
            state.applied = deepcopy(group)
            state.match_case = group.get('caseSensitive', '0') in ('1', 'true')
            state.method = 2 if group.get('sortMethod') == 'stroke' else 1
            for condition in group:
                if condition.tag != NS + 'sortCondition':
                    raise unsupported('with unknown sort conditions')
                try:
                    area = parse_area(condition.get('ref', ''))
                except ValueError:
                    raise unsupported('with an invalid key reference') from None
                if area.sheet:
                    raise unsupported('with a qualified stored key reference')
                if len(state.fields) >= MAX_FIELDS:
                    raise unsupported('beyond the supported key count')
                state.fields.append(_TableField(table, Range(table.sheet, [area]),
                                                DESCENDING if condition.get('descending') in ('1', 'true') else ASCENDING,
                                                0, state))
            state.bound_header = bool(state.fields)
    state.original = settings_signature(state)
    table.sort_state = state
    return state


def configured(table: Table) -> bool:
    state = table.sort_state
    if isinstance(state, _TableState) and (state.fields or state.applied is not None):
        return True
    # A persisted sort is configured even before the first Sort getter.
    if table.xml:
        if len(table.xml) > MAX_TEXT:
            raise unsupported('beyond the supported metadata budget')
        try:
            return any(node.tag.rsplit('}', 1)[-1] == 'sortState' for node in xml_tree(table.xml).iter())
        except ValueError:
            raise unsupported('with invalid table XML') from None
    return False


def guard_edit(table: Table, operation: str) -> None:
    if configured(table):
        raise unsupported(operation + ' a table with configured sort settings')


def _metadata(table: Table) -> None:
    if not table.xml:
        return
    if len(table.xml) > MAX_TEXT:
        raise unsupported('beyond the supported metadata budget')
    try:
        root = xml_tree(table.xml)
    except ValueError:
        raise unsupported('with invalid table XML') from None
    # The existing table metadata writer patches native unprefixed lexical
    # forms. Refuse alternate forms it cannot preserve before publishing rows.
    if '<!--' in table.xml or '<![CDATA[' in table.xml:
        raise unsupported('with unmeasured lexical table metadata')
    for match in re.finditer(r"<([A-Za-z_][\w:.-]*)((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", table.xml):
        tag = match.group(1)
        attrs = (match.group(2) or '').rstrip().removesuffix('/')
        if ':' in tag:
            raise unsupported('with prefixed table metadata')
        position = 0
        while position < len(attrs):
            if not attrs[position].isspace():
                raise unsupported('with unmeasured quoted table metadata')
            while position < len(attrs) and attrs[position].isspace():
                position += 1
            if position == len(attrs):
                break
            while position < len(attrs) and (attrs[position].isalnum() or attrs[position] in '_:.-'):
                position += 1
            if attrs[position:position + 2] != '="':
                raise unsupported('with unmeasured quoted table metadata')
            closing = attrs.find('"', position + 2)
            if closing < 0 or '<' in attrs[position + 2:closing] or '>' in attrs[position + 2:closing]:
                raise unsupported('with unmeasured quoted table metadata')
            position = closing + 1
    if root.tag != NS + 'table':
        raise unsupported('with an unknown table namespace')
    if root.get('tableType', 'worksheet') != 'worksheet' or root.get('connectionId') is not None:
        raise unsupported('external or query tables')
    if any(child.tag not in {NS + name for name in ('autoFilter', 'sortState', 'tableColumns', 'tableStyleInfo')}
           for child in root):
        raise unsupported('with table extensions or unknown children')
    if any(len(root.findall(NS + name)) > 1 for name in ('autoFilter', 'sortState', 'tableColumns', 'tableStyleInfo')):
        raise unsupported('with repeated table metadata')
    known = {NS + name for name in ('table', 'autoFilter', 'sortState', 'sortCondition', 'tableColumns',
                                      'tableColumn', 'tableStyleInfo', 'calculatedColumnFormula', 'totalsRowFormula')}
    if any(node.tag not in known for node in root.iter()):
        raise unsupported('with unknown nested table metadata')
    allowed_attributes: dict[str, set[str]] = {
        'table': {'id', 'name', 'displayName', 'ref', 'headerRowCount', 'totalsRowCount', 'totalsRowShown', 'tableType'},
        'tableColumns': {'count'},
        'tableColumn': {'id', 'name', 'totalsRowLabel', 'totalsRowFunction'},
        'tableStyleInfo': {'name', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'},
        'autoFilter': {'ref'},
        'calculatedColumnFormula': set(), 'totalsRowFormula': set(),
    }
    revisions = {'{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid',
                 '{http://schemas.microsoft.com/office/spreadsheetml/2016/revision3}uid'}
    for node in root.iter():
        local = node.tag.removeprefix(NS)
        if local not in allowed_attributes:
            continue
        if any(key not in allowed_attributes[local] and key not in revisions
               and key != '{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable'
               and key != 'xmlns' and not key.startswith('xmlns:') for key in node.attrib):
            raise unsupported('with unmeasured table attributes')
    if root.get('headerRowCount', '1') not in ('0', '1') or root.get('totalsRowCount', '0') not in ('0', '1'):
        raise unsupported('with unmeasured header or totals geometry')
    for group in root.findall(NS + 'sortState'):
        if (group.text or '').strip():
            raise unsupported('with text in sort metadata')
        try:
            area = parse_area(group.get('ref', ''))
        except ValueError:
            raise unsupported('with an invalid stored sort range') from None
        if area.sheet or group.get('caseSensitive', '0') not in ('0', '1', 'false', 'true') \
                or group.get('sortMethod', 'pinYin') not in ('pinYin', 'stroke'):
            raise unsupported('with unrecognized sort settings')
        allowed = {'ref', 'caseSensitive', 'sortMethod'}
        if any(key not in allowed and key != 'xmlns' and not key.startswith('xmlns:') for key in group.attrib):
            raise unsupported('with custom sort metadata')
        for condition in group:
            if (condition.tag != NS + 'sortCondition' or len(condition)
                    or (condition.text or '').strip() or (condition.tail or '').strip()
                    or any(key not in ('ref', 'descending', 'sortBy') for key in condition.attrib)
                    or condition.get('sortBy', 'value') != 'value'
                    or condition.get('descending', '0') not in ('0', '1', 'false', 'true')):
                raise unsupported('with custom sort conditions')
    for group in root.findall(NS + 'autoFilter'):
        if len(group):
            raise unsupported('tables with active filter metadata')


def changed(table: Table) -> bool:
    state = table.sort_state
    return (isinstance(state, _TableState) and state.applied is not None
            and (state.dirty or settings_signature(state) != state.original))


def _state_root(table: Table, state: _TableState, reference: str) -> XmlElement:
    if len(state.fields) > MAX_FIELDS:
        raise unsupported('beyond the supported key count')
    root = XmlElement(NS + 'sortState', {'xmlns': NS[1:-1], 'ref': reference,
                                        'xmlns:xlrd2': 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2'})
    if state.match_case:
        root.set('caseSensitive', '1')
    if state.method == 2:
        root.set('sortMethod', 'stroke')
    for field in state.fields:
        key = field.key
        if (key.sheet is not table.sheet or len(key.areas) != 1 or key.first.columns != 1
                or not table.area.left <= key.first.left <= table.area.right
                or key.first.bottom < table.area.top or key.first.top > table.area.bottom
                or field.order not in (ASCENDING, DESCENDING) or field.option not in (0, 1)):
            raise unsupported('persisting unmeasured field settings')
        attrs = {'ref': key.first.address(absolute=False)}
        if field.order == DESCENDING:
            attrs['descending'] = '1'
        root.append(XmlElement(NS + 'sortCondition', attrs))
    return root


def write_state(table: Table, xml: str, *, planned: XmlElement | None = None) -> str:
    if planned is None:
        state = table.sort_state
        if not isinstance(state, _TableState) or not changed(table) or state.applied is None:
            return xml
        _metadata(table)
        if state.cleared and state.fields:
            raise unsupported('persisting new fields after Clear without another Apply')
        planned = _state_root(table, state, state.applied.get('ref', '')) if state.fields else None
    part = opc_part(xml.encode('utf8'), NS[1:-1], 'table')
    found = [(start, end) for name, _attrs, start, end in part.children if name == NS[1:] + 'sortState']
    if len(found) > 1:
        raise unsupported('with repeated sort metadata')
    markup = xml_bytes(planned).split(b'?>', 1)[1].lstrip() if planned is not None else b''
    if found:
        start, end = found[0]
        return (part.data[:start] + markup + part.data[end:]).decode('utf8')
    before = {NS[1:] + name for name in ('tableColumns', 'tableStyleInfo', 'extLst')}
    at = next((start for name, _attrs, start, _end in part.children if name in before), part.closing_start)
    return part.insert(at, markup).decode('utf8') if markup else xml


class _TableField(SortField):
    vba_type_name = "SortField"

    def __init__(self, table: Table, key: Range, order: int, option: int, state: _TableState) -> None:
        super().__init__(key, order, option)
        self.table = table
        self.sheet = table.sheet
        self.state = state

    def _check_alive(self) -> None:
        super()._check_alive()
        alive(self.table)
        if self not in self.state.fields:
            raise unsupported('through a removed sort-field proxy')

    @setter('Order')
    def _set_order(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in (ASCENDING, DESCENDING):
            raise error(5)
        self.order = wanted


class _TableFields(SortFields):
    vba_type_name = "SortFields"

    def __init__(self, owner: TableSortObject) -> None:
        super().__init__(owner)
        self.table_owner = owner

    @method
    def Clear(self) -> object:
        state = state_for(self.table_owner.table)
        state.fields.clear()
        state.cleared = state.applied is not None
        return EMPTY

    @method
    def Add(self, Key: object = MISSING, SortOn: object = MISSING, Order: object = MISSING,
            CustomOrder: object = MISSING, DataOption: object = MISSING) -> object:
        if not isinstance(Key, Range):
            raise error(1004)
        if SortOn is not MISSING and int(to_integer(SortOn, 'Long')) != 0:
            raise unsupported('on colors, icons or rich values')
        if CustomOrder is not MISSING:
            raise unsupported('with custom ordering')
        order = ASCENDING if Order is MISSING else int(to_integer(Order, 'Long'))
        option = 0 if DataOption is MISSING else int(to_integer(DataOption, 'Long'))
        if order not in (ASCENDING, DESCENDING):
            raise unsupported('with an unmeasured key order')
        if len(self.owner.state.fields) >= MAX_FIELDS:
            raise unsupported('beyond the supported key count')
        state = state_for(self.table_owner.table)
        field = _TableField(self.table_owner.table, Key, order, option, state)
        state.header = int(self.table_owner.Header())
        state.bound_header = True
        state.fields.append(field)
        return field


class TableSortObject(SortObject):
    vba_type_name = "Sort"

    def __init__(self, table: Table) -> None:
        self.table = table
        self.sheet = table.sheet
        self.state = state_for(table)

    def _check_alive(self) -> None:
        super()._check_alive()
        alive(self.table)

    @member
    def Parent(self) -> object:
        from pyopenvba.apps.excel._tables import view

        return view(self.table)

    @member
    def Rng(self) -> object:
        return SortRange(self, Range(self.sheet, [self.table.area]))

    @member
    def SortFields(self) -> object:
        return _TableFields(self)

    @method
    def SetRange(self, Rng: object = MISSING) -> object:
        raise error(1004)

    @member
    def Header(self) -> object:
        state = state_for(self.table)
        return VBAInt(state.header if state.bound_header else YES if self.table.headers else NO, 'Long')

    @setter('Header')
    def _set_header(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        current = int(self.Header())
        if wanted not in (0, current):
            raise error(5)
        state = state_for(self.table)
        state.header, state.bound_header = current, True

    @setter('Orientation')
    def _set_orientation(self, value: object) -> None:
        if int(to_integer(value, 'Long')) != 1:
            raise error(5)
        self.state.orientation = 1

    @setter('SortMethod')
    def _set_sort_method(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if wanted not in (1, 2):
            raise error(5)
        self.state.method = wanted

    @method
    def Apply(self) -> object:
        from pyopenvba.apps.excel import _events
        from pyopenvba.apps.excel._conditional import meets
        from pyopenvba.apps.excel._conditional_transfer import guard_extensions
        from pyopenvba.apps.excel._hyperlinks import touches
        from pyopenvba.apps.excel._visible import filtering

        table, sheet, state = self.table, self.sheet, state_for(self.table)
        if not state.fields:
            return EMPTY
        _metadata(table)
        data = table.data
        if data is None:
            raise unsupported('tables without data rows')
        if len(table.columns) != data.columns:
            raise unsupported('with an inconsistent column schema')
        if state.method != 1:
            raise unsupported('with an unmeasured stroke ordering')
        if (data.rows * data.columns > MAX_CELLS
                or data.rows * max(1, data.rows.bit_length()) * len(state.fields) > MAX_WORK):
            raise unsupported('beyond the supported transaction work budget')
        if filtering(sheet):
            raise unsupported('on filtered sheets')
        guard_extensions(sheet)
        blocks = [*sheet.merged_areas, *sheet.array_formulas.values(), *(spill.area for spill in sheet.spills.values())]
        blocks.extend(area for rule in [*sheet.conditional_formats, *sheet.validations] for area in rule.areas)
        if (any(meets(data, area) for area in blocks) or any(data.contains(*key) for key in sheet.notes)
                or touches(sheet, [data]) or any(data.contains(*key) and cell.dynamic for key, cell in sheet.cells_.items())):
            raise unsupported('over arrays, rules, annotations or merged cells')
        keys: list[Key] = []
        for field in state.fields:
            selected = field.key
            if selected.sheet is not sheet:
                raise error(1004)
            if len(selected.areas) != 1 or selected.first.columns != 1:
                raise unsupported('with multi-area or multi-column keys')
            if selected.first.bottom < table.area.top or selected.first.top > table.area.bottom:
                raise unsupported('with keys outside the table rows')
            index = selected.first.left - data.left
            if not 0 <= index < data.columns:
                raise error(1004)
            if field.option not in (0, 1):
                raise unsupported('with unmeasured data options')
            keys.append(Key(index, field.order == DESCENDING, field.option == 1))
        remaining = MAX_TEXT

        def charge(value: object) -> None:
            nonlocal remaining
            if isinstance(value, str):
                remaining -= len(value)
                if remaining < 0:
                    raise unsupported('beyond the supported transaction text budget')

        for (row, column), cell in sheet.cells_.items():
            if data.contains(row, column):
                charge(cell.formula)
        try:
            plan = plan_sort_area(sheet, data, keys, header=False, match_case=state.match_case,
                                  across=False, charge_value=charge)
        except FormulaError:
            raise unsupported('with an unrecognized moved formula') from None
        applied = _state_root(table, state, data.address(absolute=False))
        # Validate lexical replacement before publishing either cells or settings.
        from pyopenvba.apps.excel._tables import base_table_xml

        write_state(table, table.xml or base_table_xml(table), planned=applied)
        if plan is not None and (not sheet.book.application.udf_callers
                                 or any(source != target for source, target in plan.destination.items())):
            apply_sort(plan, preserve_cached=not sheet.book.calculator.automatic)
        state.applied, state.dirty, state.cleared = applied, True, False
        table.changed = True
        sheet.touched()
        if not sheet.book.application.udf_callers:
            _events.recalculated(sheet.book)
        return EMPTY
