"""Worksheet settings macros set beside Protect or on dashboards: ScrollArea, EnableAutoFilter, EnableOutlining,
EnablePivotTable, EnableCalculation, TransitionExpEval, TransitionFormEntry and DisplayRightToLeft.

Measured in live Excel (scripts/measure_worksheet_settings.py, tests/fixtures/worksheet_settings.json):

- The switches take any value VBA turns into a Boolean, Empty as False;
  Null and text that is no Boolean leave one as it was, with no error, as
  Application's switches do.
- ScrollArea takes text Range takes on the sheet, one area of it, spaces
  round it dropped, a name too, and reads it back as the area's absolute
  address. Anything else -- two areas, another sheet's cells, text that is
  no reference, Null, a number -- changes nothing, with no error; "" and
  Empty clear it. Set while the active cell is outside the area, it
  selects the area's first cell, and Goto outside it does too, while a
  macro's Select and Activate still go there.
- EnableCalculation False keeps every formula on the sheet as it is,
  through Calculate, Application.Calculate and CalculateFull alike, and
  another sheet reading one reads the value it keeps; a formula written
  there reads 0. True again works the sheet out under automatic
  calculation, and leaves it for the next calculation under manual.
- A copy of the sheet keeps EnableAutoFilter, EnableOutlining,
  EnablePivotTable, the Transition settings and DisplayRightToLeft, and
  not ScrollArea or EnableCalculation.
- A save writes TransitionExpEval and TransitionFormEntry as sheetPr's
  transitionEvaluation and transitionEntry, and DisplayRightToLeft as the
  sheetView's rightToLeft (see _windows); the rest are not saved.

The model draws nothing, so EnableAutoFilter, EnableOutlining and EnablePivotTable read back what a macro set and change
nothing; TransitionExpEval's 1-2-3 rules are not modelled, so formulas on such a sheet are worked out by Excel's own.
EnableFormatConditionsCalculation is left out: Excel's reading of it lags behind a setting, on a new workbook True until
a sheet is added or copied, by no rule found.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel._application_settings import flag
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import MemberSpec
from pyopenvba.interpreter._values import EMPTY

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
#: The switches a sheet keeps for the session, each with its value on a new sheet.
SWITCHES = {'EnableAutoFilter': False, 'EnableOutlining': False, 'EnablePivotTable': False}
#: What a copy of the sheet keeps of them.
COPIED = ('EnableAutoFilter', 'EnableOutlining', 'EnablePivotTable')
#: The Transition settings by sheetPr's attribute for each.
TRANSITIONS = {'TransitionExpEval': 'transitionEvaluation', 'TransitionFormEntry': 'transitionEntry'}
#: sheetPr's attributes in the schema's order.
_SHEET_PR_ORDER = ('syncHorizontal', 'syncVertical', 'syncRef', 'transitionEvaluation', 'transitionEntry', 'published',
                   'codeName', 'filterMode', 'enableFormatConditionsCalculation')
_SHEET_PR = re.compile(r'<(?:\w+:)?sheetPr\b[^>]*?/?>')


def transitions(sheet: Worksheet) -> dict[str, bool]:
    """The sheet's Transition settings, read from its part the first time."""
    if sheet.transitions is None:
        found: dict[str, bool] = {}
        package = sheet.book.package
        if package is not None and sheet.part_name and package.has(sheet.part_name):
            part = opc_part(package.read(sheet.part_name), NS, 'worksheet')
            empty: dict[str, str] = {}
            attrs = next((attrs for name, attrs, _start, _end in part.children if name == NS + '}sheetPr'), empty)
            for member, attribute in TRANSITIONS.items():
                found[member] = attrs.get(attribute) in ('1', 'true')
        sheet.transitions = found
    return sheet.transitions


def with_transitions(sheet: Worksheet, xml: str) -> str:
    """The sheet's XML with its Transition settings on sheetPr, in the schema's order, a sheetPr made for them."""
    from pyopenvba.apps.excel._windows import with_attribute

    found = transitions(sheet)
    tag = _SHEET_PR.search(xml)
    if tag is None:
        if not any(found.values()):
            return xml
        root = opc_part(xml.encode('utf-8'), NS, 'worksheet')
        prefix = root.qname.rpartition(':')[0]
        opening = f'<{prefix + ":" if prefix else ""}sheetPr/>'
        xml = root.insert(root.opening_end, opening.encode('utf-8')).decode('utf-8')
        tag = _SHEET_PR.search(xml)
        assert tag is not None
    opening = tag.group(0)
    closing = '/>' if opening.endswith('/>') else '>'
    body = opening[:-len(closing)].rstrip()
    for member, attribute in TRANSITIONS.items():
        body = with_attribute(body, attribute, '1' if found.get(member) else '', _SHEET_PR_ORDER)
    xml = xml[:tag.start()] + body + closing + xml[tag.end():]
    # A sheetPr with nothing left in it goes, as Excel writes none.
    return re.sub(r'<((?:\w+:)?sheetPr)\s*(?:/>|></\1>)', '', xml, count=1)


def scroll_area(sheet: Worksheet, value: object) -> None:
    """Set ScrollArea as Excel sets it (see the module's notes)."""
    if value is EMPTY or value == '':
        sheet.settings['ScrollArea'] = ''
        return
    if not isinstance(value, str):
        return
    try:
        found = sheet.vba_get('Range', [value.strip()])
    except VBARuntimeError:
        return
    from pyopenvba.apps.excel._model import Range

    if not isinstance(found, Range) or found.sheet is not sheet or len(found.areas) != 1:
        return
    sheet.settings['ScrollArea'] = str(found.vba_get('Address'))
    _keep_inside(sheet)


def _area(sheet: Worksheet) -> Range | None:
    text = sheet.settings.get('ScrollArea')
    if not text:
        return None
    from pyopenvba.apps.excel._model import Range

    found = sheet.vba_get('Range', [text])
    return found if isinstance(found, Range) else None


def _select_corner(sheet: Worksheet, area: Range) -> None:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Range

    first = area.first
    Range(sheet, [Area(first.top, first.left, first.top, first.left, sheet.name)]).select_as_excel_does()


def _keep_inside(sheet: Worksheet) -> None:
    """The area's first cell selected, where the active cell of the active sheet is outside the scroll area."""
    area = _area(sheet)
    active = sheet.book.active_sheet is sheet and sheet.book.application.active_book is sheet.book
    if area is None or not active:
        return
    cell = sheet.active_cell_range.first if sheet.active_cell_range is not None else None
    if not area.first.contains(cell.top if cell else 1, cell.left if cell else 1):
        _select_corner(sheet, area)


def after_goto(target: Range) -> None:
    """Goto outside the sheet's scroll area: Excel selects the area's first cell instead (scroll_goto)."""
    area = _area(target.sheet)
    if area is None:
        return
    first, inside = target.first, area.first
    if not (inside.contains(first.top, first.left) and inside.contains(first.bottom, first.right)):
        _select_corner(target.sheet, area)


def enable_calculation(sheet: Worksheet, value: object) -> None:
    """Set EnableCalculation: False keeps the sheet's formulas as they are; True again marks them all to be worked
    out, and works them out under automatic calculation (calc_back_on, manual_back_on)."""
    found = flag(value)
    if found is None or found == sheet.enable_calculation:
        return
    sheet.enable_calculation = found
    if found:
        calculator = sheet.book.calculator
        calculator.thawed(sheet)


def copied(source: Worksheet, copy: Worksheet) -> None:
    """What a copy of a sheet keeps of its settings (see the module's notes)."""
    copy.settings.update({name: value for name, value in source.settings.items() if name in COPIED})
    found = transitions(source)
    if any(found.values()):
        copy.transitions = dict(found)
        copy.transitions_changed = True


def _switch(name: str) -> MemberSpec:
    def get(sheet: Worksheet) -> object:
        return sheet.settings.get(name, SWITCHES[name])

    def put(sheet: Worksheet, value: object) -> None:
        found = flag(value)
        if found is not None:
            sheet.settings[name] = found

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _transition(name: str) -> MemberSpec:
    def get(sheet: Worksheet) -> object:
        return transitions(sheet).get(name, False)

    def put(sheet: Worksheet, value: object) -> None:
        found = flag(value)
        if found is not None:
            transitions(sheet)[name] = found
            sheet.transitions_changed = True
            sheet.book.saved = False

    return MemberSpec(name=name, kind='property', getter=get, setter=put)


def _right_to_left() -> MemberSpec:
    def get(sheet: Worksheet) -> object:
        return sheet.view.right_to_left

    def put(sheet: Worksheet, value: object) -> None:
        found = flag(value)
        if found is None:
            return
        if sheet.view.split_xml:
            raise VBAUnsupportedError("changing the view of a window whose panes the model does not lay out, a split "
                                      "Excel kept in twips of what the window shows, is not implemented")
        if found != sheet.view.right_to_left:
            sheet.view.right_to_left = found
            sheet.view.changed = True

    return MemberSpec(name='DisplayRightToLeft', kind='property', getter=get, setter=put)


def members() -> list[MemberSpec]:
    """Every Worksheet setting this module keeps, as a member."""

    def get_scroll(sheet: Worksheet) -> object:
        return sheet.settings.get('ScrollArea', '')

    def get_calculation(sheet: Worksheet) -> object:
        return sheet.enable_calculation

    return [*(_switch(name) for name in SWITCHES), *(_transition(name) for name in TRANSITIONS), _right_to_left(),
            MemberSpec(name='ScrollArea', kind='property', getter=get_scroll, setter=scroll_area),
            MemberSpec(name='EnableCalculation', kind='property', getter=get_calculation, setter=enable_calculation)]
