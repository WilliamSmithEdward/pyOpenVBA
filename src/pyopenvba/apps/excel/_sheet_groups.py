"""Groups of sheets: Sheets(Array(...)) and the window's selected sheets, as live Excel keeps them.

Measured in live Excel (scripts/measure_sheet_groups.py, tests/fixtures/sheet_groups.json, and
tests/fixtures/sheet_group_moves.json):

- Sheets(Array(...)) and Worksheets(Array(...)) are a Sheets object of the sheets the array names, by name or by
  place, in the array's order, a sheet named twice standing twice. A name no sheet has, or a place past the last
  sheet, is error 9, and an empty array error 13. Its Item takes a place or a name.
- Select makes the array's first sheet active and the group the window's selected sheets, which SelectedSheets
  lists in tab order; with Replace:=False the group, or one sheet, joins the selection and the active sheet stays.
  A hidden sheet in the group is error 1004. Activating a sheet of the group keeps the group; any other sheet
  stands alone, and so does one that Select picks. While sheets are grouped, Worksheets.Add is error 1004.
- Copy with no place makes a new workbook of the group's sheets in tab order, their references to one another kept
  as they are; with Before or After the copies go in tab order at that place, named as one sheet's copy is, their
  references to one another the copies'. The first copy is active then.
- Move with no place makes a new workbook of them. With Before or After they go together, in tab order, before
  the first sheet outside the group that stands at the place the anchor gives or past it -- one place further
  when the array's first sheet stands past that place and some of the group before it, a rule 608 moves bear
  out. The group's first sheet in tab order is active then.
- Delete takes them all, error 1004 where no sheet, or no visible sheet, would be left; the formulas and names
  that read them read #REF! in their place.
- A workbook's whole Sheets or Worksheets collection is a group of every sheet in tab order: Worksheets.Copy
  makes a new workbook of them all, Worksheets.Select selects them all with the first active, and setting
  Visible selects them first, so it is error 1004 while any is hidden.
- Setting Visible selects the group first, then hides or shows every sheet of it; hiding every sheet there is is
  error 1004, and reading Visible is error 1004 too.
- When the active sheet goes, deleted or hidden, the next visible sheet after it is active, or else the last
  visible one before it.
"""

from __future__ import annotations

import re

from collections.abc import Sequence
from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _events
from pyopenvba.apps.excel._model import ExcelObject, Worksheet
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba._a1 import quote_sheet
from pyopenvba.formula._parse import FormulaError, split_sheet, tokenize
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import (ERR_SUBSCRIPT_OUT_OF_RANGE, EMPTY, MISSING, VBAArray, VBAInt, error,
                                           to_bool, to_integer)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_sheets import ChartSheet
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import NameEntry, Workbook

VISIBLE = -1


def group(book: Workbook, array: VBAArray) -> SheetGroup:
    """The sheets an array names, by name or by place, in its order."""
    elements = array.elements()
    if not elements:
        raise error(13, "Type mismatch")
    sheets: list[Worksheet] = []
    for element in elements:
        if isinstance(element, str):
            sheets.append(book.sheet_found(element))
            continue
        place = int(to_integer(element, "Long"))
        if not 1 <= place <= len(book.sheets_):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        sheets.append(book.sheets_[place - 1])
    return SheetGroup(book, sheets)


def in_tab_order(book: Workbook, sheets: list[Worksheet]) -> list[Worksheet]:
    """``sheets`` once each, in the workbook's tab order."""
    return [sheet for sheet in book.sheets_ if any(one is sheet for one in sheets)]


class SheetGroup(VBACollection, ExcelObject):
    """Some of a workbook's sheets, in the order they were named."""

    vba_type_name = "Sheets"

    def __init__(self, book: Workbook, sheets: list[Worksheet]) -> None:
        self.book = book
        self.sheets = sheets

    def vba_items(self) -> list[object]:
        return list(self.sheets)

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            from pyopenvba.apps.excel._names import same_name

            found = next((sheet for sheet in self.sheets if same_name(sheet.name, index)), None)
            if found is None:
                raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
            return found
        return super().vba_lookup(index, items)

    @method
    def Select(self, Replace: object = MISSING) -> object:
        select(self.book, self.sheets, Replace is MISSING or to_bool(Replace))
        return EMPTY

    @method
    def Copy(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import COPYING
        from pyopenvba.apps.excel._sheet_copy import copy_sheets

        self.sheets[0].guard_structure(Before, After, COPYING, "Copying")
        return copy_sheets(in_tab_order(self.book, self.sheets), Before, After)

    @method
    def Move(self, Before: object = MISSING, After: object = MISSING) -> object:
        from pyopenvba.apps.excel._protection import MOVING
        from pyopenvba.apps.excel._sheet_move import move_sheets

        self.sheets[0].guard_structure(Before, After, MOVING, "Moving")
        return move_sheets(self.sheets, Before, After)

    @method
    def Delete(self) -> object:
        delete(self.book, self.sheets, "Delete method of Sheets class failed")
        return EMPTY

    @member
    def Visible(self) -> object:
        raise error(1004, "Unable to get the Visible property of the Sheets class")

    @setter("Visible")
    def _set_visible(self, value: object) -> None:
        # Excel selects the group before it hides a sheet of it (tests/fixtures/sheet_groups.json, hidden_all).
        select(self.book, self.sheets, True)
        set_visible(self.book, self.sheets, value, "Unable to set the Visible property of the Sheets class")

    @method
    def FillAcrossSheets(self, Range: object = MISSING, Type: object = MISSING) -> object:
        fill_across(self.sheets, Range, Type)
        return EMPTY

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def Creator(self) -> object:
        return VBAInt(1480803660, "Long")

    @member
    def Parent(self) -> object:
        return self.book


#: XlFillWith: xlFillWithAll, xlFillWithContents and xlFillWithFormats.
FILL_ALL, FILL_CONTENTS, FILL_FORMATS = -4104, 2, -4122


def fill_across(sheets: list[Worksheet], source: object, kind: object) -> None:
    """Sheets.FillAcrossSheets: the range's cells, their values and formulas, their formats or both, written to the
    same cells of every other sheet of the group, a formula as it is spelled and a blank cell clearing its place;
    one Change for each sheet filled, the last first (tests/fixtures/fill_across_sheets.json). The range has to be
    one area on a sheet of the group, and the group more than that sheet, else error 1004."""
    from pyopenvba.apps.excel import _events
    from pyopenvba.apps.excel._model import Range

    failed = "FillAcrossSheets method of Sheets class failed"
    if source is MISSING:
        raise error(449)
    if not isinstance(source, Range):
        raise error(1004, failed)
    if kind is MISSING:
        chosen = FILL_ALL
    else:
        try:
            chosen = int(to_integer(kind, "Long"))
        except VBARuntimeError:
            raise error(1004, failed) from None
    if chosen not in (FILL_ALL, FILL_CONTENTS, FILL_FORMATS):
        raise error(1004, failed)
    if len(source.areas) != 1:
        raise error(1004, "This can't be done on a multiple range selection.\n\nSelect a single range and try again.")
    origin = source.sheet
    targets = [sheet for sheet in sheets if sheet is not origin]
    if not targets or not any(sheet is origin for sheet in sheets):
        raise error(1004, failed)
    area = source.first
    for sheet in (origin, *targets):
        _refuse_unmodelled(sheet, area)
    contents, formats = chosen != FILL_FORMATS, chosen != FILL_CONTENTS
    for target in reversed(targets):
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                held = origin.cells_.get((row, column))
                cell = target.cell(row, column, create=True)
                assert cell is not None
                if contents:
                    cell.formula = held.formula if held is not None else ""
                    cell.value = EMPTY if held is None or held.formula else held.value
                    cell.stale = bool(cell.formula)
                    cell.shared, cell.dynamic, cell.spilled_from, cell.runs = None, False, None, None
                if formats:
                    target.restyle(row, column, origin.style_at(row, column))
                target.settle(row, column)
                target.cell_changed(row, column)
        _events.after_edit(Range(target, [area]))


def _refuse_unmodelled(sheet: Worksheet, area: Area) -> None:
    """Report what a fill across sheets would carry that the model does not: merged cells, array formulas, spills,
    rich text, notes, hyperlinks, conditional formats and validation."""
    def meets(other: Area) -> bool:
        return other.top <= area.bottom and area.top <= other.bottom and other.left <= area.right \
            and area.left <= other.right

    cells = [cell for (row, column), cell in sheet.cells_.items() if area.contains(row, column)]
    if (any(meets(one) for one in sheet.merged_areas) or any(meets(one) for one in sheet.array_formulas.values())
            or any(area.contains(*anchor) for anchor in sheet.spills)
            or any(cell.runs or cell.dynamic or cell.spilled_from for cell in cells)
            or any(area.contains(*position) for position in sheet.notes)
            or any(meets(link.area) for link in sheet.hyperlinks)
            or any(meets(one) for rule in sheet.conditional_formats for one in rule.areas)
            or any(meets(one) for rule in sheet.validations for one in rule.areas)):
        raise VBAUnsupportedError("FillAcrossSheets over merged cells, array formulas, spills, rich text, notes, "
                                  "hyperlinks, conditional formats or validation is not implemented")


def select(book: Workbook, sheets: list[Worksheet], replace: bool) -> None:
    """Select ``sheets``: in their stead with the first of them active, or beside what is selected."""
    if any(sheet.visible != VISIBLE for sheet in sheets):
        raise error(1004, "Select method of Sheets class failed")
    if replace:
        sheets[0].Activate()
        book.selected_sheets = in_tab_order(book, sheets)
        return
    held = book.selected()
    book.selected_sheets = in_tab_order(book, held + sheets)


def survivor(book: Workbook, going: list[Worksheet]) -> Worksheet | None:
    """The sheet that is active once ``going`` is deleted or hidden: the active sheet where it stays, else the next
    visible sheet after it, or the last visible one before it."""
    active = book.active_sheet
    if active is None or not any(one is active for one in going):
        return active
    place = book.sheets_.index(active)
    staying = [sheet for sheet in book.sheets_
               if sheet.visible == VISIBLE and not any(one is sheet for one in going)]
    later = [sheet for sheet in staying if book.sheets_.index(sheet) > place]
    return later[0] if later else (staying[-1] if staying else None)


def set_visible(book: Workbook, tabs: Sequence[Worksheet | ChartSheet], value: object, refusal: str) -> None:
    """Hide or show ``sheets``; hiding every visible sheet is error 1004, and hiding the active one makes another
    active. 1 shows them as xlSheetVisible does, and what Excel does not name is 1004 (tests/fixtures/settable_audit.json)."""
    from pyopenvba.apps.excel._protection import HIDING, refuse_structure

    refuse_structure(book, HIDING)
    state = VISIBLE if value is True else (0 if value is False else int(to_integer(value, "Long")))
    if state == 1:
        state = VISIBLE
    if state not in (VISIBLE, 0, 2):
        raise error(1004, refusal)
    if state == VISIBLE:
        for tab in tabs:
            tab.visible = state
        return
    if book.chart_sheets:
        from pyopenvba.apps.excel._chart_sheets import hide

        hide(book, tabs, state, refusal)
        return
    # Without a chart sheet in the workbook, every tab is a worksheet.
    sheets = [tab for tab in tabs if isinstance(tab, Worksheet)]
    after = survivor(book, sheets)
    if after is None or not any(sheet.visible == VISIBLE and not any(one is sheet for one in sheets)
                                for sheet in book.sheets_):
        raise error(1004, refusal)
    for sheet in sheets:
        sheet.visible = state
    if after is not book.active_sheet:
        book.activate_sheet(after)


def _reads(text: str, names: set[str]) -> bool:
    """Whether a formula, or a reference, reads a sheet ``names`` holds by its lower-case name."""
    try:
        tokens = tokenize(text if text.startswith("=") else "=" + text)
    except FormulaError:
        return True
    return any(token.kind in ("ref", "name") and split_sheet(token.text)[0].casefold() in names for token in tokens)


def _refused_carriers(book: Workbook, names: set[str], staying: list[Worksheet]) -> None:
    """Report a delete whose sheets something the model does not rewrite reads: a conditional format, a
    validation, a link, a control or a chart on a sheet that stays."""
    from pyopenvba.apps.excel._xstring import decode

    def refuse(what: str) -> None:
        raise VBAUnsupportedError(f"Deleting a worksheet {what} reads is not implemented")

    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    for sheet in staying:
        for rule in sheet.conditional_formats:
            if any(_reads(decode(node.text or ""), names) for node in rule.root.iter(namespace + "formula")):
                refuse("that a conditional format")
        for validation in sheet.validations:
            if any(text and _reads(text, names) for text in (validation.formula1, validation.formula2)):
                refuse("that a validation")
        for link in sheet.hyperlinks:
            if link.sub and _reads(link.sub, names):
                refuse("that a hyperlink")
        for shape in sheet.shapes_:
            control = shape.control
            if control is not None and any(text and _reads(text, names)
                                           for text in (control.linked_cell, control.list_range)):
                refuse("that a control")
            chart = sheet.chart_states.get(shape.shape_id)
            if chart is not None:
                if any(text and _reads(text, names) for series in chart.full_series or chart.series
                       for text in (series.name_reference, series.category_reference, series.value_reference)):
                    refuse("that a chart")
            elif shape.kind == "chart" and shape.source:
                refuse("beside a chart the model has not read, which may read it,")
            else:
                # A picture or a shape reads no sheet but through text linked to a cell.
                linked = re.search(r'\btextlink="([^"]+)"', shape.source)
                if linked is not None and _reads(linked.group(1), names):
                    refuse("that a shape's text")
    # A pivot cache keeps the source its file or Create gave it, as before; Excel lets the sheet go, the cache
    # holding its records (tests/fixtures/pivot_column_fit/, whose macro deletes a table's source sheet).


def gone_references(text: str, names: set[str], order: list[str], tables: frozenset[str] = frozenset()) -> str:
    """A formula with each reference to a sheet ``names`` holds read as #REF! in the sheet's place: =Two!A1+1 is
    =#REF!A1+1 once Two is deleted (tests/fixtures/sheet_groups.json, delete_read and delete_name_read). A span of
    sheets, First:Last!A1, shrinks to the sheets of ``order``, the tabs before the deletion, left at its ends, and a
    reference to one of the ``tables`` that went with them is #REF! whole (tests/fixtures/sheet_delete_refs.json)."""
    for token in reversed(tokenize(text)):
        if token.kind == "structured" and token.text.partition("[")[0].casefold() in tables:
            text = text[:token.at] + "#REF!" + text[token.at + len(token.text):]
            continue
        if token.kind not in ("ref", "name"):
            continue
        sheet, body = split_sheet(token.text)
        kept = _span_left(sheet, names, order) if ":" in sheet else None
        if kept is not None:
            text = text[:token.at] + kept + body + text[token.at + len(token.text):]
        elif sheet and sheet.casefold() in names:
            # A name of the sheet stays a name, of no sheet now, which a file spells [0]!Name; a cell is #REF!.
            kept_text = "[0]!" if token.kind == "name" else "#REF!"
            text = text[:token.at] + kept_text + body + text[token.at + len(token.text):]
    return text


def _keep_names_still_read(book: Workbook, leaving: list[NameEntry], names: set[str]) -> None:
    """A name of a deleted sheet that what is left reads by the sheet's name, =Gone!Mine, stays as the workbook's, unless
    the workbook has one of that name, or a formula left reads the workbook's by the bare name, =Mine, which then has
    none: Excel keeps it reading #REF!, or a constant its value. A formula on the deleted sheet does not count, and a name
    going with it does (tests/fixtures/deleted_local_names.json)."""
    from pyopenvba._a1 import split_sheet as split
    from pyopenvba.apps.excel._tables import rewrite_formulas

    qualified: set[str] = set()
    bare: set[str] = set()

    def noted(text: str) -> str:
        try:
            tokens = tokenize(text)
        except FormulaError:
            return text
        for token in tokens:
            if token.kind == "name":
                sheet, name = split_sheet(token.text)
                if not sheet:
                    bare.add(name.casefold())
                elif sheet.casefold() in names:
                    qualified.add(name.casefold())
        return text

    rewrite_formulas(book, noted)
    for entry in leaving:
        noted(entry.refers_to)
    held = {split(entry.name)[1].casefold() for entry in book.names_.entries if not split(entry.name)[0]}
    for entry in leaving:
        name = split(entry.name)[1]
        if name.casefold() in qualified and name.casefold() not in bare and name.casefold() not in held:
            entry.name = name
            entry.sheet = None
            entry.attributes = re.sub(r'\slocalSheetId="[^"]*"', "", entry.attributes)
            book.names_.entries.append(entry)
            held.add(name.casefold())


def _span_left(span: str, names: set[str], order: list[str]) -> str | None:
    """What stands before a reference's ! for a span of sheets with ``names`` gone, or None if the span stays as it is."""
    first, _, last = span.partition(":")
    folded = [name.casefold() for name in order]
    if first.casefold() not in folded or last.casefold() not in folded:
        return None
    if first.casefold() not in names and last.casefold() not in names:
        return None
    ends = sorted((folded.index(first.casefold()), folded.index(last.casefold())))
    left = [name for name in order[ends[0]:ends[1] + 1] if name.casefold() not in names]
    if not left:
        return "#REF!"
    if len(left) == 1:
        return quote_sheet(left[0]) + "!"
    if quote_sheet(left[0]) == left[0] and quote_sheet(left[-1]) == left[-1]:
        return f"{left[0]}:{left[-1]}!"
    return quote_sheet(f"{left[0]}:{left[-1]}") + "!"


def delete(book: Workbook, sheets: list[Worksheet], refusal: str) -> None:
    """Delete ``sheets``, each formula and name that reads one of them reading #REF! instead."""
    from pyopenvba._a1 import split_sheet as split
    from pyopenvba.apps.excel._documents import remove_document
    from pyopenvba.apps.excel._names import changed
    from pyopenvba.apps.excel._protection import DELETING, refuse_structure
    from pyopenvba.apps.excel._tables import rewrite_formulas

    refuse_structure(book, DELETING)
    going = in_tab_order(book, sheets)
    staying = [sheet for sheet in book.sheets_ if not any(one is sheet for one in going)]
    if not staying:
        raise error(1004, refusal)
    if not any(sheet.visible == VISIBLE for sheet in staying) and not any(chart.visible == VISIBLE for chart in book.chart_sheets):
        # tests/fixtures/sheet_groups.json, delete_last_visible.
        raise error(1004, "A workbook must contain at least one visible worksheet.")
    names = {sheet.name.casefold() for sheet in going}
    order = [sheet.name for sheet in book.sheets_]
    tables = frozenset(table.name.casefold() for sheet in going for table in sheet.tables)
    _refused_carriers(book, names, staying)
    after = survivor(book, going)
    was_active = book.active_sheet
    next_tab = None
    if book.chart_sheets:
        from pyopenvba.apps.excel._chart_sheets import next_visible

        next_tab = next_visible(book, list(going))
    for sheet in reversed(going):
        _events.before_delete(sheet)
    _events.sheet_left(sheets[0])
    # The tabs after a deleted sheet move up, and the file's active tab and first tab with them, which left as they were
    # open the file on another sheet or, past the last, have Excel refuse it (tests/fixtures/deleted_local_names.json).
    before_first = sum(1 for sheet in going if book.tabs().index(sheet) < book.view.first_sheet)
    book.view.first_sheet = max(0, min(book.view.first_sheet - before_first, len(book.tabs()) - len(going) - 1))
    book.view.changed = True
    for sheet in going:
        remove_document(sheet)
        book.sheets_.remove(sheet)
        if book.chart_sheets:
            book.file_tabs.remove(sheet)
    # Local names belong to the removed sheets, independently of RefersTo; retaining them would make the writer
    # silently globalize their scope.
    retained = [entry for entry in book.names_.entries if split(entry.name)[0].casefold() not in names]
    if len(retained) != len(book.names_.entries):
        leaving = [entry for entry in book.names_.entries if split(entry.name)[0].casefold() in names]
        book.names_.entries[:] = retained
        _keep_names_still_read(book, leaving, names)
        changed(book)
    rewrite_formulas(book, lambda text: gone_references(text, names, order, tables))
    if after is not None and any(one is after for one in book.sheets_):
        book.active_sheet_index = book.sheets_.index(after)
    if next_tab is not None and not isinstance(next_tab, Worksheet):
        book.active_chart_sheet = next_tab
    book.selected_sheets = []
    book.saved = False
    for sheet in going:
        # A variable holding a deleted sheet, or a range of it, fails as Excel's does, and a new sheet of the old
        # name never takes its place (tests/fixtures/deleted_sheet.json).
        sheet.deleted = True
        sheet.invalidated = True
    if book.active_sheet is not None and book.active_sheet is not was_active:
        _events.sheet_entered(book.active_sheet)
